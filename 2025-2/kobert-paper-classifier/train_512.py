import json
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
import torch
from sklearn.metrics import f1_score, classification_report
from sklearn.utils.class_weight import compute_class_weight
from transformers import AutoTokenizer, AutoModelForSequenceClassification, get_linear_schedule_with_warmup
from torch.utils.data import Dataset, DataLoader

BASE_DIR = Path("/content/drive/MyDrive/2025/02")
OUT_DIR_512 = BASE_DIR / "kobert_outputs_512"  # 원본 버그 수정: 384와 분리
OUT_DIR_512.mkdir(parents=True, exist_ok=True)

TRAIN_XLSX = BASE_DIR / "경진대회" / "train_dataset.xlsx"
VALID_XLSX = BASE_DIR / "경진대회" / "validation_dataset.xlsx"

device = "cuda" if torch.cuda.is_available() else "cpu"

MODEL_NAME = "skt/kobert-base-v1"
MAX_LEN_512 = 512
BATCH_512 = 8
EPOCHS = 50
LR = 2e-5
PATIENCE = 8


class PaperDS(Dataset):
    def __init__(self, df: pd.DataFrame, tok, max_len: int):
        self.df = df.reset_index(drop=True)
        self.tok = tok
        self.max_len = max_len

    def __len__(self):
        return len(self.df)

    def __getitem__(self, i: int):
        text = self.df.loc[i, "text"]
        y = int(self.df.loc[i, "y"])
        enc = self.tok(text, truncation=True, padding="max_length", max_length=self.max_len)
        item = {k: torch.tensor(v) for k, v in enc.items()}
        item["labels"] = torch.tensor(y)
        return item


def load_and_preprocess(train_xlsx: Path, valid_xlsx: Path):
    train = pd.read_excel(train_xlsx)
    valid = pd.read_excel(valid_xlsx)

    rename = {
        "ID": "id", "Datastamp": "datastamp", "title": "title", "abstract": "abstract",
        "publisher": "publisher", "issn": "issn", "creator": "creator", "label": "label",
        "게재일자": "datastamp", "제목": "title", "요약": "abstract", "학회지": "publisher",
        "저자": "creator", "주제분류": "label",
    }
    train = train.rename(columns=rename)
    valid = valid.rename(columns=rename)

    for c in ["title", "abstract", "publisher", "issn", "creator", "label"]:
        if c in train:
            train[c] = train[c].astype(str).replace({"nan": ""}).str.strip()
        if c in valid:
            valid[c] = valid[c].astype(str).replace({"nan": ""}).str.strip()

    train["text"] = (train.get("title", "") + " [SEP] " + train.get("abstract", "")).str.strip()
    valid["text"] = (valid.get("title", "") + " [SEP] " + valid.get("abstract", "")).str.strip()

    labels: List[str] = sorted(train["label"].dropna().unique().tolist())
    lab2id: Dict[str, int] = {l: i for i, l in enumerate(labels)}
    id2lab: Dict[int, str] = {i: l for l, i in lab2id.items()}

    train["y"] = train["label"].map(lab2id)
    valid["y"] = valid["label"].map(lab2id)

    return train, valid, lab2id, id2lab


def main():
    train, valid, lab2id, id2lab = load_and_preprocess(TRAIN_XLSX, VALID_XLSX)
    (OUT_DIR_512 / "label2id.json").write_text(json.dumps(lab2id, ensure_ascii=False, indent=2), encoding="utf-8")

    tok = AutoTokenizer.from_pretrained(MODEL_NAME)  # 384 모델과 같은 토크나이저 사용
    model = AutoModelForSequenceClassification.from_pretrained(MODEL_NAME, num_labels=len(lab2id)).to(device)

    train_dl = DataLoader(PaperDS(train, tok, MAX_LEN_512), batch_size=BATCH_512, shuffle=True)
    valid_dl = DataLoader(PaperDS(valid, tok, MAX_LEN_512), batch_size=BATCH_512, shuffle=False)

    cw = compute_class_weight(class_weight="balanced", classes=np.arange(len(lab2id)), y=train["y"].values)
    cw = torch.tensor(cw, dtype=torch.float32, device=device)
    criterion = torch.nn.CrossEntropyLoss(weight=cw)

    optim = torch.optim.AdamW(model.parameters(), lr=LR)
    total_steps = EPOCHS * len(train_dl)
    sched = get_linear_schedule_with_warmup(optim, int(0.1 * total_steps), total_steps)

    best_f1, wait = 0.0, 0
    for epoch in range(1, EPOCHS + 1):
        model.train()
        for batch in train_dl:
            batch = {k: v.to(device) for k, v in batch.items()}
            logits = model(input_ids=batch["input_ids"], attention_mask=batch["attention_mask"]).logits
            loss = criterion(logits, batch["labels"])
            optim.zero_grad()
            loss.backward()
            torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
            optim.step()
            sched.step()

        model.eval()
        preds, trues = [], []
        with torch.no_grad():
            for batch in valid_dl:
                batch = {k: v.to(device) for k, v in batch.items()}
                logits = model(input_ids=batch["input_ids"], attention_mask=batch["attention_mask"]).logits
                preds.extend(torch.argmax(logits, dim=-1).cpu().tolist())
                trues.extend(batch["labels"].cpu().tolist())
        f1 = f1_score(trues, preds, average="macro")
        print(f"[512] Epoch {epoch} | Val Macro-F1: {f1:.4f}")

        pd.DataFrame({
            "id": valid["id"] if "id" in valid.columns else range(len(valid)),
            "gold": [id2lab[i] for i in trues],
            "pred": [id2lab[i] for i in preds],
        }).to_csv(OUT_DIR_512 / f"preds_epoch{epoch}.csv", index=False, encoding="utf-8")

        if f1 > best_f1:
            best_f1, wait = f1, 0
            model.save_pretrained(OUT_DIR_512 / "best_model")
            tok.save_pretrained(OUT_DIR_512 / "best_model")
        else:
            wait += 1
            if wait > PATIENCE:
                break

    print("[512] Best Val Macro-F1:", round(best_f1, 4))
    print(classification_report(trues, preds, target_names=[id2lab[i] for i in range(len(id2lab))]))


if __name__ == "__main__":
    main()
