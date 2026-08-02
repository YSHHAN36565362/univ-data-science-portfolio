from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
import torch
from sklearn.metrics import f1_score, classification_report
from torch.nn import functional as F
from transformers import AutoTokenizer, AutoModelForSequenceClassification
from torch.utils.data import Dataset, DataLoader

BASE_DIR = Path("/content/drive/MyDrive/2025/02")
OUT_DIR_384 = BASE_DIR / "kobert_outputs_384"
OUT_DIR_512 = BASE_DIR / "kobert_outputs_512"
VALID_XLSX = BASE_DIR / "경진대회" / "validation_dataset.xlsx"

device = "cuda" if torch.cuda.is_available() else "cpu"


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


def load_valid_with_labels(valid_xlsx: Path, lab2id: Dict[str, int]):
    valid = pd.read_excel(valid_xlsx)
    rename = {
        "제목": "title", "요약": "abstract", "주제분류": "label",
        "title": "title", "abstract": "abstract", "label": "label",
    }
    valid = valid.rename(columns=rename)
    for c in ["title", "abstract", "label"]:
        if c in valid:
            valid[c] = valid[c].astype(str).replace({"nan": ""}).str.strip()
    valid["text"] = (valid.get("title", "") + " [SEP] " + valid.get("abstract", "")).str.strip()
    valid["y"] = valid["label"].map(lab2id)
    return valid


def probs_from(model_dir: str, df: pd.DataFrame, max_len: int, bs: int = 32):
    tok = AutoTokenizer.from_pretrained(model_dir)
    ds = PaperDS(df, tok, max_len)
    dl = DataLoader(ds, batch_size=bs, shuffle=False)
    model = AutoModelForSequenceClassification.from_pretrained(model_dir).to(device).eval()

    probs: List[np.ndarray] = []
    ys: List[int] = []
    with torch.no_grad():
        for batch in dl:
            b = {k: v.to(device) for k, v in batch.items()}
            logits = model(input_ids=b["input_ids"], attention_mask=b["attention_mask"]).logits
            probs.append(F.softmax(logits, dim=-1).cpu().numpy())
            ys.extend(b["labels"].cpu().tolist())
    return np.vstack(probs), np.array(ys)


def main():
    import json

    lab2id = json.loads((OUT_DIR_384 / "label2id.json").read_text(encoding="utf-8"))
    id2lab = {v: k for k, v in lab2id.items()}

    valid = load_valid_with_labels(VALID_XLSX, lab2id)

    p384, yv = probs_from(str(OUT_DIR_384 / "best_model"), valid, max_len=384, bs=32)
    p512, _ = probs_from(str(OUT_DIR_512 / "best_model"), valid, max_len=512, bs=32)

    p_ens = (p384 + p512) / 2.0
    pred_ens = p_ens.argmax(axis=1)

    macro_f1 = f1_score(yv, pred_ens, average="macro")
    print("Ensemble Macro-F1:", macro_f1)
    print(classification_report(yv, pred_ens, target_names=[id2lab[i] for i in range(len(id2lab))]))

    pd.DataFrame({
        "gold": [id2lab[i] for i in yv],
        "pred": [id2lab[i] for i in pred_ens],
    }).to_csv(BASE_DIR / "ensemble_preds.csv", index=False, encoding="utf-8")


if __name__ == "__main__":
    main()
