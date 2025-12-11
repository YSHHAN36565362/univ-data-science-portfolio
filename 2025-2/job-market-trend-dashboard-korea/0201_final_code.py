#설치 안했다면 라이브러리 설치
#pip install openpyxl

import pandas as pd
import matplotlib.pyplot as plt 
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg 
from matplotlib.figure import Figure 
from matplotlib.patches import Polygon 
from matplotlib.collections import PatchCollection 
from tkinter import * 
from tkinter import ttk 
import numpy as np 
import os
import platform
import json 
import matplotlib.dates as mdates


# 1. 설정 및 데이터 로드


# 한글 폰트 설정
font_path = ''
if platform.system() == 'Windows':
    font_name = 'Malgun Gothic' 
    font_path = 'c:/Windows/Fonts/malgun.ttf'
elif platform.system() == 'Darwin': 
    font_name = 'AppleGothic'
    font_path = '/System/Library/Fonts/AppleGothic.ttf'

plt.rc('font', family=font_name)
plt.rc('axes', unicode_minus=False)

# 전국 17개 시도 중심 좌표
PROVINCE_COORDS = {
    '서울': (37.5665, 126.9780), '부산': (35.1796, 129.0756), '대구': (35.8714, 128.6014),
    '인천': (37.4563, 126.7052), '광주': (35.1601, 126.8517), '대전': (36.3504, 127.3845),
    '울산': (35.5384, 129.3114), '세종': (36.4800, 127.2890), '경기': (37.4138, 127.5183),
    '강원': (37.8228, 128.1555), '충북': (36.6350, 127.4914), '충남': (36.6588, 126.6728),
    '전북': (35.7175, 127.1530), '전남': (34.8679, 126.9910), '경북': (36.5760, 128.5056),
    '경남': (35.4606, 128.2132), '제주': (33.4996, 126.5312)
}

# 줌 레벨
ZOOM_LEVELS = {
    '서울': 0.15, '부산': 0.15, '대구': 0.2, '인천': 0.2, '광주': 0.15, '대전': 0.15, '울산': 0.15, '세종': 0.1,
    '경기': 0.8, '강원': 1.5, '충북': 0.8, '충남': 0.8, '전북': 0.8, '전남': 1.0, '경북': 1.2, '경남': 1.0, '제주': 0.4
}

# 지도 이름 매핑
GEO_NAME_MAP = {
    'Seoul': '서울', 'Busan': '부산', 'Daegu': '대구', 'Incheon': '인천',
    'Gwangju': '광주', 'Daejeon': '대전', 'Ulsan': '울산', 'Sejong': '세종',
    'Gyeonggi': '경기', 'Gangwon': '강원', 'Chungbuk': '충북', 'Chungnam': '충남',
    'Jeonbuk': '전북', 'Jeonnam': '전남', 'Gyeongbuk': '경북', 'Gyeongnam': '경남',
    'Jeju': '제주'
}

# 인구순 정렬
POPULATION_ORDER = [
    '서울', '경기', '부산', '인천', '경남', '대구', '경북', '전남', '충남', '전북', 
    '충북', '강원', '대전', '광주', '울산', '제주', '세종'
]

# 학력 정렬 순서
EDU_ORDER = ['초졸이하', '중졸', '고졸', '전문대졸', '대졸', '석박사', '학력무관']

#---------------------------------------------
#경로설정 ★★
#---------------------------------------------
# 데이터 파일 경로 설정
BASE_PATH_TREND = r"C:\Users\UserPc\Desktop\최종데이터셋\최종데이터셋\최종데이터셋"  #2022~2025 데이터셋 있는 폴더 경로 복붙
# json 폴더 경로
JSON_DIR = r"C:\Users\UserPc\Desktop\최종데이터셋\json" # 압축 풀자마자 있는 json폴더 경로
# 실제 geojson 파일 경로 (파일명 맞게 수정!)
JSON_PATH = os.path.join(JSON_DIR, "skorea-provinces-geo.json") 
#  통합 데이터가 있는 폴더 경로
BASE_PATH_INTEGRATED = r"C:\Users\UserPc\Desktop\최종데이터셋" # 압축헤제한 폴더 경로

class DataLoader:
    def __init__(self):
        self.df_open, self.df_seek = self.load_trend_data()
        # [NEW] 통합 데이터 로드 (전국 직종별 통계)
        self.df_integrated = self.load_integrated_data()
        self.geo_data = self.load_map_geojson()

    def load_integrated_data(self):
        """유효구인인원 통합 엑셀 데이터 로드"""
        file_path = os.path.join(BASE_PATH_INTEGRATED, '유효구인인원_직종규모형태학력_통합데이터.xlsx')
        if not os.path.exists(file_path): 
            print(f"경고: 통합 데이터 파일 없음 ({file_path})")
            return pd.DataFrame()
        try:
            df = pd.read_excel(file_path)   # openpyxl 없으면 경고만 찍고 넘어감
            return df
        except Exception as e:
            print(f"통합 데이터 로드 실패: {e}")
            return pd.DataFrame()

    def load_trend_data(self):
        openings_list = []
        seekers_list = []
        
        for year in range(2022, 2026):
            # --- 구인 데이터 ---
            o_path = os.path.join(BASE_PATH_TREND, f'{year}_구인인원.csv')
            if os.path.exists(o_path):
                try:
                    try:
                        temp = pd.read_csv(o_path, encoding='cp949')
                    except:
                        temp = pd.read_csv(o_path, encoding='utf-8')

                    cols = ['시도', '직종', '학력', '규모', '인원수', '연도', '월']
                    valid_cols = [c for c in cols if c in temp.columns]
                    openings_list.append(temp[valid_cols])
                except Exception as e:
                    print(f"{year}년 구인 데이터 로드 실패:", e)

            # --- 구직 데이터 ---
            s_name = '2024_데이터셋.csv' if year == 2024 else f'{year}_구직자수.csv'
            s_path = os.path.join(BASE_PATH_TREND, s_name)
            if os.path.exists(s_path):
                try:
                    try:
                        temp = pd.read_csv(s_path, encoding='cp949')
                    except:
                        temp = pd.read_csv(s_path, encoding='utf-8')
                    
                    if year == 2025:
                        if '학력' in temp.columns:
                            temp = temp.rename(columns={'학력': '나이'})
                        if '규모' in temp.columns:
                            temp = temp.rename(columns={'규모': '학력'})
                    elif '경력' in temp.columns and '학력' not in temp.columns:
                        temp = temp.rename(columns={'경력': '학력'})

                    temp = temp.loc[:, ~temp.columns.duplicated()]
                    cols = ['시도', '직종', '학력', '인원수', '연도', '월']
                    valid_cols = [c for c in cols if c in temp.columns]
                    seekers_list.append(temp[valid_cols])
                except Exception as e:
                    print(f"{year}년 구직 데이터 로드 실패:", e)

        # ✅ 구인 데이터 통합 (없어도 컬럼은 확보)
        if openings_list:
            df_o = pd.concat(openings_list, ignore_index=True)
        else:
            print("[경고] 구인 데이터 파일을 하나도 읽지 못했습니다.")
            df_o = pd.DataFrame(columns=['시도', '직종', '학력', '규모', '인원수', '연도', '월'])

        # ✅ 구직 데이터 통합 (없어도 컬럼은 확보)
        if seekers_list:
            df_s = pd.concat(seekers_list, ignore_index=True)
        else:
            print("[경고] 구직 데이터 파일을 하나도 읽지 못했습니다.")
            df_s = pd.DataFrame(columns=['시도', '직종', '학력', '인원수', '연도', '월'])

        # 날짜/학력 전처리
        if not df_o.empty:
            df_o['학력'] = df_o['학력'].replace('무관', '학력무관')
            df_o['날짜'] = pd.to_datetime(df_o['연도'].astype(str) + '-' + df_o['월'].astype(str))
        else:
            df_o['날짜'] = pd.Series(dtype='datetime64[ns]')

        if not df_s.empty:
            df_s['학력'] = df_s['학력'].replace('무관', '학력무관')
            df_s['날짜'] = pd.to_datetime(df_s['연도'].astype(str) + '-' + df_s['월'].astype(str))
        else:
            df_s['날짜'] = pd.Series(dtype='datetime64[ns]')

        # 디버그용 출력
        print("df_open shape:", df_o.shape, "columns:", list(df_o.columns))
        print("df_seek shape:", df_s.shape, "columns:", list(df_s.columns))

        return df_o, df_s

    def load_map_geojson(self):
        if os.path.exists(JSON_PATH):
            try:
                with open(JSON_PATH, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print("지도 JSON 로드 실패:", e)
                return None
        else:
            print(f"지도 파일 없음: {JSON_PATH}")
            return None


# 2. 메인 윈도우 클래스
class JobDashboard(Tk): 
    def __init__(self): 
        Tk.__init__(self) 
        self.title("수원투용 - 대한민국 채용/구직 트렌드 대시보드") 
        self.geometry("1600x950") 

        self.loader = DataLoader()
        self.geo_data = self.loader.load_map_geojson()

        self.paned_window = PanedWindow(self, orient=HORIZONTAL) 
        self.paned_window.pack(fill=BOTH, expand=1) 

        self.left_frame = Frame(self.paned_window, width=280, bg='#f0f0f0', relief="raised") 
        self.paned_window.add(self.left_frame) 

        self.right_frame = Frame(self.paned_window, bg='white') 
        self.paned_window.add(self.right_frame) 

        self.create_controls() 
        self.create_plots()    

    def create_controls(self):
        lbl_title = Label(self.left_frame, text="🇰🇷 전국 분석 필터", font=("맑은 고딕", 16, "bold"), bg='#f0f0f0') 
        lbl_title.pack(pady=20, padx=10, anchor="w") 

        # 1. 지역 선택
        lbl_region = Label(self.left_frame, text="1. 지역 선택 (인구순):", font=("맑은 고딕", 11), bg='#f0f0f0') 
        lbl_region.pack(anchor="w", padx=10) 
        
        self.region_var = StringVar(value="전체") 
        if not self.loader.df_open.empty:
            raw_regions = list(self.loader.df_open['시도'].unique())
            valid_regions = [x for x in raw_regions if isinstance(x, str) and x != '총계']
            sorted_regions = sorted(valid_regions, key=lambda x: POPULATION_ORDER.index(x) if x in POPULATION_ORDER else 999)
            regions = ["전체"] + sorted_regions
        else:
            regions = ["전체"]
        self.combo_region = ttk.Combobox(self.left_frame, textvariable=self.region_var, values=regions, state="readonly")
        self.combo_region.pack(fill=X, padx=10, pady=5)
        self.combo_region.bind("<<ComboboxSelected>>", self.update_graph) 

        # 2. 직종 선택
        lbl_job = Label(self.left_frame, text="2. 직종 선택:", font=("맑은 고딕", 11), bg='#f0f0f0') 
        lbl_job.pack(anchor="w", padx=10, pady=(20, 0)) 

        self.job_var = StringVar(value="전체")
        if not self.loader.df_open.empty:
            job_list = ["전체"] + sorted([x for x in self.loader.df_open['직종'].unique() if isinstance(x, str)])
        else:
            job_list = ["전체"]
        self.combo_job = ttk.Combobox(self.left_frame, textvariable=self.job_var, values=job_list, state="readonly")
        self.combo_job.pack(fill=X, padx=10, pady=5)
        self.combo_job.bind("<<ComboboxSelected>>", self.update_graph)

        # 3. 학력 선택
        lbl_edu = Label(self.left_frame, text="3. 학력 선택:", font=("맑은 고딕", 11), bg='#f0f0f0') 
        lbl_edu.pack(anchor="w", padx=10, pady=(20, 0)) 

        self.edu_var = StringVar(value="전체")
        if not self.loader.df_open.empty:
            raw_edu = list(self.loader.df_open['학력'].unique())
            valid_edu = [x for x in raw_edu if isinstance(x, str) and x != '전체']
            sorted_edu = sorted(valid_edu, key=lambda x: EDU_ORDER.index(x) if x in EDU_ORDER else 999)
            edu_list = ["전체"] + sorted_edu
        else:
            edu_list = ["전체"]
        self.combo_edu = ttk.Combobox(self.left_frame, textvariable=self.edu_var, values=edu_list, state="readonly")
        self.combo_edu.pack(fill=X, padx=10, pady=5)
        self.combo_edu.bind("<<ComboboxSelected>>", self.update_graph)

        info = (
            "📌 [대시보드 안내]\n\n"
            "1. 좌측: 전국 채용 지도\n"
            "   - 지역별 구인 규모\n"
            "   - 지역 선택 시 줌인\n\n"
            "2. 우측 상단: 시계열 추세\n"
            "   - 2022~2025 변화\n\n"
            "3. 우측 중단: 직종별 채용\n"
            "   - 통합 데이터 기반\n"
            "   - 전국/지역별 수요 파악\n\n"
            "4. 우측 하단: 수급 불일치\n"
            "   - (구직자 - 구인인원)\n"
            "   - 모든 학력 비교"
        )
        lbl_info = Label(self.left_frame, text=info, justify="left", bg="white", relief="solid", bd=1, padx=10, pady=20)
        lbl_info.pack(fill=X, padx=10)

        lbl_team = Label(self.left_frame, text="Team 수원투용", font=("Arial", 12, "bold"), fg="gray", bg='#f0f0f0')
        lbl_team.pack(side=BOTTOM, pady=20)

    def create_plots(self):
        self.fig = Figure(figsize=(14, 10), dpi=100) 
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.right_frame) 
        self.canvas.get_tk_widget().pack(fill=BOTH, expand=True) 
        self.draw_graphs()

    def draw_graphs(self):
        self.fig.clear() 
        
        sel_region = self.region_var.get() 
        sel_job = self.job_var.get()
        sel_edu = self.edu_var.get()
        
        # 데이터 복사
        df_o_all = self.loader.df_open[self.loader.df_open['시도'] != '총계'].copy()
        df_o = df_o_all.copy()
        df_s = self.loader.df_seek[self.loader.df_seek['시도'] != '총계'].copy()
        df_int = self.loader.df_integrated.copy() # 통합 데이터

        # [필터링]
        if sel_region != "전체":
            df_o = df_o[df_o['시도'] == sel_region]
            df_s = df_s[df_s['시도'] == sel_region]

        if sel_job != "전체":
            df_o = df_o[df_o['직종'] == sel_job]
            df_s = df_s[df_s['직종'] == sel_job]
            # 통합 데이터 필터링
            if not df_int.empty and '직종' in df_int.columns:
                df_int = df_int[df_int['직종'] == sel_job]

        df_o_edu = df_o.copy()
        df_s_edu = df_s.copy()
        if sel_edu != "전체":
            df_o_edu = df_o_edu[df_o_edu['학력'] == sel_edu]
            df_s_edu = df_s_edu[df_s_edu['학력'] == sel_edu]

        gs = self.fig.add_gridspec(3, 2, width_ratios=[1.5, 1]) 

        
        # [1] 좌측: 지도 시각화
        
        ax1 = self.fig.add_subplot(gs[:, 0]) 
        
        map_source = df_o_all
        if sel_job != "전체": map_source = map_source[map_source['직종'] == sel_job]
        if sel_edu != "전체": map_source = map_source[map_source['학력'] == sel_edu]

        map_stats = map_source.groupby('시도')['인원수'].sum().reset_index()
        max_val = map_stats['인원수'].max() if not map_stats.empty else 1
        
        ax1.set_clip_on(True)

        if self.geo_data:
            patches = []
            for feature in self.geo_data['features']:
                eng_name = feature['properties'].get('name')
                kor_name = GEO_NAME_MAP.get(eng_name, eng_name)
                
                val = 0
                row = map_stats[map_stats['시도'] == kor_name]
                if not row.empty: val = row.iloc[0]['인원수']

                edge_c = 'red' if kor_name == sel_region else 'gray'
                line_w = 2 if kor_name == sel_region else 0.5
                face_c = plt.cm.Blues(val / max_val * 0.8 + 0.1) 

                for coords in feature['geometry']['coordinates']:
                    if feature['geometry']['type'] == 'Polygon':
                        patches.append(Polygon(coords, closed=True, facecolor=face_c, edgecolor=edge_c, linewidth=line_w, alpha=0.6))
                    else:
                        for sub in coords:
                            patches.append(Polygon(sub, closed=True, facecolor=face_c, edgecolor=edge_c, linewidth=line_w, alpha=0.6))
            
            if patches: ax1.add_collection(PatchCollection(patches, match_original=True))

        if not map_stats.empty:
            map_stats['lat'] = map_stats['시도'].map(lambda x: PROVINCE_COORDS.get(x, (0,0))[0])
            map_stats['lon'] = map_stats['시도'].map(lambda x: PROVINCE_COORDS.get(x, (0,0))[1])
            map_stats = map_stats[map_stats['lat'] > 0]

            scatter = ax1.scatter(
                map_stats['lon'], map_stats['lat'], 
                s=map_stats['인원수'] * 0.05, 
                c=map_stats['인원수'], 
                cmap='coolwarm', alpha=0.8, edgecolors='k', zorder=5
            )
            
            for idx, row in map_stats.iterrows():
                name = row['시도']
                if sel_region != "전체" and name != sel_region: continue
                
                weight = 'bold' if name == sel_region else 'normal'
                color = 'red' if name == sel_region else 'black'
                size = 11 if name == sel_region else 8
                label_txt = f"{name}\n{int(row['인원수']):,}"
                ax1.text(row['lon'], row['lat'], label_txt, fontsize=size, color=color, ha='center', va='center', fontweight=weight, zorder=6)
            
            cbar = self.fig.colorbar(scatter, ax=ax1, orientation='horizontal', pad=0.05, shrink=0.8)
            cbar.set_label(f'채용 공고 수 ({sel_job})')

        if sel_region != "전체" and sel_region in PROVINCE_COORDS:
            lat, lon = PROVINCE_COORDS[sel_region]
            zoom = ZOOM_LEVELS.get(sel_region, 0.5)
            ax1.set_xlim(lon - zoom, lon + zoom)
            ax1.set_ylim(lat - zoom, lat + zoom)
            title_txt = f"{sel_region} 채용 지도 (Zoom-in)"
        else:
            ax1.set_xlim(125, 130) 
            ax1.set_ylim(33, 39)
            title_txt = f"전국 채용 지도 ({sel_job})"

        ax1.set_aspect('equal') 
        ax1.set_title(title_txt, fontsize=14)
        ax1.axis('off') 

        
        # [2] 우측 상단: 시계열 추세
        
        ax2 = self.fig.add_subplot(gs[0, 1])
        if not df_o_edu.empty and not df_s_edu.empty:
            t_o = df_o_edu.groupby('날짜')['인원수'].sum()
            t_s = df_s_edu.groupby('날짜')['인원수'].sum()
            
            ax2.plot(t_o.index, t_o.values, marker='o', markersize=3, label='구인', color='blue', linewidth=2)
            ax2.plot(t_s.index, t_s.values, marker='s', markersize=3, label='구직', color='red', linestyle='--', linewidth=2)
            
            for i, (d, v) in enumerate(zip(t_o.index, t_o.values)):
                 offset = 10 if i % 2 == 0 else -10
                 ax2.annotate(f"{int(v):,}", (d, v), xytext=(0, offset), textcoords='offset points', 
                              fontsize=7, color='blue', ha='center')
            for i, (d, v) in enumerate(zip(t_s.index, t_s.values)):
                 offset = -15 if i % 2 == 0 else 15
                 ax2.annotate(f"{int(v):,}", (d, v), xytext=(0, offset), textcoords='offset points', 
                              fontsize=7, color='red', ha='center')

            all_vals = np.concatenate([t_o.values, t_s.values])
            mn, mx = all_vals.min(), all_vals.max()
            margin = (mx - mn) * 0.2 if mx != mn else mx * 0.1
            ax2.set_ylim(max(0, mn - margin), mx + margin)

            ax2.set_title(f"시기별 추세 ({sel_region}, {sel_job})", fontsize=12)
            ax2.legend(fontsize=9)
            ax2.grid(axis='y', linestyle='--', alpha=0.5)
            ax2.xaxis.set_major_formatter(mdates.DateFormatter('%y.%m'))
            ax2.tick_params(axis='x', rotation=45, labelsize=8)
        else:
            ax2.text(0.5, 0.5, "데이터 없음", ha='center', va='center')

       
        # [3] 우측 중단: 직종별 채용 규모 (전국 데이터 기반) 
        
        ax3 = self.fig.add_subplot(gs[1, 1]) 
        
        # 통합 데이터(df_int) 또는 구인 데이터(df_o) 활용하여 직종별 규모 파악
        # 여기서는 df_o(필터링된 구인 데이터)를 사용하여 현재 조건에서의 직종 분포를 보여줌
        if not df_o.empty:
            if sel_job == "전체":
                # 전체 직종 중 Top 7
                job_stats = df_o.groupby('직종')['인원수'].sum().sort_values().tail(7)
                title_sub = f"직종별 채용 규모 (Top 7) - {sel_region}"
            else:
                # 선택된 직종 내의 세부 통계가 없으면 그냥 해당 직종 총량 표시
                # (만약 통합 데이터에 세부 직종이 있다면 그걸 활용)
                job_stats = df_o.groupby('직종')['인원수'].sum()
                title_sub = f"'{sel_job}' 채용 규모 - {sel_region}"

            if not job_stats.empty:
                bars = ax3.barh(job_stats.index, job_stats.values, color='lightgreen', edgecolor='gray')
                for bar in bars:
                    width = bar.get_width()
                    ax3.text(width, bar.get_y() + bar.get_height()/2, f' {int(width):,}명', va='center', fontsize=8)
                
                ax3.set_title(title_sub, fontsize=12)
                ax3.set_xlabel("인원수 (명)", fontsize=9)
                ax3.grid(axis='x', linestyle='--', alpha=0.5)
                ax3.tick_params(labelsize=8)
            else:
                ax3.text(0.5, 0.5, "데이터 없음", ha='center')
        else:
            ax3.text(0.5, 0.5, "데이터 없음", ha='center')

        
        # [4] 우측 하단: 수급 불일치 (전체 학력 비교)
        
        ax4 = self.fig.add_subplot(gs[2, 1]) 
        
        if not df_o.empty and not df_s.empty:
            o_sum = df_o.groupby('학력')['인원수'].sum()
            s_sum = df_s.groupby('학력')['인원수'].sum()
            
            common_edu = sorted(list(set(o_sum.index) & set(s_sum.index)))
            common_edu = [e for e in common_edu if e != '전체']
            common_edu = sorted(common_edu, key=lambda x: EDU_ORDER.index(x) if x in EDU_ORDER else 999)

            gaps = [s_sum.get(e, 0) - o_sum.get(e, 0) for e in common_edu]
            
            if gaps:
                colors = ['blue' if v > 0 else 'red' for v in gaps]
                bars = ax4.barh(common_edu, gaps, color=colors, alpha=0.7)
                
                for bar in bars:
                    width = bar.get_width()
                    x_pos = width if width > 0 else width
                    align = 'left' if width > 0 else 'right'
                    offset = 5 if width > 0 else -5
                    ax4.text(x_pos + offset, bar.get_y() + bar.get_height()/2, f'{int(width):,}', 
                             va='center', ha=align, fontsize=8, color='black')

                ax4.axvline(0, color='black', linewidth=0.8) 
                ax4.set_title("학력별 수급 불일치 (전체 학력 비교)", fontsize=12)
                ax4.set_xlabel("인원 차이 (명) [ (+)구직자 과잉 / (-)일자리 과잉 ]", fontsize=9)
                ax4.grid(axis='x', linestyle='--', alpha=0.5)
                ax4.tick_params(labelsize=8)
            else:
                ax4.text(0.5, 0.5, "비교 데이터 부족", ha='center')
        else:
            ax4.text(0.5, 0.5, "데이터 없음", ha='center')

        self.fig.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.08, wspace=0.2, hspace=0.4)
        self.canvas.draw() 

    def update_graph(self, event=None):
        self.draw_graphs()

# 3. 실행
if __name__ == "__main__": 
    app = JobDashboard() 
    print("대시보드 창이 열렸습니다.")
    app.mainloop()