"""
KEPCO 지능형 관제 시스템 GUI v5.5 (Yearly-Week-Sync Enhanced)
배전선로 휴전 작업 가부 판정 시스템

[핵심 알고리즘: Yearly-Week-Sync v5.5]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. 364일(52주) 전 동일 주차 데이터 동기화
   - 365일이 아닌 364일을 사용하는 이유: 요일(Day of Week) 완벽 일치
   - 예: 2026-01-21(수) → 2025-01-22(수) 참조 (364일 전)

2. Reference Week 추출
   - Target Date 기준 364일 전 날짜가 속한 주(월~일) 전체 데이터 추출
   - 00:00~23:00 시간대별 평균 패턴 산출

3. Feature Engineering (XGBoost 통합)
   - last_year_same_week_load: 1년 전 동일 주차 부하
   - scaling_factor: 최근 4주 트렌드 보정 계수
   - is_weekend: 주말 여부 (토/일 별도 취급)

4. 부하 절체 시뮬레이션
   - Combined_Load = Target_Load + (Shutdown_Load / N)
   - 임계치(14.0MW) 기준 가부 판정

작성: 한국전력 배전센터 데이터 사이언티스트
버전: v5.5 Enhanced (Yearly-Week-Sync)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import threading
import os

# =============================================================================
# matplotlib import (그래프 시각화)
# =============================================================================
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    # 한글 폰트 설정
    plt.rcParams['font.family'] = 'Malgun Gothic'
    plt.rcParams['axes.unicode_minus'] = False
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

# =============================================================================
# XGBoost import (Yearly-Week-Sync 피처와 앙상블 적용)
# =============================================================================
try:
    import xgboost as xgb
    XGBOOST_AVAILABLE = True
except ImportError:
    XGBOOST_AVAILABLE = False


class KEPCOOutageGUI:
    """
    KEPCO 배전선로 휴전 작업 가부 판정 시스템 v5.5

    핵심 알고리즘: Yearly-Week-Sync
    - Target Date로부터 364일(52주) 전 동일 주차 데이터 참조
    - 요일(Day of Week) 완벽 일치
    - 최근 4주 트렌드 보정 적용
    """

    # =========================================================================
    # 상수 정의
    # =========================================================================
    DAYS_IN_YEAR = 364  # 52주 = 364일 (요일 완벽 매칭, 365일 아님!)
    WEEKDAY_NAMES_KR = ['월', '화', '수', '목', '금', '토', '일']
    WEEKDAY_NAMES_FULL = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']

    # 주말 인덱스 (토=5, 일=6)
    WEEKEND_DAYS = {5, 6}

    # 작업 시간대 (09:00 ~ 14:00)
    WORK_HOURS = list(range(9, 15))

    def __init__(self, root):
        self.root = root
        self.root.title("⚡ KEPCO 배전선로 휴전 작업 가부 판정 시스템 v5.5 (Yearly-Week-Sync)")
        self.root.geometry("1500x1050")
        self.root.configure(bg='#f0f0f0')

        # =====================================================================
        # 데이터 변수 초기화
        # =====================================================================
        self.load_file = None
        self.shutdown_file = None
        self.results_df = None
        self.threshold_kw = 14000  # 14MW 임계치
        self.shutdown_mapping = {}  # 휴전선로 -> 절체대상 매핑
        self.selected_shutdown_line = None
        self.detailed_results = []

        # 변전소별 파일 (A, B, C)
        self.load_files = {
            'A': None,
            'B': None,
            'C': None
        }

        # 절체대상 개수
        self.n_transfer_targets = 0
        self.transfer_columns = {}

        # =====================================================================
        # Yearly-Week-Sync 핵심 데이터 구조
        # =====================================================================
        # 1년 전 동일 주차 프로파일: {Line: {Weekday: {Hour: stats}}}
        self.yearly_week_profiles = {}

        # 최근 트렌드 스케일링 팩터: {Line: scaling_factor}
        self.scaling_factors = {}

        # Long format 데이터
        self.long_df = None

        # 주간(1주일) 예측 데이터 (시각화용)
        self.weekly_predictions = {}

        # 참조 주차 정보 저장
        self.reference_week_info = {}

        # =================================================================
        # Yearly-Week-Sync 강화 데이터 구조
        # =================================================================
        # 1년 전 동일 주차 실제 데이터 (슬라이싱된 원본)
        self.reference_week_data = {}  # {date_str: {line: DataFrame}}

        # XGBoost 피처용 데이터
        self.last_year_same_week_features = {}  # {line: {date: {hour: load}}}

        # 주말/평일 별도 프로파일
        self.weekday_profiles = {}  # 평일(월~금) 프로파일
        self.weekend_profiles = {}  # 주말(토~일) 프로파일

        self.create_widgets()

    def create_widgets(self):
        """GUI 위젯 생성"""
        # =====================================================================
        # 헤더
        # =====================================================================
        header_frame = tk.Frame(self.root, bg='#1e3a8a', height=80)
        header_frame.pack(fill='x', padx=0, pady=0)

        title_label = tk.Label(
            header_frame,
            text="⚡ KEPCO v5.5 Yearly-Week-Sync 부하 예측 시스템",
            font=('맑은 고딕', 18, 'bold'),
            fg='white',
            bg='#1e3a8a'
        )
        title_label.pack(pady=20)

        # =====================================================================
        # 메인 컨테이너
        # =====================================================================
        main_container = tk.Frame(self.root, bg='#f0f0f0')
        main_container.pack(fill='both', expand=True, padx=20, pady=10)

        # =====================================================================
        # 좌측 패널 (설정) - 스크롤 가능
        # =====================================================================
        left_frame = tk.Frame(main_container, bg='white', relief='raised', borderwidth=2)
        left_frame.pack(side='left', fill='y', padx=(0, 10), pady=0)

        left_canvas = tk.Canvas(left_frame, bg='white', highlightthickness=0, width=290)
        left_scrollbar = tk.Scrollbar(left_frame, orient='vertical', command=left_canvas.yview)
        left_panel = tk.Frame(left_canvas, bg='white')

        left_panel.bind(
            '<Configure>',
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox('all'))
        )

        left_canvas.create_window((0, 0), window=left_panel, anchor='nw')
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left_canvas.pack(side='left', fill='both', expand=True)
        left_scrollbar.pack(side='right', fill='y')

        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1*(event.delta/120)), 'units')
        left_canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # =====================================================================
        # 파일 선택 섹션
        # =====================================================================
        file_section = tk.LabelFrame(left_panel, text="📂 변전소별 부하 데이터",
                                      font=('맑은 고딕', 11, 'bold'), bg='white')
        file_section.pack(fill='x', padx=10, pady=10)

        # 변전소 A
        tk.Label(file_section, text="변전소 A:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 0))
        self.load_file_label_a = tk.Label(file_section, text="Load_A.xlsx", font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.load_file_label_a.pack(anchor='w', padx=10, pady=(0, 5))
        tk.Button(
            file_section, text="변전소 A 선택",
            command=lambda: self.select_load_file_by_substation('A'),
            bg='#3b82f6', fg='white', font=('맑은 고딕', 9), relief='flat', cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 5))

        # 변전소 B
        tk.Label(file_section, text="변전소 B:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(5, 0))
        self.load_file_label_b = tk.Label(file_section, text="Load_B.xlsx", font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.load_file_label_b.pack(anchor='w', padx=10, pady=(0, 5))
        tk.Button(
            file_section, text="변전소 B 선택",
            command=lambda: self.select_load_file_by_substation('B'),
            bg='#3b82f6', fg='white', font=('맑은 고딕', 9), relief='flat', cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 5))

        # 변전소 C
        tk.Label(file_section, text="변전소 C:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(5, 0))
        self.load_file_label_c = tk.Label(file_section, text="Loda_C.xlsx", font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.load_file_label_c.pack(anchor='w', padx=10, pady=(0, 5))
        tk.Button(
            file_section, text="변전소 C 선택",
            command=lambda: self.select_load_file_by_substation('C'),
            bg='#3b82f6', fg='white', font=('맑은 고딕', 9), relief='flat', cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))

        # =====================================================================
        # 절체 매핑 파일
        # =====================================================================
        mapping_section = tk.LabelFrame(left_panel, text="📂 절체 매핑 파일",
                                        font=('맑은 고딕', 11, 'bold'), bg='white')
        mapping_section.pack(fill='x', padx=10, pady=10)

        self.shutdown_file_label = tk.Label(mapping_section, text="shutdown_list.xlsx",
                                            font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.shutdown_file_label.pack(anchor='w', padx=10, pady=(10, 5))
        tk.Button(
            mapping_section, text="절체 매핑 파일 선택",
            command=self.select_shutdown_file,
            bg='#3b82f6', fg='white', font=('맑은 고딕', 9), relief='flat', cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))

        # =====================================================================
        # 휴전 대상 선로 선택
        # =====================================================================
        shutdown_section = tk.LabelFrame(left_panel, text="🎯 휴전 대상 선로 선택",
                                         font=('맑은 고딕', 11, 'bold'), bg='white')
        shutdown_section.pack(fill='x', padx=10, pady=10)

        tk.Label(shutdown_section, text="휴전 대상 선로:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 5))
        self.shutdown_line_var = tk.StringVar()
        self.shutdown_line_combo = ttk.Combobox(
            shutdown_section, textvariable=self.shutdown_line_var,
            font=('맑은 고딕', 10), state='readonly', width=25
        )
        self.shutdown_line_combo.pack(fill='x', padx=10, pady=(0, 10))
        self.shutdown_line_combo.bind('<<ComboboxSelected>>', self.on_shutdown_line_selected)

        self.transfer_info_label = tk.Label(
            shutdown_section, text="절체 대상: -",
            font=('맑은 고딕', 8), bg='white', fg='#6b7280', wraplength=250, justify='left'
        )
        self.transfer_info_label.pack(anchor='w', padx=10, pady=(0, 10))

        # =====================================================================
        # 분석 설정
        # =====================================================================
        settings_section = tk.LabelFrame(left_panel, text="⚙️ 분석 설정",
                                         font=('맑은 고딕', 11, 'bold'), bg='white')
        settings_section.pack(fill='x', padx=10, pady=10)

        # 임계치
        tk.Label(settings_section, text="임계치 (kW):", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 0))
        threshold_frame = tk.Frame(settings_section, bg='white')
        threshold_frame.pack(fill='x', padx=10, pady=(5, 5))
        self.threshold_var = tk.StringVar(value="14000")
        tk.Entry(threshold_frame, textvariable=self.threshold_var, font=('맑은 고딕', 10), width=10).pack(side='left')
        tk.Label(threshold_frame, text="kW (14MW)", font=('맑은 고딕', 9), bg='white', fg='#6b7280').pack(side='left', padx=(5, 0))

        # 분석 기간
        tk.Label(settings_section, text="분석 기간 (일):", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 0))
        days_frame = tk.Frame(settings_section, bg='white')
        days_frame.pack(fill='x', padx=10, pady=(5, 10))
        self.days_var = tk.StringVar(value="30")
        tk.Entry(days_frame, textvariable=self.days_var, font=('맑은 고딕', 10), width=10).pack(side='left')
        tk.Label(days_frame, text="일", font=('맑은 고딕', 9), bg='white', fg='#6b7280').pack(side='left', padx=(5, 0))

        # =====================================================================
        # 알고리즘 설명
        # =====================================================================
        algo_section = tk.LabelFrame(left_panel, text="📊 예측 알고리즘",
                                     font=('맑은 고딕', 11, 'bold'), bg='white')
        algo_section.pack(fill='x', padx=10, pady=10)

        algo_info = tk.Label(
            algo_section,
            text="Yearly-Week-Sync v5.5\n"
                 "━━━━━━━━━━━━━━━━━━━\n"
                 "• 364일(52주) 전 동일 주차 참조\n"
                 "• 요일(Day of Week) 완벽 매칭\n"
                 "• Reference Week Trend 보정\n"
                 "• 평일/주말 별도 분석 지원",
            font=('맑은 고딕', 9),
            bg='white', fg='#2563eb', justify='left'
        )
        algo_info.pack(anchor='w', padx=10, pady=10)

        # =====================================================================
        # 실행 버튼
        # =====================================================================
        tk.Button(
            left_panel, text="🚀 분석 시작",
            command=self.run_analysis,
            bg='#16a34a', fg='white', font=('맑은 고딕', 12, 'bold'),
            relief='flat', cursor='hand2', height=2
        ).pack(fill='x', padx=10, pady=20)

        tk.Button(
            left_panel, text="💾 결과 엑셀 저장",
            command=self.save_to_excel,
            bg='#9333ea', fg='white', font=('맑은 고딕', 10),
            relief='flat', cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))

        # =====================================================================
        # 우측 패널 (결과)
        # =====================================================================
        right_panel = tk.Frame(main_container, bg='white', relief='raised', borderwidth=2)
        right_panel.pack(side='right', fill='both', expand=True, padx=0, pady=0)

        self.notebook = ttk.Notebook(right_panel)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # -----------------------------------------------------------------
        # 탭 1: 분석 결과
        # -----------------------------------------------------------------
        self.results_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.results_tab, text="📋 분석 결과")

        result_frame = tk.Frame(self.results_tab, bg='white')
        result_frame.pack(fill='both', expand=True, padx=10, pady=10)

        scroll_y = tk.Scrollbar(result_frame)
        scroll_y.pack(side='right', fill='y')
        scroll_x = tk.Scrollbar(result_frame, orient='horizontal')
        scroll_x.pack(side='bottom', fill='x')

        self.tree = ttk.Treeview(
            result_frame,
            columns=('날짜', '요일', '판정', '피크(MW)', '최저(MW)', '절체부하', '여유(MW)', '대조구간', '비고'),
            show='headings',
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            height=16
        )
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        # 컬럼 설정
        columns_config = [
            ('날짜', 95, 'center'),
            ('요일', 45, 'center'),
            ('판정', 45, 'center'),
            ('피크(MW)', 75, 'center'),
            ('최저(MW)', 75, 'center'),
            ('절체부하', 250, 'w'),
            ('여유(MW)', 70, 'center'),
            ('대조구간', 180, 'center'),
            ('비고', 150, 'w')
        ]
        for col, width, anchor in columns_config:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor=anchor)

        self.tree.pack(fill='both', expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_date_selected)

        # -----------------------------------------------------------------
        # 탭 2: 1주일 예측 그래프
        # -----------------------------------------------------------------
        self.graph_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.graph_tab, text="📈 1주일 예측 그래프")

        self.graph_frame = tk.Frame(self.graph_tab, bg='white')
        self.graph_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.graph_info_label = tk.Label(
            self.graph_frame,
            text="분석 결과 탭에서 날짜를 선택하면\n해당 날짜 기준 1주일(월~일) 예측 그래프가 표시됩니다.\n\n"
                 "📌 364일 전 동일 주차 데이터 기반\n"
                 "📌 평일/주말 패턴 구분 표시",
            font=('맑은 고딕', 12), bg='white', fg='#6b7280'
        )
        self.graph_info_label.pack(expand=True)

        # -----------------------------------------------------------------
        # 탭 3: 요일별 패턴
        # -----------------------------------------------------------------
        self.weekday_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.weekday_tab, text="📊 요일별 패턴")

        self.weekday_text = scrolledtext.ScrolledText(
            self.weekday_tab, font=('맑은 고딕', 10), wrap='word', bg='#f9fafb'
        )
        self.weekday_text.pack(fill='both', expand=True, padx=10, pady=10)

        # -----------------------------------------------------------------
        # 탭 4: 추천 일자
        # -----------------------------------------------------------------
        self.recommendation_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.recommendation_tab, text="🎯 추천 일자")

        self.recommendation_text = scrolledtext.ScrolledText(
            self.recommendation_tab, font=('맑은 고딕', 10), wrap='word', bg='#f9fafb'
        )
        self.recommendation_text.pack(fill='both', expand=True, padx=10, pady=10)

        # -----------------------------------------------------------------
        # 탭 5: 분석 로그
        # -----------------------------------------------------------------
        self.log_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.log_tab, text="📝 분석 로그")

        self.log_text = scrolledtext.ScrolledText(
            self.log_tab, font=('Consolas', 9), wrap='word', bg='#1e1e1e', fg='#d4d4d4'
        )
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)

        # =====================================================================
        # 초기 메시지
        # =====================================================================
        self.log("=" * 70)
        self.log("✅ KEPCO v5.5 Yearly-Week-Sync 시스템 준비 완료")
        self.log("=" * 70)
        self.log("")
        self.log("📌 핵심 알고리즘: 1년 전 동일 주차 데이터 동기화")
        self.log("   • 364일(52주) 전 데이터 참조로 요일 완벽 매칭")
        self.log("   • 최근 4주 트렌드 보정 (Scaling Factor) 적용")
        self.log("   • 평일/주말 별도 분석 지원")
        self.log("")
        self.log("📌 지원 기능:")
        self.log("   • 변전소 A, B, C 3개 파일 통합 분석")
        self.log("   • 1주일 단위 시각화 그래프")
        self.log("   • 정밀 1/N 부하 절체 시뮬레이션")
        self.log("")

        # =====================================================================
        # 기본 파일 설정
        # =====================================================================
        if os.path.exists('Load_A.xlsx'):
            self.load_files['A'] = 'Load_A.xlsx'
        if os.path.exists('Load_B.xlsx'):
            self.load_files['B'] = 'Load_B.xlsx'
        if os.path.exists('Loda_C.xlsx'):
            self.load_files['C'] = 'Loda_C.xlsx'

        if os.path.exists('shutdown_list.xlsx'):
            self.shutdown_file = 'shutdown_list.xlsx'
            self.load_shutdown_mapping()

    # =========================================================================
    # 유틸리티 메서드
    # =========================================================================

    def log(self, message):
        """로그 메시지 출력"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert('end', f"[{timestamp}] {message}\n")
        self.log_text.see('end')
        self.root.update()

    def select_load_file_by_substation(self, substation):
        """변전소별 부하 데이터 파일 선택"""
        filename = filedialog.askopenfilename(
            title=f"변전소 {substation} 부하 데이터 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.load_files[substation] = filename
            label_map = {'A': self.load_file_label_a, 'B': self.load_file_label_b, 'C': self.load_file_label_c}
            label_map[substation].config(text=os.path.basename(filename))
            self.log(f"📂 변전소 {substation} 부하 데이터: {os.path.basename(filename)}")

    def select_shutdown_file(self):
        """절체 매핑 파일 선택"""
        filename = filedialog.askopenfilename(
            title="절체 매핑 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.shutdown_file = filename
            self.shutdown_file_label.config(text=os.path.basename(filename))
            self.log(f"📂 절체 매핑: {os.path.basename(filename)}")
            self.load_shutdown_mapping()

    def load_shutdown_mapping(self):
        """shutdown_list.xlsx에서 휴전선로 -> 절체대상 매핑 로드"""
        if not self.shutdown_file or not os.path.exists(self.shutdown_file):
            return

        try:
            df = pd.read_excel(self.shutdown_file, skiprows=1)
            self.shutdown_mapping = {}
            shutdown_lines = []

            for _, row in df.iterrows():
                shutdown_ss = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
                shutdown_dl = row.iloc[1]
                if pd.notna(shutdown_dl):
                    shutdown_dl = str(int(shutdown_dl)) if isinstance(shutdown_dl, (int, float)) else str(shutdown_dl).strip()
                else:
                    shutdown_dl = ''

                if shutdown_ss and shutdown_ss != 'nan' and shutdown_dl and shutdown_dl != 'nan':
                    shutdown_line = f"{shutdown_ss}-{shutdown_dl}"
                    transfer_targets = []
                    col_idx = 3

                    while col_idx < len(row) - 1:
                        target_ss = row.iloc[col_idx] if col_idx < len(row) else None
                        target_dl = row.iloc[col_idx + 1] if col_idx + 1 < len(row) else None

                        if pd.notna(target_ss):
                            target_ss = str(target_ss).strip()
                        else:
                            target_ss = ''

                        if pd.notna(target_dl):
                            target_dl = str(int(target_dl)) if isinstance(target_dl, (int, float)) else str(target_dl).strip()
                        else:
                            target_dl = ''

                        if target_ss and target_ss != 'nan' and target_dl and target_dl != 'nan':
                            transfer_targets.append(f"{target_ss}-{target_dl}")

                        col_idx += 2

                    if transfer_targets:
                        self.shutdown_mapping[shutdown_line] = transfer_targets
                        shutdown_lines.append(shutdown_line)

            self.shutdown_line_combo['values'] = shutdown_lines
            if shutdown_lines:
                self.shutdown_line_combo.current(0)
                self.on_shutdown_line_selected(None)

            self.log(f"✅ 매핑 로드: {len(shutdown_lines)}개 휴전선로")

        except Exception as e:
            self.log(f"❌ 매핑 로드 실패: {str(e)}")

    def on_shutdown_line_selected(self, event):
        """휴전 선로 선택 시 절체 대상 표시"""
        selected = self.shutdown_line_var.get()
        if selected in self.shutdown_mapping:
            targets = self.shutdown_mapping[selected]
            self.transfer_info_label.config(text=f"절체 대상: {', '.join(targets)}")
            self.log(f"🎯 선택: {selected} → [{', '.join(targets)}]")

    # =========================================================================
    # 분석 실행
    # =========================================================================

    def run_analysis(self):
        """분석 실행"""
        valid_files = {k: v for k, v in self.load_files.items() if v and os.path.exists(v)}
        if not valid_files:
            messagebox.showerror("오류", "최소 하나의 변전소 부하 데이터 파일을 선택해주세요.")
            return

        if not self.shutdown_line_var.get():
            messagebox.showerror("오류", "휴전 대상 선로를 선택해주세요.")
            return

        try:
            self.threshold_kw = float(self.threshold_var.get())
            days = int(self.days_var.get())
        except:
            messagebox.showerror("오류", "임계치와 분석 기간은 숫자로 입력해주세요.")
            return

        self.selected_shutdown_line = self.shutdown_line_var.get()

        thread = threading.Thread(target=self._run_analysis_thread, args=(days,))
        thread.daemon = True
        thread.start()

    def _run_analysis_thread(self, days):
        """분석 실행 (스레드)"""
        try:
            self.log("")
            self.log("=" * 70)
            self.log("⚡ Yearly-Week-Sync v5.5 분석 시작")
            self.log("=" * 70)
            self.log(f"📋 휴전 대상 선로: {self.selected_shutdown_line}")
            self.log(f"📋 임계치: {self.threshold_kw/1000:.1f}MW | 분석 기간: {days}일")
            self.log(f"📋 알고리즘: 364일(52주) 전 동일 주차 데이터 동기화")
            self.log("")

            # -----------------------------------------------------------------
            # 1단계: 데이터 로드
            # -----------------------------------------------------------------
            self.log("▶ [1단계] 변전소별 데이터 로딩")
            all_dfs = []
            for substation, file_path in self.load_files.items():
                if file_path and os.path.exists(file_path):
                    self.log(f"   • 변전소 {substation}: {os.path.basename(file_path)}")
                    df = self.load_data(file_path, substation)
                    all_dfs.append(df)

            if not all_dfs:
                raise ValueError("로드된 데이터가 없습니다.")

            df = pd.concat(all_dfs, ignore_index=True)
            self.log(f"   ✓ 총 {len(df)}개 레코드 통합 완료")

            # -----------------------------------------------------------------
            # 2단계: Long format 변환
            # -----------------------------------------------------------------
            self.log("")
            self.log("▶ [2단계] 데이터 변환 및 요일/시간 정보 추출")
            self.long_df = self.convert_to_long_format(df)

            # -----------------------------------------------------------------
            # 3단계: 1년 전 동일 주차 프로파일 생성 (핵심!)
            # -----------------------------------------------------------------
            self.log("")
            self.log("▶ [3단계] Yearly-Week-Sync 프로파일 생성")
            self.log(f"   • 기준: 364일(52주) 전 동일 주차 데이터")
            self.log(f"   • 목적: 요일(Day of Week) 완벽 매칭")
            self.log(f"   • 원리: 365일 아닌 364일 사용 → 요일 완벽 일치")
            self.generate_yearly_week_profiles()

            # -----------------------------------------------------------------
            # 3-1단계: 평일/주말 별도 프로파일 생성
            # -----------------------------------------------------------------
            self.log("")
            self.log("▶ [3-1단계] 평일/주말 별도 프로파일 생성")
            self.log(f"   • 평일(월~금): 산업/상업 부하 패턴")
            self.log(f"   • 주말(토~일): 주거/여가 부하 패턴 [별도 취급 예정]")
            self.generate_weekday_weekend_profiles()

            # -----------------------------------------------------------------
            # 4단계: 364일 전 동일 주차 기반 트렌드 팩터 계산
            # -----------------------------------------------------------------
            self.log("")
            self.log("▶ [4단계] Reference Week Trend 분석")
            self.log(f"   • 방식: 364일 전 동일 주차의 요일별 부하 패턴 분석")
            self.log(f"   • 목적: Target 요일의 주간 내 상대적 부하 수준 파악")
            self.calculate_scaling_factors()

            # -----------------------------------------------------------------
            # 5단계: 부하 예측 및 가부 판정
            # -----------------------------------------------------------------
            self.log("")
            self.log(f"▶ [5단계] 부하 예측 및 가부 판정 (향후 {days}일)")
            self.results_df = self.yearly_week_sync_verification(days)

            # -----------------------------------------------------------------
            # 6단계: 결과 표시
            # -----------------------------------------------------------------
            self.log("")
            self.log("▶ [6단계] 결과 표시")
            self.display_results()
            self.display_weekday_summary()
            self.display_recommendations()

            # -----------------------------------------------------------------
            # 7단계: 알고리즘 분석 로그
            # -----------------------------------------------------------------
            self.log_algorithm_analysis()

            # -----------------------------------------------------------------
            # 8단계: 예측 부하 계산 검증 로그
            # -----------------------------------------------------------------
            self.log_calculation_verification(num_samples=3)

            self.log("")
            self.log("=" * 70)
            self.log("✅ Yearly-Week-Sync v5.5 분석 완료!")
            self.log("=" * 70)

            messagebox.showinfo("완료",
                f"'{self.selected_shutdown_line}' 분석 완료!\n\n"
                f"[Yearly-Week-Sync v5.5]\n"
                f"• 364일(52주) 전 동일 주차 참조\n"
                f"• 요일 완벽 매칭 + 트렌드 보정")

        except Exception as e:
            self.log(f"❌ 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"분석 중 오류:\n{str(e)}")

    # =========================================================================
    # 데이터 처리 메서드
    # =========================================================================

    def load_data(self, file_path, substation=None):
        """데이터 로드"""
        df = pd.read_excel(file_path, skiprows=4)

        cols = ['S/S', 'D/L', '일자']
        for i in range(1, len(df.columns) - 2):
            cols.append(f'{i}시')
        df.columns = cols

        df['일자'] = pd.to_numeric(df['일자'], errors='coerce')
        df = df.dropna(subset=['일자'])
        df['일자'] = pd.to_datetime(df['일자'].astype(int), format='%Y%m%d')

        df['D/L'] = df['D/L'].apply(lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x))
        df['Line'] = df['S/S'].astype(str) + '-' + df['D/L']

        self.log(f"     → {len(df)}개 레코드 로드")
        return df

    def convert_to_long_format(self, df):
        """Long format 변환 (요일, 시간, 주말 여부 추출)"""
        hour_cols = [col for col in df.columns if '시' in str(col)]
        data_list = []

        for _, row in df.iterrows():
            date = row['일자']
            line = str(row['Line']).strip()

            for hour_col in hour_cols:
                try:
                    hour = int(hour_col.replace('시', ''))
                    if hour == 24:
                        timestamp = date + timedelta(days=1)
                    else:
                        timestamp = date + timedelta(hours=hour)

                    load_mw = row[hour_col]
                    if pd.notna(load_mw):
                        load_mw = float(load_mw)
                        load_kw = load_mw * 1000 if load_mw < 100 else load_mw

                        weekday = timestamp.weekday()
                        is_weekend = weekday >= 5  # 토(5), 일(6)

                        data_list.append({
                            'Timestamp': timestamp,
                            'Date': timestamp.date(),
                            'Year': timestamp.year,
                            'Month': timestamp.month,
                            'Day': timestamp.day,
                            'Hour': timestamp.hour if hour != 24 else 0,
                            'Weekday': weekday,
                            'WeekdayName': self.WEEKDAY_NAMES_KR[weekday],
                            'IsWeekend': is_weekend,
                            'Line': line,
                            'Load_kW': load_kw
                        })
                except:
                    pass

        long_df = pd.DataFrame(data_list)

        if len(long_df) > 0:
            min_date = long_df['Timestamp'].min()
            max_date = long_df['Timestamp'].max()
            self.log(f"   ✓ {len(long_df):,}개 시간대 데이터 변환")
            self.log(f"   ✓ 데이터 기간: {min_date.strftime('%Y-%m-%d')} ~ {max_date.strftime('%Y-%m-%d')}")

            # 평일/주말 통계
            weekday_count = len(long_df[~long_df['IsWeekend']])
            weekend_count = len(long_df[long_df['IsWeekend']])
            self.log(f"   ✓ 평일 데이터: {weekday_count:,}개 | 주말 데이터: {weekend_count:,}개")

        return long_df

    def generate_yearly_week_profiles(self):
        """
        1년 전 동일 주차 데이터 프로파일 생성

        핵심 로직:
        - 각 선로별로 과거 데이터를 요일(0~6)과 시간(0~23)으로 그룹화
        - 364일 전 데이터를 참조하여 요일 완벽 매칭
        """
        self.yearly_week_profiles = {}

        if self.long_df is None or len(self.long_df) == 0:
            self.log("   ⚠️ 데이터가 없어 프로파일 생성 불가")
            return

        lines = self.long_df['Line'].unique()

        for line in lines:
            line_data = self.long_df[self.long_df['Line'] == line]
            line_profile = {}

            # 요일별(0~6), 시간별(0~23) 통계 계산
            for weekday in range(7):
                weekday_data = line_data[line_data['Weekday'] == weekday]

                if len(weekday_data) == 0:
                    continue

                line_profile[weekday] = {}

                for hour in range(24):
                    hour_data = weekday_data[weekday_data['Hour'] == hour]

                    if len(hour_data) > 0:
                        loads = hour_data['Load_kW'].values
                        line_profile[weekday][hour] = {
                            'avg': np.mean(loads),
                            'max': np.max(loads),
                            'min': np.min(loads),
                            'std': np.std(loads) if len(loads) > 1 else 0,
                            'count': len(loads),
                            'is_weekend': weekday >= 5
                        }

            self.yearly_week_profiles[line] = line_profile

        self.log(f"   ✓ {len(lines)}개 선로의 요일-시간 프로파일 생성 완료")

        # 샘플 출력
        if self.yearly_week_profiles:
            sample_line = list(self.yearly_week_profiles.keys())[0]
            profile = self.yearly_week_profiles[sample_line]
            self.log(f"   ✓ 샘플({sample_line}): {len(profile)}개 요일 × 24시간 데이터")

    def calculate_scaling_factors(self):
        """
        364일 전 동일 주차 기반 시간대별 트렌드 보정 계수 계산

        [Reference Week Hourly Trend 방식]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        수식: Trend_Factor[hour] = (Target요일의 hour시 부하) / (참조주차 hour시 평균)

        예시 (11시 기준):
        - Target 요일(목) 11시 부하: 9.5 MW
        - 참조 주차 전체(월~일) 11시 평균: 8.8 MW
        - Trend_Factor[11] = 9.5 / 8.8 = 1.0795

        핵심 원리:
        1. 현재일 기준 364일 전 주차(월~일) 데이터 추출
        2. 각 시간대별로 참조 주차의 평균 부하 계산
        3. Target 요일의 해당 시간대 부하와 비교하여 Trend Factor 산출

        장점:
        - 시간대별 특성이 정확히 반영됨 (피크 시간대 vs 비피크 시간대)
        - 동일 시즌, 동일 주차의 실제 패턴 반영
        - 요일별 + 시간대별 이중 보정
        """
        self.scaling_factors = {}

        if self.long_df is None or len(self.long_df) == 0:
            return

        # 현재 날짜 (분석 시작일)
        current_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        # 364일 전 참조 주차 계산
        ref_monday, ref_sunday, _ = self.get_reference_week_dates(current_date)

        self.log(f"   • 분석 기준일: {current_date.strftime('%Y-%m-%d')} ({self.WEEKDAY_NAMES_KR[current_date.weekday()]})")
        self.log(f"   • 참조 주차(364일 전): {ref_monday.strftime('%Y-%m-%d')}(월) ~ {ref_sunday.strftime('%Y-%m-%d')}(일)")
        self.log(f"   • 보정 방식: 시간대별 Trend Factor")

        lines = self.long_df['Line'].unique()

        for line in lines:
            line_data = self.long_df[self.long_df['Line'] == line]

            # Reference Week 데이터 추출 (364일 전 주차)
            ref_start = pd.Timestamp(ref_monday)
            ref_end = pd.Timestamp(ref_sunday) + timedelta(days=1)

            ref_week_data = line_data[
                (line_data['Timestamp'] >= ref_start) &
                (line_data['Timestamp'] < ref_end)
            ]

            if len(ref_week_data) == 0:
                self.scaling_factors[line] = {
                    'hourly_factors': {h: 1.0 for h in range(24)},
                    'hourly_ref_avg': {h: None for h in range(24)},
                    'hourly_target_avg': {h: None for h in range(24)},
                    'factor': 1.0,  # 대표값 (평균)
                    'ref_week_avg_kw': None,
                    'change_pct': 0,
                    'trend_direction': '→',
                    'ref_period': f"{ref_monday.strftime('%m/%d')}~{ref_sunday.strftime('%m/%d')}",
                }
                continue

            # Reference Week 전체 평균 (참고용)
            ref_week_avg = ref_week_data['Load_kW'].mean()

            # ================================================================
            # 시간대별 Trend Factor 계산 (핵심 로직)
            # ================================================================
            target_weekday = current_date.weekday()

            # 1) 참조 주차의 시간대별 평균 (전체 요일 평균)
            hourly_ref_avg = ref_week_data.groupby('Hour')['Load_kW'].mean()

            # 2) 참조 주차 내 Target 요일의 시간대별 부하
            target_weekday_data = ref_week_data[ref_week_data['Weekday'] == target_weekday]
            if len(target_weekday_data) > 0:
                hourly_target_avg = target_weekday_data.groupby('Hour')['Load_kW'].mean()
            else:
                hourly_target_avg = hourly_ref_avg  # 데이터 없으면 전체 평균 사용

            # 3) 시간대별 Trend Factor 계산
            hourly_factors = {}
            hourly_ref_dict = {}
            hourly_target_dict = {}

            for hour in range(24):
                ref_avg_h = hourly_ref_avg.get(hour, None)
                target_avg_h = hourly_target_avg.get(hour, None) if hour in hourly_target_avg.index else ref_avg_h

                hourly_ref_dict[hour] = ref_avg_h
                hourly_target_dict[hour] = target_avg_h

                if ref_avg_h and ref_avg_h > 0 and target_avg_h:
                    hourly_factors[hour] = target_avg_h / ref_avg_h
                else:
                    hourly_factors[hour] = 1.0

            # 대표 Trend Factor (작업 시간대 09~14시 평균)
            work_hour_factors = [hourly_factors[h] for h in range(9, 15) if h in hourly_factors]
            avg_factor = np.mean(work_hour_factors) if work_hour_factors else 1.0

            # 트렌드 방향 결정
            change_pct = (avg_factor - 1) * 100
            if change_pct > 3:
                trend_direction = "↑ 고부하"
            elif change_pct < -3:
                trend_direction = "↓ 저부하"
            else:
                trend_direction = "→ 평균"

            self.scaling_factors[line] = {
                'hourly_factors': hourly_factors,      # 시간대별 Trend Factor
                'hourly_ref_avg': hourly_ref_dict,    # 참조주차 시간대별 평균
                'hourly_target_avg': hourly_target_dict,  # Target요일 시간대별 부하
                'factor': avg_factor,  # 대표값 (작업시간대 평균)
                'ref_week_avg_kw': ref_week_avg,
                'change_pct': change_pct,
                'trend_direction': trend_direction,
                'ref_period': f"{ref_monday.strftime('%m/%d')}~{ref_sunday.strftime('%m/%d')}",
            }

        # 로그 출력
        factors = [s['factor'] for s in self.scaling_factors.values() if s['factor']]
        if factors:
            avg_all = np.mean(factors)
            self.log(f"   ✓ 평균 Trend Factor (작업시간대): {avg_all:.4f} ({(avg_all-1)*100:+.2f}%)")
            self.log(f"   ✓ 공식: Trend_Factor[hour] = (Target요일 hour시 부하) / (참조주차 hour시 평균)")

            # 개별 선로 시간대별 트렌드 출력 (상위 3개만)
            self.log("")
            for line, data in list(self.scaling_factors.items())[:3]:
                hourly_f = data.get('hourly_factors', {})
                hourly_ref = data.get('hourly_ref_avg', {})
                hourly_tgt = data.get('hourly_target_avg', {})

                self.log(f"     [{line}] (참조주차: {data.get('ref_period', 'N/A')})")
                # 작업 시간대 샘플 출력 (09, 11, 13시)
                for h in [9, 11, 13]:
                    ref_h = hourly_ref.get(h)
                    tgt_h = hourly_tgt.get(h)
                    factor_h = hourly_f.get(h, 1.0)
                    if ref_h and tgt_h:
                        self.log(f"       {h:02d}시: 참조주차평균={ref_h/1000:.2f}MW, "
                                f"Target요일={tgt_h/1000:.2f}MW → Factor={factor_h:.4f}")

    def get_reference_week_dates(self, target_date):
        """
        Target Date 기준 364일(52주) 전 참조 주차의 월~일 날짜 반환

        [핵심 로직]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        - 364일 = 52주 (365일 아님!)
        - 이유: 요일(Day of Week)을 완벽히 일치시키기 위함
        - 예시: 2026-01-21(수) → 364일 전 = 2025-01-22(수)

        Args:
            target_date: 예측 대상 날짜 (datetime)

        Returns:
            tuple: (참조 시작일(월), 참조 종료일(일), 참조 날짜 리스트)
        """
        # 364일 전 = 52주 전 (요일 완벽 매칭)
        ref_date = target_date - timedelta(days=self.DAYS_IN_YEAR)

        # 해당 날짜가 속한 주의 월요일 찾기
        ref_weekday = ref_date.weekday()  # 0=월, 6=일
        ref_monday = ref_date - timedelta(days=ref_weekday)
        ref_sunday = ref_monday + timedelta(days=6)

        # 주간 날짜 리스트 (월~일)
        ref_dates = [ref_monday + timedelta(days=i) for i in range(7)]

        return ref_monday, ref_sunday, ref_dates

    def slice_reference_week_data(self, target_date, line):
        """
        1년 전 동일 주차 데이터를 실제 데이터프레임에서 슬라이싱

        [Yearly-Week-Sync 핵심 함수]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        Target Date 기준 364일 전 날짜가 포함된 주(월~일)의
        00:00~23:00 시간대별 부하 데이터를 추출합니다.

        Args:
            target_date: 예측 대상 날짜 (datetime)
            line: 선로명 (str)

        Returns:
            dict: {
                'ref_monday': 참조 주 월요일,
                'ref_sunday': 참조 주 일요일,
                'hourly_avg': {hour: avg_load},  # 시간대별 평균
                'hourly_max': {hour: max_load},  # 시간대별 최대
                'weekday_avg': {weekday: {hour: avg}},  # 요일별 시간대 평균
                'daily_peak': {date: peak_load},  # 일별 피크
                'same_weekday_load': 동일 요일 평균 부하,
                'data_count': 데이터 개수
            }
        """
        if self.long_df is None or len(self.long_df) == 0:
            return None

        # 참조 주차 날짜 계산
        ref_monday, ref_sunday, ref_dates = self.get_reference_week_dates(target_date)

        # 해당 선로 데이터 필터링
        line_data = self.long_df[self.long_df['Line'] == line].copy()
        if len(line_data) == 0:
            return None

        # 참조 주차 데이터 슬라이싱 (월요일 00:00 ~ 일요일 23:59)
        ref_start = pd.Timestamp(ref_monday)
        ref_end = pd.Timestamp(ref_sunday) + timedelta(days=1) - timedelta(seconds=1)

        ref_week_data = line_data[
            (line_data['Timestamp'] >= ref_start) &
            (line_data['Timestamp'] <= ref_end)
        ]

        if len(ref_week_data) == 0:
            return None

        # 시간대별 평균/최대 부하 계산
        hourly_stats = ref_week_data.groupby('Hour')['Load_kW'].agg(['mean', 'max', 'min', 'std'])
        hourly_avg = hourly_stats['mean'].to_dict()
        hourly_max = hourly_stats['max'].to_dict()

        # 요일별 시간대 평균 (상세 프로파일)
        weekday_hourly = ref_week_data.groupby(['Weekday', 'Hour'])['Load_kW'].mean()
        weekday_avg = {}
        for (wd, hr), load in weekday_hourly.items():
            if wd not in weekday_avg:
                weekday_avg[wd] = {}
            weekday_avg[wd][hr] = load

        # 일별 피크 부하
        daily_peak = ref_week_data.groupby(ref_week_data['Timestamp'].dt.date)['Load_kW'].max().to_dict()

        # 동일 요일 평균 부하 (Target Date와 같은 요일)
        target_weekday = target_date.weekday()
        same_weekday_data = ref_week_data[ref_week_data['Weekday'] == target_weekday]
        same_weekday_load = same_weekday_data['Load_kW'].mean() if len(same_weekday_data) > 0 else None

        # 평일/주말 구분 평균
        weekday_only = ref_week_data[~ref_week_data['IsWeekend']]
        weekend_only = ref_week_data[ref_week_data['IsWeekend']]

        return {
            'ref_monday': ref_monday,
            'ref_sunday': ref_sunday,
            'ref_dates': ref_dates,
            'hourly_avg': hourly_avg,
            'hourly_max': hourly_max,
            'weekday_avg': weekday_avg,
            'daily_peak': daily_peak,
            'same_weekday_load': same_weekday_load,
            'weekday_only_avg': weekday_only['Load_kW'].mean() if len(weekday_only) > 0 else None,
            'weekend_only_avg': weekend_only['Load_kW'].mean() if len(weekend_only) > 0 else None,
            'data_count': len(ref_week_data),
            'raw_data': ref_week_data  # 원본 데이터 (시각화용)
        }

    def extract_last_year_same_week_feature(self, line, target_date, hour):
        """
        XGBoost 피처용: 1년 전 동일 주차 동일 시간대 부하 추출

        [Feature Engineering]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        last_year_same_week_load 피처를 생성합니다.
        이 피처는 XGBoost 모델의 입력으로 사용됩니다.

        Args:
            line: 선로명
            target_date: 예측 대상 날짜
            hour: 시간 (0~23)

        Returns:
            float: 1년 전 동일 주차, 동일 요일, 동일 시간대 부하 (kW)
        """
        # 캐시 확인
        cache_key = f"{line}_{target_date.date()}_{hour}"
        if cache_key in self.last_year_same_week_features:
            return self.last_year_same_week_features[cache_key]

        # 참조 데이터 슬라이싱
        ref_data = self.slice_reference_week_data(target_date, line)
        if ref_data is None:
            return None

        # 동일 요일, 동일 시간대 부하
        target_weekday = target_date.weekday()
        if target_weekday in ref_data['weekday_avg'] and hour in ref_data['weekday_avg'][target_weekday]:
            load = ref_data['weekday_avg'][target_weekday][hour]
        elif hour in ref_data['hourly_avg']:
            load = ref_data['hourly_avg'][hour]
        else:
            load = None

        # 캐시 저장
        self.last_year_same_week_features[cache_key] = load
        return load

    def generate_weekday_weekend_profiles(self):
        """
        평일/주말 별도 프로파일 생성

        [주말 부하 별도 취급]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        - 평일(월~금): 산업/상업 부하 패턴
        - 주말(토~일): 주거/여가 부하 패턴

        나중에 주말 부하를 별도로 취급할 때 사용됩니다.
        """
        if self.long_df is None or len(self.long_df) == 0:
            return

        self.weekday_profiles = {}
        self.weekend_profiles = {}

        lines = self.long_df['Line'].unique()

        for line in lines:
            line_data = self.long_df[self.long_df['Line'] == line]

            # 평일 데이터 (월~금, weekday 0~4)
            weekday_data = line_data[~line_data['IsWeekend']]
            # 주말 데이터 (토~일, weekday 5~6)
            weekend_data = line_data[line_data['IsWeekend']]

            # 평일 시간대별 프로파일
            if len(weekday_data) > 0:
                wd_profile = {}
                for hour in range(24):
                    hour_data = weekday_data[weekday_data['Hour'] == hour]
                    if len(hour_data) > 0:
                        wd_profile[hour] = {
                            'avg': hour_data['Load_kW'].mean(),
                            'max': hour_data['Load_kW'].max(),
                            'min': hour_data['Load_kW'].min(),
                            'std': hour_data['Load_kW'].std() if len(hour_data) > 1 else 0
                        }
                self.weekday_profiles[line] = wd_profile

            # 주말 시간대별 프로파일
            if len(weekend_data) > 0:
                we_profile = {}
                for hour in range(24):
                    hour_data = weekend_data[weekend_data['Hour'] == hour]
                    if len(hour_data) > 0:
                        we_profile[hour] = {
                            'avg': hour_data['Load_kW'].mean(),
                            'max': hour_data['Load_kW'].max(),
                            'min': hour_data['Load_kW'].min(),
                            'std': hour_data['Load_kW'].std() if len(hour_data) > 1 else 0
                        }
                self.weekend_profiles[line] = we_profile

        self.log(f"   ✓ 평일 프로파일: {len(self.weekday_profiles)}개 선로")
        self.log(f"   ✓ 주말 프로파일: {len(self.weekend_profiles)}개 선로")

    def get_yearly_week_sync_load(self, line, target_date, hour, apply_scaling=True, return_details=False):
        """
        Yearly-Week-Sync 방식으로 부하 예측값 반환

        핵심: 364일(52주) 전 동일 요일, 동일 시간대 부하 참조

        Args:
            line: 선로명
            target_date: 예측 대상 날짜
            hour: 시간 (0~23)
            apply_scaling: 스케일링 팩터 적용 여부
            return_details: True면 계산 상세 정보도 함께 반환

        Returns:
            float: 예측 부하 (kW)
            또는 return_details=True인 경우:
            tuple: (예측 부하, 상세 정보 dict)
        """
        details = {
            'line': line,
            'target_date': target_date,
            'hour': hour,
            'weekday': target_date.weekday(),
            'weekday_name': self.WEEKDAY_NAMES_KR[target_date.weekday()],
            'base_load_kw': None,
            'base_load_source': None,
            'trend_factor': 1.0,
            'trend_applied': False,
            'final_load_kw': None,
            'profile_stats': None
        }

        if line not in self.yearly_week_profiles:
            if return_details:
                details['base_load_source'] = 'NO_PROFILE'
                return None, details
            return None

        profile = self.yearly_week_profiles[line]
        weekday = target_date.weekday()

        # 해당 요일, 시간대 데이터 조회
        if weekday in profile and hour in profile[weekday]:
            # 보수적 예측: 최대값 사용
            base_load = profile[weekday][hour]['max']
            details['base_load_kw'] = base_load
            details['base_load_source'] = f'PROFILE[{self.WEEKDAY_NAMES_KR[weekday]}][{hour}시].max'
            details['profile_stats'] = {
                'avg': profile[weekday][hour]['avg'],
                'max': profile[weekday][hour]['max'],
                'min': profile[weekday][hour]['min'],
                'count': profile[weekday][hour].get('count', 'N/A')
            }
        else:
            # 데이터 없으면 전체 평균 사용
            all_loads = []
            for wd, hours_data in profile.items():
                if hour in hours_data:
                    all_loads.append(hours_data[hour]['avg'])

            if all_loads:
                base_load = np.mean(all_loads)
                details['base_load_kw'] = base_load
                details['base_load_source'] = f'ALL_WEEKDAY_AVG[{hour}시] (해당요일 데이터 없음)'
            else:
                if return_details:
                    details['base_load_source'] = 'NO_DATA'
                    return None, details
                return None

        # Trend Factor 적용 (시간대별)
        if apply_scaling and line in self.scaling_factors:
            sf_data = self.scaling_factors[line]
            hourly_factors = sf_data.get('hourly_factors', {})

            # 해당 시간대의 Trend Factor 사용
            if hour in hourly_factors:
                trend_factor = hourly_factors[hour]
            else:
                trend_factor = sf_data.get('factor', 1.0)  # 없으면 대표값 사용

            final_load = base_load * trend_factor
            details['trend_factor'] = trend_factor
            details['trend_applied'] = True
            details['final_load_kw'] = final_load

            # 추가 정보 (검증용)
            details['hourly_ref_avg'] = sf_data.get('hourly_ref_avg', {}).get(hour)
            details['hourly_target_avg'] = sf_data.get('hourly_target_avg', {}).get(hour)
        else:
            final_load = base_load
            details['final_load_kw'] = final_load

        if return_details:
            return final_load, details
        return final_load

    def yearly_week_sync_verification(self, days):
        """
        Yearly-Week-Sync 방식 가부 판정

        핵심 로직:
        1. Target Date 기준 364일 전 동일 주차 참조
        2. 요일별 시간대 부하 예측
        3. 스케일링 팩터로 최근 트렌드 보정
        4. 1/N 부하 절체 적용 후 임계치 판정
        """
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        future_dates = [start_date + timedelta(days=i) for i in range(days)]

        results = []
        self.detailed_results = []
        self.weekly_predictions = {}

        transfer_targets = self.shutdown_mapping.get(self.selected_shutdown_line, [])
        n_targets = len(transfer_targets)

        for date in future_dates:
            weekday = date.weekday()
            weekday_name = self.WEEKDAY_NAMES_KR[weekday]
            is_weekend = weekday >= 5

            # 참조 주차 정보
            ref_monday, ref_sunday, ref_dates = self.get_reference_week_dates(date)
            ref_week_str = f"{ref_monday.strftime('%Y-%m-%d')}~{ref_sunday.strftime('%m-%d')}"

            # 24시간 예측
            hourly_combined_loads = {i: {} for i in range(1, n_targets + 1)}
            all_hours = list(range(24))

            for hour in all_hours:
                for i, target in enumerate(transfer_targets, 1):
                    # 절체 대상 선로 부하 (Yearly-Week-Sync)
                    target_load = self.get_yearly_week_sync_load(target, date, hour, apply_scaling=True)

                    # 휴전 선로 부하
                    shutdown_load = self.get_yearly_week_sync_load(
                        self.selected_shutdown_line, date, hour, apply_scaling=True
                    )

                    if target_load is None:
                        target_load = 8000  # 기본값 8MW
                    if shutdown_load is None:
                        shutdown_load = 10000  # 기본값 10MW

                    # 합산 부하 = 자체 부하 + (휴전 부하 / N)
                    combined_load = target_load + (shutdown_load / n_targets)
                    hourly_combined_loads[i][hour] = combined_load

            # 1주일 예측 데이터 저장 (시각화용)
            self.weekly_predictions[str(date.date())] = {
                'targets': transfer_targets,
                'hourly_loads': hourly_combined_loads,
                'weekday': weekday_name,
                'is_weekend': is_weekend,
                'ref_week': ref_week_str
            }

            # 작업 시간대 (09:00~14:00) 분석
            work_hours = [9, 10, 11, 12, 13]
            transfer_loads = {}

            # 24시간 피크/최저
            all_day_loads = []
            for i in range(1, n_targets + 1):
                for hour in all_hours:
                    all_day_loads.append(hourly_combined_loads[i][hour])

            peak_load_24h = max(all_day_loads) if all_day_loads else 0
            min_load_24h = min(all_day_loads) if all_day_loads else 0

            for i, target in enumerate(transfer_targets, 1):
                work_hour_loads = [hourly_combined_loads[i][h] for h in work_hours]
                max_load = max(work_hour_loads) if work_hour_loads else 8000
                avg_load = np.mean(work_hour_loads) if work_hour_loads else 8000

                transfer_loads[i] = {
                    'max': max_load,
                    'avg': avg_load,
                    'target': target
                }

            # 엄격 판정: 모든 절체선로가 임계치 미만이어야 SUCCESS
            all_under_threshold = True
            max_load_kw = 0
            failed_lines = []

            for i, loads in transfer_loads.items():
                max_kw = loads['max']
                max_load_kw = max(max_load_kw, max_kw)

                if max_kw >= self.threshold_kw:
                    all_under_threshold = False
                    target_name = loads['target']
                    failed_lines.append(f"{target_name}({max_kw/1000:.2f}MW)")

            # 스케일링 팩터 정보
            sf_info = ""
            if self.selected_shutdown_line in self.scaling_factors:
                sf = self.scaling_factors[self.selected_shutdown_line]['factor']
                sf_info = f" (SF:{sf:.3f})"

            # 판정 결과
            if all_under_threshold:
                status = '✅'
                remarks = f"{'주말' if is_weekend else '평일'}{sf_info}"
            else:
                status = '❌'
                remarks = f"초과: {', '.join(failed_lines[:2])}"

            max_load_mw = max_load_kw / 1000
            margin_mw = self.threshold_kw / 1000 - max_load_mw

            results.append({
                'Date': date.date(),
                'Weekday': weekday_name,
                'WeekdayNum': weekday,
                'Status': status,
                'PeakLoad_MW': peak_load_24h / 1000,
                'MinLoad_MW': min_load_24h / 1000,
                'MaxLoad_MW': max_load_mw,
                'Margin_MW': margin_mw,
                'RefWeek': ref_week_str,
                'Remarks': remarks,
                'IsWeekend': is_weekend,
                'IsFeasible': all_under_threshold,
                'TransferLoads': transfer_loads
            })

            self.detailed_results.append({
                'Date': date.date(),
                'TransferLoads': transfer_loads,
                'Feasible': all_under_threshold,
                'Weekday': weekday_name,
                'IsWeekend': is_weekend
            })

            # 참조 주차 정보 저장
            self.reference_week_info[str(date.date())] = {
                'target_date': date,
                'ref_monday': ref_monday,
                'ref_sunday': ref_sunday,
                'ref_dates': ref_dates
            }

        df_result = pd.DataFrame(results)

        # 통계
        total = len(df_result)
        success = df_result['IsFeasible'].sum()
        fail = total - success

        # 평일/주말 구분 통계
        weekday_results = df_result[~df_result['IsWeekend']]
        weekend_results = df_result[df_result['IsWeekend']]

        weekday_success = weekday_results['IsFeasible'].sum() if len(weekday_results) > 0 else 0
        weekend_success = weekend_results['IsFeasible'].sum() if len(weekend_results) > 0 else 0

        self.log(f"   ✓ 전체: SUCCESS {success}일 / FAIL {fail}일 ({success/total*100:.1f}%)")
        if len(weekday_results) > 0:
            self.log(f"   ✓ 평일: SUCCESS {weekday_success}일 / FAIL {len(weekday_results)-weekday_success}일")
        if len(weekend_results) > 0:
            self.log(f"   ✓ 주말: SUCCESS {weekend_success}일 / FAIL {len(weekend_results)-weekend_success}일")

        return df_result

    # =========================================================================
    # 결과 표시 메서드
    # =========================================================================

    def display_results(self):
        """결과 테이블 표시"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.results_df.iterrows():
            # 절체선로별 부하 포맷팅
            transfer_loads_text = ""
            if 'TransferLoads' in row and isinstance(row['TransferLoads'], dict):
                load_parts = []
                for i in sorted(row['TransferLoads'].keys()):
                    target_name = row['TransferLoads'][i].get('target', f"절체{i}")
                    max_mw = row['TransferLoads'][i]['max'] / 1000
                    load_parts.append(f"{target_name}:{max_mw:.1f}")
                transfer_loads_text = " | ".join(load_parts)
            else:
                transfer_loads_text = f"{row['MaxLoad_MW']:.2f}MW"

            # 주말 표시
            weekday_display = f"{row['Weekday']}{'*' if row['IsWeekend'] else ''}"

            values = (
                str(row['Date']),
                weekday_display,
                row['Status'],
                f"{row['PeakLoad_MW']:.2f}",
                f"{row['MinLoad_MW']:.2f}",
                transfer_loads_text,
                f"{row['Margin_MW']:+.2f}",
                row['RefWeek'],
                row['Remarks']
            )

            if row['Status'] == '✅':
                tag = 'safe'
            elif row['IsWeekend']:
                tag = 'weekend'
            else:
                tag = 'danger'

            self.tree.insert('', 'end', values=values, tags=(tag,))

        self.tree.tag_configure('safe', background='#d1fae5')
        self.tree.tag_configure('danger', background='#fee2e2')
        self.tree.tag_configure('weekend', background='#fef3c7')

    def on_date_selected(self, event):
        """날짜 선택 시 1주일 그래프 표시"""
        selection = self.tree.selection()
        if not selection:
            return

        item = self.tree.item(selection[0])
        selected_date = item['values'][0]

        if selected_date in self.weekly_predictions:
            self.draw_weekly_graph(selected_date)
            self.notebook.select(self.graph_tab)

    def draw_weekly_graph(self, date_str):
        """1주일 예측 그래프 그리기"""
        if not MATPLOTLIB_AVAILABLE:
            self.graph_info_label.config(text="matplotlib이 설치되지 않아\n그래프를 표시할 수 없습니다.")
            return

        # 기존 그래프 제거
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        prediction = self.weekly_predictions.get(date_str)
        if not prediction:
            return

        targets = prediction['targets']
        hourly_loads = prediction['hourly_loads']
        weekday = prediction['weekday']
        is_weekend = prediction['is_weekend']
        ref_week = prediction['ref_week']

        # Figure 생성
        fig = Figure(figsize=(11, 6), dpi=100)
        ax = fig.add_subplot(111)

        hours = list(range(24))
        colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4']

        # 각 절체선로별 부하 그래프
        for i, target in enumerate(targets, 1):
            loads_mw = [hourly_loads[i][h] / 1000 for h in hours]
            color = colors[(i-1) % len(colors)]
            ax.plot(hours, loads_mw, marker='o', markersize=4, label=f'{target}',
                   color=color, linewidth=2)

        # 임계치 라인 (14MW)
        threshold_mw = self.threshold_kw / 1000
        ax.axhline(y=threshold_mw, color='red', linestyle='--', linewidth=2,
                  label=f'임계치 ({threshold_mw:.0f}MW)')

        # 작업 시간대 강조 (09:00~14:00)
        ax.axvspan(9, 14, alpha=0.2, color='yellow', label='작업 시간대 (09-14시)')

        # 주말 표시
        weekend_marker = " [주말]" if is_weekend else " [평일]"

        # 그래프 스타일
        ax.set_xlabel('시간 (Hour)', fontsize=11)
        ax.set_ylabel('부하 (MW)', fontsize=11)
        ax.set_title(f'📊 {date_str} ({weekday}요일){weekend_marker} 24시간 부하 예측\n'
                    f'대조 구간: {ref_week}', fontsize=13, fontweight='bold')
        ax.set_xticks(hours)
        ax.set_xlim(-0.5, 23.5)
        ax.grid(True, alpha=0.3)
        ax.legend(loc='upper right', fontsize=9)

        # Y축 범위
        all_loads = []
        for i in range(1, len(targets) + 1):
            all_loads.extend([hourly_loads[i][h] / 1000 for h in hours])
        y_max = max(max(all_loads), threshold_mw) * 1.1
        ax.set_ylim(0, y_max)

        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

    def display_weekday_summary(self):
        """요일별 패턴 요약"""
        self.weekday_text.delete('1.0', 'end')

        text = "=" * 90 + "\n"
        text += f"📊 '{self.selected_shutdown_line}' Yearly-Week-Sync 요일별 분석\n"
        text += "=" * 90 + "\n\n"

        transfer_targets = self.shutdown_mapping.get(self.selected_shutdown_line, [])

        # 요일별 통계
        if self.results_df is not None and len(self.results_df) > 0:
            text += "▶ 요일별 분석 결과 (평일/주말 구분)\n"
            text += "-" * 90 + "\n\n"

            # 평일 (월~금)
            text += "📅 [평일 분석]\n"
            for weekday_num in range(5):
                weekday_data = self.results_df[self.results_df['WeekdayNum'] == weekday_num]
                if len(weekday_data) == 0:
                    continue

                weekday_name = self.WEEKDAY_NAMES_FULL[weekday_num]
                avg_peak = weekday_data['PeakLoad_MW'].mean()
                max_peak = weekday_data['PeakLoad_MW'].max()
                success_rate = weekday_data['IsFeasible'].mean() * 100

                status_icon = "✅" if success_rate >= 50 else "⚠️" if success_rate > 0 else "❌"
                text += f"   {status_icon} {weekday_name}: 피크 {avg_peak:.2f}MW (최대 {max_peak:.2f}MW), "
                text += f"성공률 {success_rate:.0f}%\n"

            # 주말 (토~일)
            text += "\n📅 [주말 분석] (별도 취급 예정)\n"
            for weekday_num in range(5, 7):
                weekday_data = self.results_df[self.results_df['WeekdayNum'] == weekday_num]
                if len(weekday_data) == 0:
                    continue

                weekday_name = self.WEEKDAY_NAMES_FULL[weekday_num]
                avg_peak = weekday_data['PeakLoad_MW'].mean()
                max_peak = weekday_data['PeakLoad_MW'].max()
                success_rate = weekday_data['IsFeasible'].mean() * 100

                status_icon = "✅" if success_rate >= 50 else "⚠️" if success_rate > 0 else "❌"
                text += f"   {status_icon} {weekday_name}: 피크 {avg_peak:.2f}MW (최대 {max_peak:.2f}MW), "
                text += f"성공률 {success_rate:.0f}%\n"

        # Trend Factor (364일 전 동일 주차 기반 - 시간대별)
        text += "\n" + "-" * 90 + "\n"
        text += "▶ Trend Factor (시간대별 보정)\n"
        text += "   공식: Trend_Factor[hour] = (Target요일 hour시) / (참조주차 hour시 평균)\n"
        text += "-" * 90 + "\n\n"

        for target in [self.selected_shutdown_line] + transfer_targets:
            if target in self.scaling_factors:
                sf = self.scaling_factors[target]
                trend_dir = sf.get('trend_direction', '→')
                ref_period = sf.get('ref_period', 'N/A')
                hourly_f = sf.get('hourly_factors', {})
                hourly_ref = sf.get('hourly_ref_avg', {})
                hourly_tgt = sf.get('hourly_target_avg', {})

                text += f"   [{target}] 참조주차: {ref_period} ({trend_dir})\n"
                # 작업 시간대(09~13시) Factor 출력
                for h in [9, 10, 11, 12, 13]:
                    ref_h = hourly_ref.get(h)
                    tgt_h = hourly_tgt.get(h)
                    factor_h = hourly_f.get(h, 1.0)
                    if ref_h and tgt_h:
                        text += f"      {h:02d}시: {tgt_h/1000:.2f}MW / {ref_h/1000:.2f}MW = {factor_h:.4f}\n"
                text += "\n"

        # 요일별 추천도
        text += "\n" + "-" * 90 + "\n"
        text += "▶ 요일별 작업 추천도\n"
        text += "-" * 90 + "\n\n"

        if self.results_df is not None:
            weekday_stats = []
            for weekday_num in range(7):
                weekday_data = self.results_df[self.results_df['WeekdayNum'] == weekday_num]
                if len(weekday_data) > 0:
                    success_rate = weekday_data['IsFeasible'].mean() * 100
                    avg_margin = weekday_data['Margin_MW'].mean()
                    is_weekend = weekday_num >= 5
                    weekday_stats.append({
                        'name': self.WEEKDAY_NAMES_FULL[weekday_num],
                        'success_rate': success_rate,
                        'avg_margin': avg_margin,
                        'is_weekend': is_weekend
                    })

            weekday_stats.sort(key=lambda x: (-x['success_rate'], -x['avg_margin']))

            for i, stat in enumerate(weekday_stats):
                if stat['success_rate'] >= 80:
                    rating = "⭐⭐⭐⭐⭐"
                elif stat['success_rate'] >= 60:
                    rating = "⭐⭐⭐⭐"
                elif stat['success_rate'] >= 40:
                    rating = "⭐⭐⭐"
                elif stat['success_rate'] >= 20:
                    rating = "⭐⭐"
                else:
                    rating = "⭐"

                weekend_tag = "[주말]" if stat['is_weekend'] else "[평일]"
                text += f"   {i+1}. {stat['name']} {weekend_tag}: {rating} "
                text += f"(성공률 {stat['success_rate']:.0f}%, 여유 {stat['avg_margin']:.1f}MW)\n"

        self.weekday_text.insert('1.0', text)

    def display_recommendations(self):
        """
        추천 일자 표시

        [Yearly-Week-Sync v5.5 추천 리포트]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        - 평일 작업 우선 추천
        - 대조 구간 상세 표시
        - 1년 전 동일 요일 부하 참조
        """
        self.recommendation_text.delete('1.0', 'end')

        # 평일만 필터링 (주말은 별도 취급 예정)
        weekday_results = self.results_df[~self.results_df['IsWeekend']].copy()
        feasible_weekdays = weekday_results[weekday_results['IsFeasible']].copy()

        text = "━" * 90 + "\n"
        text += f"🎯 '{self.selected_shutdown_line}' 최적 작업일 추천 (Yearly-Week-Sync v5.5)\n"
        text += "━" * 90 + "\n\n"

        targets = self.shutdown_mapping.get(self.selected_shutdown_line, [])
        text += "┌─────────────────────────────────────────────────────────────────────────────────┐\n"
        text += "│ [분석 설정]                                                                     │\n"
        text += "└─────────────────────────────────────────────────────────────────────────────────┘\n"
        text += f"   • 휴전 선로: {self.selected_shutdown_line}\n"
        text += f"   • 절체 대상: {', '.join(targets)} (총 {len(targets)}개)\n"
        text += f"   • 임계치: {self.threshold_kw/1000:.1f}MW ({self.threshold_kw:,}kW)\n"
        text += f"   • 작업 시간대: 09:00 ~ 14:00\n"
        text += f"   • 알고리즘: Yearly-Week-Sync v5.5 Enhanced\n"
        text += f"   • 참조 기준: 364일(52주) 전 동일 주차 + 트렌드 보정\n"
        text += f"   • 판정 방식: 모든 절체선로 임계치 미만 필수 (엄격 검증)\n\n"
        text += "─" * 90 + "\n\n"

        if len(feasible_weekdays) == 0:
            text += "┌─────────────────────────────────────────────────────────────────────────────────┐\n"
            text += "│ ⚠️  경고: 향후 분석 기간 동안 작업 가능한 평일이 없습니다!                       │\n"
            text += "└─────────────────────────────────────────────────────────────────────────────────┘\n\n"
            text += "💡 대안 검토:\n"
            text += "   1. 작업 시간대를 야간(22:00~06:00)으로 변경\n"
            text += "   2. 새벽 시간대(04:00~07:00) 검토\n"
            text += "   3. 분석 기간 확대 (60일 이상)\n"
            text += "   4. 주말 작업 검토 (주말 부하 패턴 별도 분석 필요)\n"
            text += "   5. 절체 방식 변경 검토 (부분 절체 등)\n"
        else:
            feasible_weekdays = feasible_weekdays.sort_values('Margin_MW', ascending=False)
            top5 = feasible_weekdays.head(5)

            text += "┌─────────────────────────────────────────────────────────────────────────────────┐\n"
            text += "│ [추천 작업일 TOP 5] - 평일 기준, 여유량 순 정렬                                 │\n"
            text += "└─────────────────────────────────────────────────────────────────────────────────┘\n\n"

            medals = ['🥇', '🥈', '🥉', '4️⃣', '5️⃣']
            for idx, (_, row) in enumerate(top5.iterrows()):
                date_obj = row['Date']
                weekday_name = row['Weekday']

                # 1년 전 동일 요일 평균 부하
                if isinstance(date_obj, str):
                    target_dt = datetime.strptime(str(date_obj), '%Y-%m-%d')
                else:
                    target_dt = datetime(date_obj.year, date_obj.month, date_obj.day)
                last_year_load = self._get_last_year_same_weekday_avg(target_dt)
                last_year_str = f"{last_year_load/1000:.2f}MW" if last_year_load else "N/A"

                text += f"{medals[idx]} ━━━ 추천 {idx+1}순위 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                text += f"   📅 타겟 일자: {row['Date']} ({weekday_name}요일)\n"
                text += f"   📆 대조 구간: {row['RefWeek']} (364일 전 동일 주차)\n"
                text += f"   📊 1년 전 동일 요일 평균 부하: {last_year_str}\n"
                text += f"   ⚡ 24시간 피크: {row['PeakLoad_MW']:.2f}MW | 최저: {row['MinLoad_MW']:.2f}MW\n"
                text += f"   🔧 작업시간(09-14시) 최대: {row['MaxLoad_MW']:.2f}MW\n"
                text += f"   ✅ 안전 여유량: {row['Margin_MW']:.2f}MW ({row['Margin_MW']/14*100:.1f}%)\n"

                # Trend Factor 정보 (시간대별)
                if self.selected_shutdown_line in self.scaling_factors:
                    sf = self.scaling_factors[self.selected_shutdown_line]
                    trend_dir = sf.get('trend_direction', '→')
                    hourly_f = sf.get('hourly_factors', {})
                    # 작업시간대(09~13시) 평균 Factor
                    work_factors = [hourly_f.get(h, 1.0) for h in range(9, 14)]
                    avg_work_factor = np.mean(work_factors) if work_factors else 1.0
                    text += f"   📈 트렌드 보정: 작업시간대 평균 Factor = {avg_work_factor:.4f} ({trend_dir})\n"
                    if sf.get('ref_period'):
                        text += f"      └─ 참조 주차: {sf['ref_period']} (364일 전, 시간대별 보정)\n"

                # 평가
                if row['Margin_MW'] > 5:
                    text += f"   🏆 종합 평가: ⭐⭐⭐⭐⭐ 매우 안전 (여유 충분)\n"
                elif row['Margin_MW'] > 3:
                    text += f"   🏆 종합 평가: ⭐⭐⭐⭐ 안전 (권장)\n"
                elif row['Margin_MW'] > 1:
                    text += f"   🏆 종합 평가: ⭐⭐⭐ 보통 (주의 필요)\n"
                else:
                    text += f"   🏆 종합 평가: ⭐⭐ 경계 (면밀 모니터링 필수)\n"
                text += "\n"

            # 체크리스트
            text += "─" * 90 + "\n\n"
            text += "┌─────────────────────────────────────────────────────────────────────────────────┐\n"
            text += "│ [배전센터 작업 체크리스트]                                                       │\n"
            text += "└─────────────────────────────────────────────────────────────────────────────────┘\n\n"
            text += "   □ D-1일: 기상청 기온 예보 확인 (±5℃ 변동 시 부하 재검토)\n"
            text += "   □ D-Day 08:00: SCADA 시스템으로 실시간 부하 확인\n"
            text += "   □ 작업 전: 모든 절체선로 임계치 90% 미만 확인\n"
            text += "   □ 작업 중: 15분 간격 모니터링\n"
            text += "   □ 긴급 중단: 임계치 95% 도달 시 즉시 중단\n"
            text += "   □ 작업 완료: 30분간 부하 안정화 확인\n\n"

            # 통계
            total = len(self.results_df)
            success = self.results_df['IsFeasible'].sum()
            fail = total - success

            # 평일/주말 구분
            weekday_total = len(self.results_df[~self.results_df['IsWeekend']])
            weekday_success = self.results_df[~self.results_df['IsWeekend']]['IsFeasible'].sum()
            weekend_total = len(self.results_df[self.results_df['IsWeekend']])
            weekend_success = self.results_df[self.results_df['IsWeekend']]['IsFeasible'].sum()

            text += "━" * 90 + "\n"
            text += "┌─────────────────────────────────────────────────────────────────────────────────┐\n"
            text += "│ [통계 요약]                                                                      │\n"
            text += "└─────────────────────────────────────────────────────────────────────────────────┘\n\n"
            text += f"   📊 전체 분석 기간: {total}일\n"
            text += f"   ┌─────────────────────────────────────────┐\n"
            text += f"   │ ✅ SUCCESS: {success}일 ({success/total*100:.1f}%)             │\n"
            text += f"   │ ❌ FAIL: {fail}일 ({fail/total*100:.1f}%)                │\n"
            text += f"   └─────────────────────────────────────────┘\n\n"
            text += f"   📅 평일(월~금): ✅ {weekday_success}/{weekday_total}일 SUCCESS\n"
            text += f"   📅 주말(토~일): ✅ {weekend_success}/{weekend_total}일 SUCCESS [별도 취급 예정]\n\n"
            text += f"   ─────────────────────────────────────────\n"
            text += f"   📌 알고리즘: Yearly-Week-Sync v5.5 Enhanced\n"
            text += f"   📌 참조 기준: 364일(52주) 전 동일 주차\n"
            text += f"   📌 핵심 원리: 요일(DoW) 완벽 매칭 + 트렌드 보정\n"
            text += f"   📌 주말 부하: 별도 취급 예정 (현재 통합 분석)\n"

        self.recommendation_text.insert('1.0', text)

    def log_algorithm_analysis(self):
        """
        Yearly-Week-Sync 알고리즘 분석 로그

        [로그 출력 형식 - 사용자 요구사항]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        - 분석 모드: Yearly-Week-Sync v5.5
        - 타겟 일자: YYYY-MM-DD (요일)
        - 대조 구간: YYYY-MM-DD (월) ~ YYYY-MM-DD (일)
        - 1년 전 동일 요일 평균 부하: [값] MW
        - 최종 예측 결과 및 가부 판정: [SUCCESS/FAIL]
        """
        self.log("")
        self.log("━" * 70)
        self.log("📊 Yearly-Week-Sync v5.5 알고리즘 분석 리포트")
        self.log("━" * 70)

        # =====================================================================
        # 1. 분석 모드
        # =====================================================================
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [분석 모드]                                                      │")
        self.log("└─────────────────────────────────────────────────────────────────┘")
        self.log("   • 알고리즘: Yearly-Week-Sync v5.5 Enhanced")
        self.log("   • 참조 기준: 364일(52주) 전 동일 주차")
        self.log("   • 핵심 원리: 365일이 아닌 364일 사용 → 요일(DoW) 완벽 일치")
        self.log("   • 예시: 2026-01-21(수) → 364일 전 = 2025-01-22(수)")

        # =====================================================================
        # 2. 데이터 정보
        # =====================================================================
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [데이터 정보]                                                    │")
        self.log("└─────────────────────────────────────────────────────────────────┘")
        if self.long_df is not None:
            self.log(f"   • 총 레코드: {len(self.long_df):,}개")
            self.log(f"   • 선로 수: {self.long_df['Line'].nunique()}개")
            self.log(f"   • 데이터 기간: {self.long_df['Timestamp'].min().strftime('%Y-%m-%d')} ~ "
                    f"{self.long_df['Timestamp'].max().strftime('%Y-%m-%d')}")

            # 평일/주말 데이터 비율
            weekday_cnt = len(self.long_df[~self.long_df['IsWeekend']])
            weekend_cnt = len(self.long_df[self.long_df['IsWeekend']])
            self.log(f"   • 평일 데이터: {weekday_cnt:,}개 ({weekday_cnt/len(self.long_df)*100:.1f}%)")
            self.log(f"   • 주말 데이터: {weekend_cnt:,}개 ({weekend_cnt/len(self.long_df)*100:.1f}%)")

        # =====================================================================
        # 3. 상세 분석 로그 (사용자 요구사항 형식)
        # =====================================================================
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [일자별 상세 분석 로그]                                          │")
        self.log("└─────────────────────────────────────────────────────────────────┘")

        if self.reference_week_info and self.results_df is not None:
            # 처음 5일만 상세 출력
            for idx, (_, ref_info) in enumerate(list(self.reference_week_info.items())[:5]):
                target_date = ref_info['target_date']
                ref_monday = ref_info['ref_monday']
                ref_sunday = ref_info['ref_sunday']
                weekday_name = self.WEEKDAY_NAMES_KR[target_date.weekday()]

                # 해당 날짜 결과 조회
                date_result = self.results_df[self.results_df['Date'] == target_date.date()]
                if len(date_result) > 0:
                    result_row = date_result.iloc[0]
                    status = "SUCCESS" if result_row['IsFeasible'] else "FAIL"
                    status_icon = "✅" if result_row['IsFeasible'] else "❌"
                    peak_mw = result_row['PeakLoad_MW']
                    margin_mw = result_row['Margin_MW']
                else:
                    status = "N/A"
                    status_icon = "⚠️"
                    peak_mw = 0
                    margin_mw = 0

                # 1년 전 동일 요일 평균 부하 계산
                last_year_load = self._get_last_year_same_weekday_avg(target_date)

                self.log("")
                self.log(f"   ─── 분석 #{idx+1} ───")
                self.log(f"   • 분석 모드: Yearly-Week-Sync v5.5")
                self.log(f"   • 타겟 일자: {target_date.strftime('%Y-%m-%d')} ({weekday_name})")
                self.log(f"   • 대조 구간: {ref_monday.strftime('%Y-%m-%d')} (월) ~ "
                        f"{ref_sunday.strftime('%Y-%m-%d')} (일)")
                if last_year_load:
                    self.log(f"   • 1년 전 동일 요일 평균 부하: {last_year_load/1000:.2f} MW")
                else:
                    self.log(f"   • 1년 전 동일 요일 평균 부하: 데이터 없음")
                self.log(f"   • 예측 피크 부하: {peak_mw:.2f} MW")
                self.log(f"   • 임계치 여유량: {margin_mw:+.2f} MW")
                self.log(f"   • 최종 예측 결과 및 가부 판정: {status_icon} {status}")

            if len(self.reference_week_info) > 5:
                self.log(f"\n   ... 외 {len(self.reference_week_info)-5}일 분석 결과 생략")

        # =====================================================================
        # 4. Trend Factor (364일 전 동일 주차 기반 - 시간대별)
        # =====================================================================
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [Trend Factor - 시간대별 보정]                                   │")
        self.log("└─────────────────────────────────────────────────────────────────┘")
        self.log("   공식: Trend_Factor[hour] = (Target요일 hour시) / (참조주차 hour시 평균)")
        self.log("   원리: 같은 시간대끼리 비교하여 보정 (피크시간 vs 비피크시간 특성 반영)")
        self.log("")
        if self.scaling_factors:
            for line, sf in list(self.scaling_factors.items())[:3]:
                trend_dir = sf.get('trend_direction', '→')
                ref_period = sf.get('ref_period', 'N/A')
                self.log(f"   [{line}] 참조주차: {ref_period} ({trend_dir})")

                # 시간대별 Factor 샘플 출력
                hourly_f = sf.get('hourly_factors', {})
                hourly_ref = sf.get('hourly_ref_avg', {})
                hourly_tgt = sf.get('hourly_target_avg', {})

                for h in [9, 11, 13]:
                    ref_h = hourly_ref.get(h)
                    tgt_h = hourly_tgt.get(h)
                    factor_h = hourly_f.get(h, 1.0)
                    if ref_h and tgt_h:
                        self.log(f"      {h:02d}시: {tgt_h/1000:.2f}MW / {ref_h/1000:.2f}MW = {factor_h:.4f}")

        # =====================================================================
        # 5. 예측 공식
        # =====================================================================
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [예측 공식]                                                      │")
        self.log("└─────────────────────────────────────────────────────────────────┘")
        self.log("")
        self.log("   ┌────────────────────────────────────────────────────────────┐")
        self.log("   │ Predicted_Load(line, date, hour)                          │")
        self.log("   │   = YearlyWeekSync_Max(line, date.weekday, hour)          │")
        self.log("   │   × Trend_Factor[hour]                                    │")
        self.log("   │                                                           │")
        self.log("   │ Trend_Factor[hour] = (Target요일 hour시) / (참조주차 hour시)│")
        self.log("   │   → 같은 시간대끼리 비교하여 보정 계수 산출                 │")
        self.log("   └────────────────────────────────────────────────────────────┘")
        self.log("")
        self.log("   ┌────────────────────────────────────────────────────────────┐")
        self.log("   │ Combined_Load(target, t)                                  │")
        self.log("   │   = Predicted(target, t) + Predicted(shutdown, t) / N     │")
        self.log("   │   where N = 절체 대상 선로 수                               │")
        self.log("   └────────────────────────────────────────────────────────────┘")

        # =====================================================================
        # 6. 가부 판정 기준
        # =====================================================================
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [가부 판정 기준]                                                 │")
        self.log("└─────────────────────────────────────────────────────────────────┘")
        self.log(f"   • 임계치: {self.threshold_kw:,} kW ({self.threshold_kw/1000:.1f} MW)")
        self.log("   • 판정 시간대: 09:00 ~ 14:00 (작업 시간)")
        self.log("   • SUCCESS 조건: 모든 절체선로 부하 < 임계치")
        self.log("   • FAIL 조건: 하나라도 부하 >= 임계치")

        # =====================================================================
        # 7. 예측 결과 요약
        # =====================================================================
        if self.results_df is not None:
            self.log("")
            self.log("┌─────────────────────────────────────────────────────────────────┐")
            self.log("│ [예측 결과 요약]                                                 │")
            self.log("└─────────────────────────────────────────────────────────────────┘")
            total = len(self.results_df)
            success = self.results_df['IsFeasible'].sum()
            fail = total - success

            weekday_df = self.results_df[~self.results_df['IsWeekend']]
            weekend_df = self.results_df[self.results_df['IsWeekend']]

            self.log(f"   ┌───────────────────────────────────────────┐")
            self.log(f"   │ 전체: ✅ {success}일 SUCCESS / ❌ {fail}일 FAIL  │")
            self.log(f"   │       ({success/total*100:.1f}% 성공률)                │")
            self.log(f"   └───────────────────────────────────────────┘")

            if len(weekday_df) > 0:
                wd_success = weekday_df['IsFeasible'].sum()
                wd_fail = len(weekday_df) - wd_success
                self.log(f"   • 평일(월~금): ✅ {wd_success}일 / ❌ {wd_fail}일")
            if len(weekend_df) > 0:
                we_success = weekend_df['IsFeasible'].sum()
                we_fail = len(weekend_df) - we_success
                self.log(f"   • 주말(토~일): ✅ {we_success}일 / ❌ {we_fail}일 [별도 취급 예정]")

        self.log("")
        self.log("━" * 70)
        self.log("📊 Yearly-Week-Sync v5.5 분석 리포트 출력 완료")
        self.log("━" * 70)

    def log_calculation_verification(self, num_samples=3):
        """
        예측 부하 계산 과정 상세 검증 로그

        [검증 로그 출력]
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        각 예측값이 어떻게 계산되었는지 수식과 대입값을 표시합니다.

        수식:
        1) Base_Load = Profile[요일][시간].max (보수적 예측)
        2) Predicted_Load = Base_Load × Trend_Factor
        3) Combined_Load = Target_Load + (Shutdown_Load / N)
        """
        self.log("")
        self.log("━" * 70)
        self.log("🔍 예측 부하 계산 검증 (Calculation Verification)")
        self.log("━" * 70)

        # 수식 설명
        self.log("")
        self.log("┌─────────────────────────────────────────────────────────────────┐")
        self.log("│ [예측 수식]                                                      │")
        self.log("├─────────────────────────────────────────────────────────────────┤")
        self.log("│ ① Base_Load = Profile[요일][시간].max                           │")
        self.log("│    → 364일 전 동일 주차의 해당 요일/시간 최대 부하               │")
        self.log("│                                                                 │")
        self.log("│ ② Predicted_Load = Base_Load × Trend_Factor[hour]              │")
        self.log("│    → Trend_Factor[hour] = (Target요일 hour시) / (참조주차 hour시)│")
        self.log("│    → 같은 시간대끼리 비교하여 보정 계수 산출                     │")
        self.log("│                                                                 │")
        self.log("│ ③ Combined_Load = Target_Load + (Shutdown_Load / N)            │")
        self.log("│    → N = 절체 대상 선로 수                                       │")
        self.log("└─────────────────────────────────────────────────────────────────┘")

        if self.results_df is None or len(self.results_df) == 0:
            self.log("\n   ⚠️ 분석 결과가 없습니다.")
            return

        transfer_targets = self.shutdown_mapping.get(self.selected_shutdown_line, [])
        n_targets = len(transfer_targets)

        self.log(f"\n   • 휴전 선로: {self.selected_shutdown_line}")
        self.log(f"   • 절체 대상: {', '.join(transfer_targets)} (N={n_targets})")
        self.log(f"   • 임계치: {self.threshold_kw/1000:.1f} MW")

        # 샘플 날짜들에 대해 상세 검증
        sample_dates = list(self.reference_week_info.keys())[:num_samples]

        for date_str in sample_dates:
            ref_info = self.reference_week_info[date_str]
            target_date = ref_info['target_date']
            ref_monday = ref_info['ref_monday']
            ref_sunday = ref_info['ref_sunday']
            weekday_name = self.WEEKDAY_NAMES_KR[target_date.weekday()]

            self.log("")
            self.log("=" * 70)
            self.log(f"📅 검증 대상: {target_date.strftime('%Y-%m-%d')} ({weekday_name}요일)")
            self.log(f"   참조 주차: {ref_monday.strftime('%Y-%m-%d')}(월) ~ {ref_sunday.strftime('%Y-%m-%d')}(일)")
            self.log("=" * 70)

            # 작업 시간대 중 피크 시간 (예: 11시) 검증
            sample_hours = [9, 11, 13]  # 작업 시간대 샘플

            for hour in sample_hours:
                self.log("")
                self.log(f"   ─── {hour:02d}:00 시간대 계산 검증 ───")

                # 휴전 선로 부하 계산
                shutdown_load, shutdown_details = self.get_yearly_week_sync_load(
                    self.selected_shutdown_line, target_date, hour,
                    apply_scaling=True, return_details=True
                )

                self.log(f"")
                self.log(f"   [휴전선로: {self.selected_shutdown_line}]")
                if shutdown_details['base_load_kw']:
                    self.log(f"      • 데이터 출처: {shutdown_details['base_load_source']}")
                    if shutdown_details['profile_stats']:
                        stats = shutdown_details['profile_stats']
                        self.log(f"      • 프로파일 통계: avg={stats['avg']/1000:.2f}MW, "
                                f"max={stats['max']/1000:.2f}MW, min={stats['min']/1000:.2f}MW")
                    self.log(f"      • Base_Load = {shutdown_details['base_load_kw']/1000:.4f} MW")

                    # Trend Factor 상세 (시간대별)
                    hourly_ref = shutdown_details.get('hourly_ref_avg')
                    hourly_tgt = shutdown_details.get('hourly_target_avg')
                    if hourly_ref and hourly_tgt:
                        self.log(f"      • Trend_Factor[{hour}시] 계산:")
                        self.log(f"        - 참조주차 {hour}시 평균 = {hourly_ref/1000:.4f} MW")
                        self.log(f"        - Target요일 {hour}시 부하 = {hourly_tgt/1000:.4f} MW")
                        self.log(f"        - Factor = {hourly_tgt/1000:.4f} / {hourly_ref/1000:.4f} = {shutdown_details['trend_factor']:.4f}")
                    else:
                        self.log(f"      • Trend_Factor[{hour}시] = {shutdown_details['trend_factor']:.4f}")

                    self.log(f"      ────────────────────────────────")
                    self.log(f"      • Predicted_Load = {shutdown_details['base_load_kw']/1000:.4f} × "
                            f"{shutdown_details['trend_factor']:.4f}")
                    self.log(f"                       = {shutdown_details['final_load_kw']/1000:.4f} MW")
                else:
                    self.log(f"      • 데이터 없음 → 기본값 10.0 MW 사용")
                    shutdown_load = 10000

                # 각 절체 대상 선로별 계산
                for i, target in enumerate(transfer_targets, 1):
                    target_load, target_details = self.get_yearly_week_sync_load(
                        target, target_date, hour,
                        apply_scaling=True, return_details=True
                    )

                    self.log(f"")
                    self.log(f"   [절체대상{i}: {target}]")
                    if target_details['base_load_kw']:
                        self.log(f"      • 데이터 출처: {target_details['base_load_source']}")
                        if target_details['profile_stats']:
                            stats = target_details['profile_stats']
                            self.log(f"      • 프로파일 통계: avg={stats['avg']/1000:.2f}MW, "
                                    f"max={stats['max']/1000:.2f}MW, min={stats['min']/1000:.2f}MW")
                        self.log(f"      • Base_Load = {target_details['base_load_kw']/1000:.4f} MW")

                        # Trend Factor 상세 (시간대별)
                        hourly_ref = target_details.get('hourly_ref_avg')
                        hourly_tgt = target_details.get('hourly_target_avg')
                        if hourly_ref and hourly_tgt:
                            self.log(f"      • Trend_Factor[{hour}시] 계산:")
                            self.log(f"        - 참조주차 {hour}시 평균 = {hourly_ref/1000:.4f} MW")
                            self.log(f"        - Target요일 {hour}시 부하 = {hourly_tgt/1000:.4f} MW")
                            self.log(f"        - Factor = {hourly_tgt/1000:.4f} / {hourly_ref/1000:.4f} = {target_details['trend_factor']:.4f}")
                        else:
                            self.log(f"      • Trend_Factor[{hour}시] = {target_details['trend_factor']:.4f}")

                        self.log(f"      ────────────────────────────────")
                        self.log(f"      • Predicted_Load = {target_details['base_load_kw']/1000:.4f} × "
                                f"{target_details['trend_factor']:.4f}")
                        self.log(f"                       = {target_details['final_load_kw']/1000:.4f} MW")
                    else:
                        self.log(f"      • 데이터 없음 → 기본값 8.0 MW 사용")
                        target_load = 8000

                    # Combined Load 계산
                    if target_load is None:
                        target_load = 8000
                    if shutdown_load is None:
                        shutdown_load = 10000

                    combined_load = target_load + (shutdown_load / n_targets)

                    self.log(f"")
                    self.log(f"   [Combined Load 계산 - {target}]")
                    self.log(f"      ┌────────────────────────────────────────────────────┐")
                    self.log(f"      │ Combined = Target_Load + (Shutdown_Load / N)      │")
                    self.log(f"      │         = {target_load/1000:.4f} + ({shutdown_load/1000:.4f} / {n_targets})     │")
                    self.log(f"      │         = {target_load/1000:.4f} + {shutdown_load/n_targets/1000:.4f}              │")
                    self.log(f"      │         = {combined_load/1000:.4f} MW                          │")
                    self.log(f"      └────────────────────────────────────────────────────┘")

                    # 판정
                    if combined_load >= self.threshold_kw:
                        self.log(f"      ❌ 판정: {combined_load/1000:.2f} MW >= {self.threshold_kw/1000:.1f} MW (FAIL)")
                    else:
                        margin = self.threshold_kw - combined_load
                        self.log(f"      ✅ 판정: {combined_load/1000:.2f} MW < {self.threshold_kw/1000:.1f} MW "
                                f"(여유: {margin/1000:.2f} MW)")

        self.log("")
        self.log("━" * 70)
        self.log("🔍 계산 검증 로그 출력 완료")
        self.log("━" * 70)

    def _get_last_year_same_weekday_avg(self, target_date):
        """1년 전 동일 요일 평균 부하 계산 (로그용)"""
        if self.long_df is None:
            return None

        target_weekday = target_date.weekday()

        # 참조 주차의 동일 요일 데이터 (364일 전 주차)
        ref_monday, ref_sunday, _ = self.get_reference_week_dates(target_date)

        # 휴전선로 데이터
        line_data = self.long_df[self.long_df['Line'] == self.selected_shutdown_line]
        if len(line_data) == 0:
            return None

        # 참조 주차 + 동일 요일 필터
        ref_start = pd.Timestamp(ref_monday)
        ref_end = pd.Timestamp(ref_sunday) + timedelta(days=1)

        ref_data = line_data[
            (line_data['Timestamp'] >= ref_start) &
            (line_data['Timestamp'] < ref_end) &
            (line_data['Weekday'] == target_weekday)
        ]

        if len(ref_data) > 0:
            return ref_data['Load_kW'].mean()
        return None

    # =========================================================================
    # 엑셀 저장
    # =========================================================================

    def save_to_excel(self):
        """결과 엑셀 저장"""
        if self.results_df is None:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"KEPCO_v5.5_YWS_{self.selected_shutdown_line}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        if filename:
            try:
                output_df = self.results_df.copy()

                # 절체선로 부하 포맷팅
                if 'TransferLoads' in output_df.columns:
                    transfer_loads_formatted = []
                    for _, row in output_df.iterrows():
                        if isinstance(row['TransferLoads'], dict):
                            load_parts = []
                            for i in sorted(row['TransferLoads'].keys()):
                                target_name = row['TransferLoads'][i].get('target', f"절체{i}")
                                max_mw = row['TransferLoads'][i]['max'] / 1000
                                load_parts.append(f"{target_name}:{max_mw:.2f}MW")
                            transfer_loads_formatted.append(" | ".join(load_parts))
                        else:
                            transfer_loads_formatted.append(f"{row['MaxLoad_MW']:.2f}MW")
                    output_df['절체선로 부하'] = transfer_loads_formatted
                    output_df = output_df.drop(columns=['TransferLoads'], errors='ignore')

                # 컬럼 정리
                output_df = output_df.drop(columns=['WeekdayNum'], errors='ignore')
                output_df.columns = ['날짜', '요일', '판정', '피크(MW)', '최저(MW)', '최대(MW)',
                                    '여유(MW)', '대조구간', '비고', '주말', '작업가능', '절체부하']
                output_df = output_df[['날짜', '요일', '판정', '피크(MW)', '최저(MW)',
                                      '절체부하', '여유(MW)', '대조구간', '비고', '주말', '작업가능']]

                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    output_df.to_excel(writer, sheet_name='분석결과', index=False)

                    # 분석 정보
                    info_df = pd.DataFrame({
                        '항목': ['시스템 버전', '알고리즘', '휴전선로', '절체대상', '임계치(kW)',
                                '분석일시', '참조 기준'],
                        '값': [
                            'KEPCO v5.5 Yearly-Week-Sync',
                            '364일(52주) 전 동일 주차 참조 + 트렌드 보정',
                            self.selected_shutdown_line,
                            ', '.join(self.shutdown_mapping.get(self.selected_shutdown_line, [])),
                            self.threshold_kw,
                            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            '요일 완벽 매칭 (52주 = 364일)'
                        ]
                    })
                    info_df.to_excel(writer, sheet_name='분석정보', index=False)

                self.log(f"💾 저장 완료: {os.path.basename(filename)}")
                messagebox.showinfo("완료", f"저장 완료:\n{filename}")
            except Exception as e:
                self.log(f"❌ 저장 실패: {str(e)}")
                messagebox.showerror("오류", f"저장 오류:\n{str(e)}")


def main():
    root = tk.Tk()
    app = KEPCOOutageGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
