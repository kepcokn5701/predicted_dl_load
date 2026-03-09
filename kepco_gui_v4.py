"""
KEPCO 지능형 관제 시스템 GUI v4.0
배전선로 휴전 작업 가부 판정 시스템 (휴전 대상 선택 기능 추가)
작성: 한국전력 신입 시스템 엔지니어
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import threading
import os


class KEPCOOutageGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ KEPCO 배전선로 휴전 작업 가부 판정 시스템 v4.0")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # 데이터 변수
        self.load_file = None
        self.shutdown_file = None
        self.results_df = None
        self.threshold_kw = 14000
        self.shutdown_mapping = {}  # 휴전선로 -> 절체대상 매핑
        self.selected_shutdown_line = None
        
        self.create_widgets()
    
    def create_widgets(self):
        """GUI 위젯 생성"""
        # 헤더
        header_frame = tk.Frame(self.root, bg='#1e3a8a', height=80)
        header_frame.pack(fill='x', padx=0, pady=0)
        
        title_label = tk.Label(
            header_frame, 
            text="⚡ KEPCO 배전선로 휴전 작업 가부 판정 시스템 v4.0",
            font=('맑은 고딕', 18, 'bold'),
            fg='white',
            bg='#1e3a8a'
        )
        title_label.pack(pady=20)
        
        # 메인 컨테이너
        main_container = tk.Frame(self.root, bg='#f0f0f0')
        main_container.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 좌측 패널 (설정)
        left_panel = tk.Frame(main_container, bg='white', relief='raised', borderwidth=2)
        left_panel.pack(side='left', fill='y', padx=(0, 10), pady=0, ipadx=10, ipady=10)
        
        # 파일 선택 섹션
        file_section = tk.LabelFrame(left_panel, text="📂 데이터 파일 선택", font=('맑은 고딕', 11, 'bold'), bg='white')
        file_section.pack(fill='x', padx=10, pady=10)
        
        # 부하 데이터 파일
        tk.Label(file_section, text="부하 데이터:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 0))
        self.load_file_label = tk.Label(file_section, text="load_data.xlsx", font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.load_file_label.pack(anchor='w', padx=10, pady=(0, 5))
        
        tk.Button(
            file_section, 
            text="부하 데이터 선택", 
            command=self.select_load_file,
            bg='#3b82f6',
            fg='white',
            font=('맑은 고딕', 9),
            relief='flat',
            cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))
        
        # 절체 선로 매핑 파일
        tk.Label(file_section, text="절체 선로 매핑:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(5, 0))
        self.shutdown_file_label = tk.Label(file_section, text="shutdown_list.xlsx", font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.shutdown_file_label.pack(anchor='w', padx=10, pady=(0, 5))
        
        tk.Button(
            file_section, 
            text="절체 매핑 파일 선택", 
            command=self.select_shutdown_file,
            bg='#3b82f6',
            fg='white',
            font=('맑은 고딕', 9),
            relief='flat',
            cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))
        
        # 휴전 대상 선로 선택 섹션 (NEW!)
        shutdown_select_section = tk.LabelFrame(left_panel, text="🎯 휴전 대상 선로 선택", font=('맑은 고딕', 11, 'bold'), bg='white')
        shutdown_select_section.pack(fill='x', padx=10, pady=10)
        
        tk.Label(shutdown_select_section, text="휴전 대상 선로:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 5))
        
        self.shutdown_line_var = tk.StringVar()
        self.shutdown_line_combo = ttk.Combobox(
            shutdown_select_section,
            textvariable=self.shutdown_line_var,
            font=('맑은 고딕', 10),
            state='readonly',
            width=25
        )
        self.shutdown_line_combo.pack(fill='x', padx=10, pady=(0, 10))
        self.shutdown_line_combo.bind('<<ComboboxSelected>>', self.on_shutdown_line_selected)
        
        # 절체 대상 표시
        self.transfer_info_label = tk.Label(
            shutdown_select_section,
            text="절체 대상: -",
            font=('맑은 고딕', 8),
            bg='white',
            fg='#6b7280',
            wraplength=250,
            justify='left'
        )
        self.transfer_info_label.pack(anchor='w', padx=10, pady=(0, 10))
        
        # 분석 설정 섹션
        settings_section = tk.LabelFrame(left_panel, text="⚙️ 분석 설정", font=('맑은 고딕', 11, 'bold'), bg='white')
        settings_section.pack(fill='x', padx=10, pady=10)
        
        # 임계치 설정
        tk.Label(settings_section, text="임계치 (kW):", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 0))
        
        threshold_frame = tk.Frame(settings_section, bg='white')
        threshold_frame.pack(fill='x', padx=10, pady=(5, 5))
        
        self.threshold_var = tk.StringVar(value="14000")
        threshold_entry = tk.Entry(threshold_frame, textvariable=self.threshold_var, font=('맑은 고딕', 10), width=10)
        threshold_entry.pack(side='left')
        tk.Label(threshold_frame, text="kW (14MW)", font=('맑은 고딕', 9), bg='white', fg='#6b7280').pack(side='left', padx=(5, 0))
        
        # 분석 기간
        tk.Label(settings_section, text="분석 기간 (일):", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(10, 0))
        
        days_frame = tk.Frame(settings_section, bg='white')
        days_frame.pack(fill='x', padx=10, pady=(5, 10))
        
        self.days_var = tk.StringVar(value="30")
        days_entry = tk.Entry(days_frame, textvariable=self.days_var, font=('맑은 고딕', 10), width=10)
        days_entry.pack(side='left')
        tk.Label(days_frame, text="일", font=('맑은 고딕', 9), bg='white', fg='#6b7280').pack(side='left', padx=(5, 0))
        
        # 실행 버튼
        tk.Button(
            left_panel,
            text="🚀 분석 시작",
            command=self.run_analysis,
            bg='#16a34a',
            fg='white',
            font=('맑은 고딕', 12, 'bold'),
            relief='flat',
            cursor='hand2',
            height=2
        ).pack(fill='x', padx=10, pady=20)
        
        # 결과 저장 버튼
        tk.Button(
            left_panel,
            text="💾 결과 엑셀 저장",
            command=self.save_to_excel,
            bg='#9333ea',
            fg='white',
            font=('맑은 고딕', 10),
            relief='flat',
            cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))
        
        # 우측 패널 (결과)
        right_panel = tk.Frame(main_container, bg='white', relief='raised', borderwidth=2)
        right_panel.pack(side='right', fill='both', expand=True, padx=0, pady=0)
        
        # 탭 생성
        self.notebook = ttk.Notebook(right_panel)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 탭 1: 분석 결과
        self.results_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.results_tab, text="📋 분석 결과")
        
        # 결과 테이블
        result_frame = tk.Frame(self.results_tab, bg='white')
        result_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 테이블 스크롤바
        scroll_y = tk.Scrollbar(result_frame)
        scroll_y.pack(side='right', fill='y')
        
        scroll_x = tk.Scrollbar(result_frame, orient='horizontal')
        scroll_x.pack(side='bottom', fill='x')
        
        # 트리뷰 (테이블)
        self.tree = ttk.Treeview(
            result_frame,
            columns=('날짜', '요일', '판정', '최대부하(MW)', '여유량(MW)', '비고'),
            show='headings',
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            height=20
        )
        
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        
        # 컬럼 설정
        self.tree.heading('날짜', text='날짜')
        self.tree.heading('요일', text='요일')
        self.tree.heading('판정', text='판정')
        self.tree.heading('최대부하(MW)', text='최대부하(MW)')
        self.tree.heading('여유량(MW)', text='여유량(MW)')
        self.tree.heading('비고', text='비고')
        
        self.tree.column('날짜', width=100, anchor='center')
        self.tree.column('요일', width=60, anchor='center')
        self.tree.column('판정', width=80, anchor='center')
        self.tree.column('최대부하(MW)', width=120, anchor='center')
        self.tree.column('여유량(MW)', width=120, anchor='center')
        self.tree.column('비고', width=300, anchor='w')
        
        self.tree.pack(fill='both', expand=True)
        
        # 탭 2: 추천 일자
        self.recommendation_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.recommendation_tab, text="🎯 추천 일자")
        
        self.recommendation_text = scrolledtext.ScrolledText(
            self.recommendation_tab,
            font=('맑은 고딕', 10),
            wrap='word',
            bg='#f9fafb'
        )
        self.recommendation_text.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 탭 3: 로그
        self.log_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.log_tab, text="📝 분석 로그")
        
        self.log_text = scrolledtext.ScrolledText(
            self.log_tab,
            font=('Consolas', 9),
            wrap='word',
            bg='#1e1e1e',
            fg='#d4d4d4'
        )
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 초기 메시지
        self.log("시스템이 준비되었습니다. 데이터 파일을 선택하고 분석을 시작하세요.")
        
        # 기본 파일 설정
        if os.path.exists('load_data.xlsx'):
            self.load_file = 'load_data.xlsx'
        if os.path.exists('shutdown_list.xlsx'):
            self.shutdown_file = 'shutdown_list.xlsx'
            self.load_shutdown_mapping()
    
    def log(self, message):
        """로그 메시지 출력"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert('end', f"[{timestamp}] {message}\n")
        self.log_text.see('end')
        self.root.update()
    
    def select_load_file(self):
        """부하 데이터 파일 선택"""
        filename = filedialog.askopenfilename(
            title="부하 데이터 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.load_file = filename
            self.load_file_label.config(text=os.path.basename(filename))
            self.log(f"부하 데이터 파일 선택: {os.path.basename(filename)}")
    
    def select_shutdown_file(self):
        """절체 선로 매핑 파일 선택"""
        filename = filedialog.askopenfilename(
            title="절체 매핑 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.shutdown_file = filename
            self.shutdown_file_label.config(text=os.path.basename(filename))
            self.log(f"절체 매핑 파일 선택: {os.path.basename(filename)}")
            self.load_shutdown_mapping()
    
    def load_shutdown_mapping(self):
        """shutdown_list.xlsx 파일에서 휴전선로 -> 절체대상 매핑 로드"""
        if not self.shutdown_file or not os.path.exists(self.shutdown_file):
            return
        
        try:
            df = pd.read_excel(self.shutdown_file)
            
            # 칼럼명 찾기 (유연하게 매칭)
            cols = df.columns.tolist()
            
            shutdown_col = None
            transfer_cols = []
            
            for col in cols:
                col_str = str(col).lower()
                if '휴전' in col_str or 'shutdown' in col_str:
                    shutdown_col = col
                elif '절체' in col_str or 'transfer' in col_str:
                    transfer_cols.append(col)
            
            if not shutdown_col:
                # 첫 번째 칼럼을 휴전선로로 가정
                shutdown_col = cols[0]
            
            if len(transfer_cols) == 0:
                # 나머지 칼럼들을 절체대상으로 가정
                transfer_cols = cols[1:4] if len(cols) > 1 else []
            
            # 매핑 생성
            self.shutdown_mapping = {}
            shutdown_lines = []
            
            for _, row in df.iterrows():
                shutdown_line = str(row[shutdown_col]).strip()
                if shutdown_line and shutdown_line != 'nan' and shutdown_line != '':
                    transfer_targets = []
                    for tcol in transfer_cols[:3]:  # 최대 3개
                        if tcol in row:
                            target = str(row[tcol]).strip()
                            if target and target != 'nan' and target != '':
                                transfer_targets.append(target)
                    
                    if transfer_targets:
                        self.shutdown_mapping[shutdown_line] = transfer_targets
                        shutdown_lines.append(shutdown_line)
            
            # 드롭다운 업데이트
            self.shutdown_line_combo['values'] = shutdown_lines
            if shutdown_lines:
                self.shutdown_line_combo.current(0)
                self.on_shutdown_line_selected(None)
            
            self.log(f"✅ 휴전선로 매핑 로드: {len(shutdown_lines)}개 선로")
            
        except Exception as e:
            self.log(f"❌ 매핑 파일 로드 실패: {str(e)}")
    
    def on_shutdown_line_selected(self, event):
        """휴전 선로 선택 시 절체 대상 표시"""
        selected = self.shutdown_line_var.get()
        if selected in self.shutdown_mapping:
            targets = self.shutdown_mapping[selected]
            self.transfer_info_label.config(
                text=f"절체 대상: {', '.join(targets)}"
            )
            self.log(f"선택된 휴전선로: {selected} → 절체대상: {', '.join(targets)}")
    
    def run_analysis(self):
        """분석 실행"""
        # 파일 확인
        if not self.load_file or not os.path.exists(self.load_file):
            messagebox.showerror("오류", "부하 데이터 파일을 선택해주세요.")
            return
        
        # 휴전 선로 선택 확인
        if not self.shutdown_line_var.get():
            messagebox.showerror("오류", "휴전 대상 선로를 선택해주세요.")
            return
        
        # 임계치 확인
        try:
            self.threshold_kw = float(self.threshold_var.get())
            days = int(self.days_var.get())
        except:
            messagebox.showerror("오류", "임계치와 분석 기간은 숫자로 입력해주세요.")
            return
        
        self.selected_shutdown_line = self.shutdown_line_var.get()
        
        # 별도 스레드에서 분석 실행
        thread = threading.Thread(target=self._run_analysis_thread, args=(days,))
        thread.daemon = True
        thread.start()
    
    def _run_analysis_thread(self, days):
        """분석 실행 (스레드)"""
        try:
            self.log("=" * 60)
            self.log(f"분석 시작: 휴전선로 '{self.selected_shutdown_line}'")
            
            # 1. 데이터 로드
            self.log("📂 데이터 로딩 중...")
            df = self.load_data(self.load_file)
            
            # 2. 데이터 변환
            self.log("🔧 데이터 형식 변환 중...")
            long_df = self.convert_to_long_format(df)
            
            # 3. 맞춤형 부하 분산 (선택된 휴전선로 기준)
            self.log("⚙️ 맞춤형 부하 분산 시뮬레이션 중...")
            pivot_df = self.simulate_custom_load_distribution(long_df)
            
            # 4. 분석
            self.log(f"🔍 휴전 가부 판정 중 (향후 {days}일)...")
            self.results_df = self.analyze_outage_feasibility(pivot_df, days)
            
            # 5. 결과 표시
            self.log("📊 결과 표시 중...")
            self.display_results()
            self.display_recommendations()
            
            self.log("✅ 분석이 완료되었습니다!")
            messagebox.showinfo("완료", f"휴전선로 '{self.selected_shutdown_line}' 분석이 완료되었습니다!")
            
        except Exception as e:
            self.log(f"❌ 오류 발생: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"분석 중 오류가 발생했습니다:\n{str(e)}")
    
    def load_data(self, file_path):
        """데이터 로드"""
        df = pd.read_excel(file_path, skiprows=4)
        cols = ['S/S', 'D/L', '일자']
        for i in range(1, len(df.columns) - 2):
            cols.append(f'{i}시')
        df.columns = cols
        df['일자'] = pd.to_numeric(df['일자'], errors='coerce')
        df = df.dropna(subset=['일자'])
        df['일자'] = pd.to_datetime(df['일자'].astype(int), format='%Y%m%d')
        self.log(f"✅ 데이터 로드: {len(df)}개 레코드")
        return df
    
    def convert_to_long_format(self, df):
        """Long format 변환"""
        hour_cols = [col for col in df.columns if '시' in str(col)]
        data_list = []
        
        for _, row in df.iterrows():
            date = row['일자']
            line = str(row['D/L']).strip()
            
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
                        data_list.append({
                            'Timestamp': timestamp,
                            'Line': line,
                            'Load_kW': load_kw
                        })
                except:
                    pass
        
        long_df = pd.DataFrame(data_list)
        self.log(f"✅ 변환 완료: {len(long_df):,}개 시간대")
        return long_df
    
    def simulate_custom_load_distribution(self, long_df):
        """맞춤형 부하 분산 (선택된 휴전선로 기준)"""
        # 피벗: 시간대 x 선로 형식
        pivot_df = long_df.pivot_table(
            index='Timestamp',
            columns='Line',
            values='Load_kW',
            aggfunc='mean'
        ).reset_index()
        
        # 선택된 휴전선로의 절체대상 가져오기
        if self.selected_shutdown_line not in self.shutdown_mapping:
            raise ValueError(f"선택된 휴전선로 '{self.selected_shutdown_line}'의 매핑 정보를 찾을 수 없습니다.")
        
        transfer_targets = self.shutdown_mapping[self.selected_shutdown_line]
        
        self.log(f"   휴전 선로: {self.selected_shutdown_line}")
        self.log(f"   절체 대상: {', '.join(transfer_targets)}")
        
        # 휴전선로 부하 가져오기
        shutdown_line_col = None
        for col in pivot_df.columns:
            if str(col) == self.selected_shutdown_line or self.selected_shutdown_line in str(col):
                shutdown_line_col = col
                break
        
        if shutdown_line_col is None:
            # 선로명이 정확히 일치하지 않으면 첫 번째 선로 사용
            available_lines = [c for c in pivot_df.columns if c != 'Timestamp']
            shutdown_line_col = available_lines[0] if available_lines else None
            self.log(f"   ⚠️ 휴전선로 데이터를 찾지 못해 '{shutdown_line_col}' 사용")
        
        if shutdown_line_col is None:
            raise ValueError("부하 데이터에서 휴전선로를 찾을 수 없습니다.")
        
        # 휴전선로 부하
        shutdown_load = pivot_df[shutdown_line_col].fillna(10000)
        distributed_load = shutdown_load / 3  # 1/3씩 배분
        
        # 각 절체 대상의 합산 부하 계산
        self.transfer_results = {}
        
        for i, target in enumerate(transfer_targets, 1):
            # 절체 대상 선로 찾기
            target_col = None
            for col in pivot_df.columns:
                if str(col) == target or target in str(col):
                    target_col = col
                    break
            
            if target_col is None:
                # 데이터가 없으면 가상 부하 사용
                pivot_df[f'절체대상{i}'] = 8000
                target_col = f'절체대상{i}'
                self.log(f"   ⚠️ '{target}' 데이터 없음, 가상 부하(8MW) 사용")
            
            # 합산 부하 = 자체 부하 + (휴전선로 부하 / 3)
            original_load = pivot_df[target_col].fillna(8000)
            combined_load = original_load + distributed_load
            
            pivot_df[f'절체대상{i}_합산_kW'] = combined_load
            self.transfer_results[f'절체대상{i}'] = target
        
        # 전체 절체 대상 중 최대 부하 (가장 위험한 선로)
        combined_cols = [f'절체대상{i}_합산_kW' for i in range(1, len(transfer_targets) + 1)]
        pivot_df['최대합산부하_kW'] = pivot_df[combined_cols].max(axis=1)
        
        # 어느 선로가 최대인지 저장
        def find_max_line(row):
            max_val = row['최대합산부하_kW']
            for i in range(1, len(transfer_targets) + 1):
                if abs(row[f'절체대상{i}_합산_kW'] - max_val) < 0.01:
                    return i
            return 1
        
        pivot_df['최대부하선로번호'] = pivot_df.apply(find_max_line, axis=1)
        
        # 각 절체 대상별 부하도 저장
        for i in range(1, len(transfer_targets) + 1):
            pivot_df[f'절체대상{i}_부하'] = pivot_df[f'절체대상{i}_합산_kW']
        
        avg_load = pivot_df['최대합산부하_kW'].mean() / 1000
        max_load = pivot_df['최대합산부하_kW'].max() / 1000
        
        self.log(f"✅ 부하 분산 완료")
        self.log(f"   평균 합산 부하: {avg_load:.2f}MW")
        self.log(f"   최대 합산 부하: {max_load:.2f}MW")
        
        return pivot_df
    
    def analyze_outage_feasibility(self, pivot_df, days):
        """휴전 가부 판정 (09:00~14:00 시간대)"""
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        future_dates = [start_date + timedelta(days=i) for i in range(days)]
        
        results = []
        
        for date in future_dates:
            # 해당 날짜의 09:00~14:00 데이터
            work_hours_data = pivot_df[
                (pivot_df['Timestamp'].dt.date == date.date()) &
                (pivot_df['Timestamp'].dt.hour >= 9) &
                (pivot_df['Timestamp'].dt.hour < 14)
            ]
            
            if len(work_hours_data) > 0:
                # 실제 데이터
                max_load_kw = work_hours_data['최대합산부하_kW'].max()
                max_line_no = work_hours_data.loc[work_hours_data['최대합산부하_kW'].idxmax(), '최대부하선로번호']
                
                # 각 절체대상별 최대 부하
                transfer_loads = {}
                for i in range(1, 4):
                    col_name = f'절체대상{i}_부하'
                    if col_name in work_hours_data.columns:
                        transfer_loads[i] = work_hours_data[col_name].max()
            else:
                # 데이터 없으면 패턴 기반 예측
                all_work_hours = pivot_df[
                    (pivot_df['Timestamp'].dt.hour >= 9) &
                    (pivot_df['Timestamp'].dt.hour < 14)
                ]
                max_load_kw = all_work_hours['최대합산부하_kW'].quantile(0.95)
                max_line_no = 1
                
                transfer_loads = {}
                for i in range(1, 4):
                    col_name = f'절체대상{i}_부하'
                    if col_name in all_work_hours.columns:
                        transfer_loads[i] = all_work_hours[col_name].quantile(0.95)
            
            max_load_mw = max_load_kw / 1000
            margin_mw = self.threshold_kw / 1000 - max_load_mw
            
            # 판정
            if max_load_kw < 13000:
                status = '✅'
            elif max_load_kw < self.threshold_kw:
                status = '⚠️'
            else:
                status = '❌'
            
            weekday_kr = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
            is_weekend = date.weekday() >= 5
            
            # 비고 생성
            remarks = []
            if is_weekend:
                remarks.append('주말')
            
            # 어느 절체대상이 임계치를 넘었는지 표시
            over_threshold_lines = []
            for i, load_kw in transfer_loads.items():
                if load_kw >= self.threshold_kw:
                    target_name = self.transfer_results.get(f'절체대상{i}', f'절체대상{i}')
                    over_threshold_lines.append(f"{target_name}({load_kw/1000:.2f}MW)")
            
            if over_threshold_lines:
                remarks.append(f"부하초과: {', '.join(over_threshold_lines)}")
            elif max_load_kw < 11000:
                remarks.append('안전마진 충분')
            elif max_load_kw < 13000:
                remarks.append('양호')
            else:
                max_target = self.transfer_results.get(f'절체대상{int(max_line_no)}', f'절체대상{int(max_line_no)}')
                remarks.append(f'최대부하: {max_target}({max_load_mw:.2f}MW)')
            
            results.append({
                'Date': date.date(),
                'Weekday': weekday_kr,
                'Status': status,
                'MaxLoad_MW': max_load_mw,
                'Margin_MW': margin_mw,
                'Remarks': ', '.join(remarks),
                'IsWeekend': is_weekend,
                'IsFeasible': max_load_kw < self.threshold_kw
            })
        
        return pd.DataFrame(results)
    
    def display_results(self):
        """결과 표시"""
        # 기존 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 데이터 추가
        for _, row in self.results_df.iterrows():
            values = (
                str(row['Date']),
                row['Weekday'],
                row['Status'],
                f"{row['MaxLoad_MW']:.2f}",
                f"{row['Margin_MW']:+.2f}",
                row['Remarks']
            )
            
            # 색상 태그
            if row['Status'] == '✅':
                tag = 'safe'
            elif row['Status'] == '⚠️':
                tag = 'caution'
            else:
                tag = 'danger'
            
            self.tree.insert('', 'end', values=values, tags=(tag,))
        
        # 태그 색상 설정
        self.tree.tag_configure('safe', background='#d1fae5')
        self.tree.tag_configure('caution', background='#fef3c7')
        self.tree.tag_configure('danger', background='#fee2e2')
    
    def display_recommendations(self):
        """추천 일자 표시"""
        self.recommendation_text.delete('1.0', 'end')
        
        weekday_results = self.results_df[~self.results_df['IsWeekend']].copy()
        feasible_weekdays = weekday_results[weekday_results['IsFeasible']].copy()
        
        text = "=" * 80 + "\n"
        text += f"🎯 휴전선로 '{self.selected_shutdown_line}' 최적 작업일 추천\n"
        text += "=" * 80 + "\n\n"
        
        # 절체 대상 정보
        targets = self.shutdown_mapping.get(self.selected_shutdown_line, [])
        text += f"📋 절체 대상 선로: {', '.join(targets)}\n"
        text += f"⚙️  임계치: {self.threshold_kw/1000:.1f}MW\n"
        text += f"⏰ 작업 시간대: 09:00 ~ 14:00\n\n"
        text += "-" * 80 + "\n\n"
        
        if len(feasible_weekdays) == 0:
            text += "⚠️ 향후 분석 기간 동안 작업 가능한 평일이 없습니다!\n\n"
            text += "💡 대안:\n"
            text += "  • 작업 시간을 야간(22:00~06:00)으로 변경\n"
            text += "  • 부하가 낮은 새벽 시간대 검토\n"
            text += "  • 분석 기간을 더 길게 설정\n"
            text += "  • 다른 휴전선로 검토\n"
        else:
            feasible_weekdays = feasible_weekdays.sort_values('Margin_MW', ascending=False)
            top3 = feasible_weekdays.head(3)
            
            medals = ['🥇', '🥈', '🥉']
            for idx, (_, row) in enumerate(top3.iterrows()):
                text += f"{medals[idx]} 추천 {idx+1}순위: {row['Date']} ({row['Weekday']}요일)\n"
                text += f"   • 예상 최대 부하: {row['MaxLoad_MW']:.2f}MW\n"
                text += f"   • 안전 여유량: {row['Margin_MW']:.2f}MW ({row['Margin_MW']/14*100:.1f}%)\n"
                text += f"   • 비고: {row['Remarks']}\n"
                
                if row['MaxLoad_MW'] < 10:
                    text += f"   • 평가: ⭐⭐⭐⭐⭐ 최상의 작업 조건\n"
                elif row['MaxLoad_MW'] < 12:
                    text += f"   • 평가: ⭐⭐⭐⭐ 매우 안정적\n"
                elif row['MaxLoad_MW'] < 13:
                    text += f"   • 평가: ⭐⭐⭐ 안정적\n"
                else:
                    text += f"   • 평가: ⭐⭐ 주의 필요\n"
                text += "\n"
            
            text += "-" * 80 + "\n\n"
            text += "💡 작업 시 체크리스트\n\n"
            text += "  ✓ D-1: 기상 예보 확인 (온도 변화에 따른 부하 변동)\n"
            text += "  ✓ D-Day 08:00: 실시간 부하 재확인 및 GO/NO-GO 결정\n"
            text += "  ✓ 작업 중: 각 절체선로별 15분 간격 모니터링\n"
            text += "  ✓ 임계치 90% 도달: 즉시 작업 중단 및 부하 재평가\n"
            text += "  ✓ 작업 완료: 선로 복구 후 부하 정상화 확인\n\n"
            
            # 통계
            total = len(self.results_df)
            safe = (self.results_df['MaxLoad_MW'] < 13).sum()
            caution = ((self.results_df['MaxLoad_MW'] >= 13) & 
                      (self.results_df['MaxLoad_MW'] < 14)).sum()
            danger = (self.results_df['MaxLoad_MW'] >= 14).sum()
            
            text += "=" * 80 + "\n"
            text += "📊 통계 요약\n"
            text += "=" * 80 + "\n\n"
            text += f"  • 총 분석 기간: {total}일\n"
            text += f"  • ✅ 작업 가능 (13MW 미만): {safe}일 ({safe/total*100:.1f}%)\n"
            text += f"  • ⚠️ 작업 주의 (13-14MW): {caution}일 ({caution/total*100:.1f}%)\n"
            text += f"  • ❌ 작업 불가 (14MW 초과): {danger}일 ({danger/total*100:.1f}%)\n"
        
        self.recommendation_text.insert('1.0', text)
    
    def save_to_excel(self):
        """결과를 엑셀로 저장"""
        if self.results_df is None:
            messagebox.showwarning("경고", "저장할 분석 결과가 없습니다. 먼저 분석을 실행해주세요.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"KEPCO_휴전분석_{self.selected_shutdown_line}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                output_df = self.results_df.copy()
                output_df.columns = ['날짜', '요일', '판정', '최대부하(MW)', '여유량(MW)', '비고', '주말여부', '작업가능']
                
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    output_df.to_excel(writer, sheet_name='분석결과', index=False)
                    
                    # 설정 정보도 함께 저장
                    info_df = pd.DataFrame({
                        '항목': ['휴전선로', '절체대상', '임계치(kW)', '분석일시'],
                        '값': [
                            self.selected_shutdown_line,
                            ', '.join(self.shutdown_mapping.get(self.selected_shutdown_line, [])),
                            self.threshold_kw,
                            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        ]
                    })
                    info_df.to_excel(writer, sheet_name='분석정보', index=False)
                
                self.log(f"💾 결과 저장 완료: {os.path.basename(filename)}")
                messagebox.showinfo("저장 완료", f"결과가 저장되었습니다:\n{filename}")
            except Exception as e:
                self.log(f"❌ 저장 실패: {str(e)}")
                messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{str(e)}")


def main():
    root = tk.Tk()
    app = KEPCOOutageGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
