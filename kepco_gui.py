"""
KEPCO 지능형 관제 시스템 GUI v3.0
배전선로 휴전 작업 가부 판정 시스템
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
        self.root.title("⚡ KEPCO 배전선로 휴전 작업 가부 판정 시스템 v3.0")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # 데이터 변수
        self.load_file = None
        self.shutdown_file = None
        self.results_df = None
        self.threshold_kw = 14000
        
        self.create_widgets()
    
    def create_widgets(self):
        """GUI 위젯 생성"""
        # 헤더
        header_frame = tk.Frame(self.root, bg='#1e3a8a', height=80)
        header_frame.pack(fill='x', padx=0, pady=0)
        
        title_label = tk.Label(
            header_frame, 
            text="⚡ KEPCO 배전선로 휴전 작업 가부 판정 시스템",
            font=('맑은 고딕', 20, 'bold'),
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
        
        # 절체 선로 파일
        tk.Label(file_section, text="절체 선로:", font=('맑은 고딕', 9), bg='white').pack(anchor='w', padx=10, pady=(5, 0))
        self.shutdown_file_label = tk.Label(file_section, text="shutdown_list.xlsx", font=('맑은 고딕', 9), fg='#059669', bg='white')
        self.shutdown_file_label.pack(anchor='w', padx=10, pady=(0, 5))
        
        tk.Button(
            file_section, 
            text="절체 선로 선택", 
            command=self.select_shutdown_file,
            bg='#3b82f6',
            fg='white',
            font=('맑은 고딕', 9),
            relief='flat',
            cursor='hand2'
        ).pack(fill='x', padx=10, pady=(0, 10))
        
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
        self.tree.column('비고', width=250, anchor='w')
        
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
        """절체 선로 파일 선택"""
        filename = filedialog.askopenfilename(
            title="절체 선로 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.shutdown_file = filename
            self.shutdown_file_label.config(text=os.path.basename(filename))
            self.log(f"절체 선로 파일 선택: {os.path.basename(filename)}")
    
    def run_analysis(self):
        """분석 실행"""
        # 파일 확인
        if not self.load_file or not os.path.exists(self.load_file):
            messagebox.showerror("오류", "부하 데이터 파일을 선택해주세요.")
            return
        
        # 임계치 확인
        try:
            self.threshold_kw = float(self.threshold_var.get())
            days = int(self.days_var.get())
        except:
            messagebox.showerror("오류", "임계치와 분석 기간은 숫자로 입력해주세요.")
            return
        
        # 별도 스레드에서 분석 실행
        thread = threading.Thread(target=self._run_analysis_thread, args=(days,))
        thread.daemon = True
        thread.start()
    
    def _run_analysis_thread(self, days):
        """분석 실행 (스레드)"""
        try:
            self.log("=" * 60)
            self.log("분석을 시작합니다...")
            
            # 1. 데이터 로드
            self.log("📂 데이터 로딩 중...")
            df = self.load_data(self.load_file)
            
            # 2. 데이터 변환
            self.log("🔧 데이터 형식 변환 중...")
            long_df = self.convert_to_long_format(df)
            
            # 3. 부하 분산
            self.log("⚙️ 부하 분산 시뮬레이션 중...")
            pivot_df = self.simulate_load_distribution(long_df)
            
            # 4. 분석
            self.log(f"🔍 휴전 가부 판정 중 (향후 {days}일)...")
            self.results_df = self.analyze_outage_feasibility(pivot_df, days)
            
            # 5. 결과 표시
            self.log("📊 결과 표시 중...")
            self.display_results()
            self.display_recommendations()
            
            self.log("✅ 분석이 완료되었습니다!")
            messagebox.showinfo("완료", "분석이 성공적으로 완료되었습니다!")
            
        except Exception as e:
            self.log(f"❌ 오류 발생: {str(e)}")
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
    
    def simulate_load_distribution(self, long_df):
        """부하 분산"""
        pivot_df = long_df.pivot_table(
            index='Timestamp',
            columns='Line',
            values='Load_kW',
            aggfunc='mean'
        ).reset_index()
        
        lines = [col for col in pivot_df.columns if col != 'Timestamp']
        shutdown_line = lines[0] if lines else None
        transfer_lines = lines[1:4] if len(lines) > 1 else []
        
        while len(transfer_lines) < 3:
            transfer_lines.append(f'가상{len(transfer_lines)+1}')
            pivot_df[transfer_lines[-1]] = 8000
        
        shutdown_load = pivot_df[shutdown_line].fillna(10000)
        distributed_load = shutdown_load / 3
        
        for line in transfer_lines[:3]:
            if line in pivot_df.columns:
                pivot_df[f'{line}_합산'] = pivot_df[line].fillna(8000) + distributed_load
            else:
                pivot_df[f'{line}_합산'] = 8000 + distributed_load
        
        combined_cols = [f'{line}_합산' for line in transfer_lines[:3]]
        pivot_df['최대합산부하_kW'] = pivot_df[combined_cols].max(axis=1)
        
        self.log(f"✅ 평균 합산 부하: {pivot_df['최대합산부하_kW'].mean()/1000:.2f}MW")
        return pivot_df
    
    def analyze_outage_feasibility(self, pivot_df, days):
        """휴전 가부 판정"""
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        future_dates = [start_date + timedelta(days=i) for i in range(days)]
        
        results = []
        for date in future_dates:
            work_hours_data = pivot_df[
                (pivot_df['Timestamp'].dt.date == date.date()) &
                (pivot_df['Timestamp'].dt.hour >= 9) &
                (pivot_df['Timestamp'].dt.hour < 14)
            ]
            
            if len(work_hours_data) > 0:
                max_load_kw = work_hours_data['최대합산부하_kW'].max()
            else:
                all_work_hours = pivot_df[
                    (pivot_df['Timestamp'].dt.hour >= 9) &
                    (pivot_df['Timestamp'].dt.hour < 14)
                ]
                max_load_kw = all_work_hours['최대합산부하_kW'].quantile(0.95)
            
            max_load_mw = max_load_kw / 1000
            margin_mw = self.threshold_kw / 1000 - max_load_mw
            
            if max_load_kw < 13000:
                status = '✅'
            elif max_load_kw < self.threshold_kw:
                status = '⚠️'
            else:
                status = '❌'
            
            weekday_kr = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
            is_weekend = date.weekday() >= 5
            
            remarks = []
            if is_weekend:
                remarks.append('주말')
            if max_load_kw < 11000:
                remarks.append('안전마진 충분')
            elif max_load_kw < 13000:
                remarks.append('양호')
            elif max_load_kw >= self.threshold_kw:
                remarks.append(f'초과 {(max_load_kw - self.threshold_kw)/1000:.2f}MW')
            else:
                remarks.append(f'여유 {margin_mw:.2f}MW')
            
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
        
        text = "=" * 70 + "\n"
        text += "🎯 최적의 평일 작업 추천 TOP 3\n"
        text += "=" * 70 + "\n\n"
        
        if len(feasible_weekdays) == 0:
            text += "⚠️ 향후 분석 기간 동안 작업 가능한 평일이 없습니다!\n\n"
            text += "💡 대안:\n"
            text += "  • 작업 시간을 야간(22:00~06:00)으로 변경\n"
            text += "  • 부하가 낮은 새벽 시간대 검토\n"
            text += "  • 분석 기간을 더 길게 설정\n"
        else:
            feasible_weekdays = feasible_weekdays.sort_values('Margin_MW', ascending=False)
            top3 = feasible_weekdays.head(3)
            
            medals = ['🥇', '🥈', '🥉']
            for idx, (_, row) in enumerate(top3.iterrows()):
                text += f"{medals[idx]} 추천 {idx+1}순위: {row['Date']} ({row['Weekday']}요일)\n"
                text += f"   • 예상 최대 부하: {row['MaxLoad_MW']:.2f}MW\n"
                text += f"   • 안전 여유량: {row['Margin_MW']:.2f}MW ({row['Margin_MW']/14*100:.1f}%)\n"
                
                if row['MaxLoad_MW'] < 10:
                    text += f"   • 추천 사유: 부하가 매우 낮아 최상의 작업 조건\n"
                elif row['MaxLoad_MW'] < 12:
                    text += f"   • 추천 사유: 안정적인 작업 환경\n"
                else:
                    text += f"   • 추천 사유: 작업 가능하나 모니터링 권장\n"
                text += "\n"
            
            text += "-" * 70 + "\n\n"
            text += "💡 작업 시 체크리스트\n\n"
            text += "  ✓ D-1: 기상 예보 확인\n"
            text += "  ✓ D-Day 08:00: 실시간 부하 재확인\n"
            text += "  ✓ 작업 중: 15분 간격 모니터링\n"
            text += "  ✓ 임계치 90% 도달: 즉시 작업 중단\n"
            text += "  ✓ 작업 완료: 부하 정상화 확인\n\n"
            
            # 통계
            total = len(self.results_df)
            safe = (self.results_df['MaxLoad_MW'] < 13).sum()
            caution = ((self.results_df['MaxLoad_MW'] >= 13) & 
                      (self.results_df['MaxLoad_MW'] < 14)).sum()
            danger = (self.results_df['MaxLoad_MW'] >= 14).sum()
            
            text += "=" * 70 + "\n"
            text += "📊 통계 요약\n"
            text += "=" * 70 + "\n\n"
            text += f"  • 총 분석 기간: {total}일\n"
            text += f"  • ✅ 작업 가능 (안전): {safe}일 ({safe/total*100:.1f}%)\n"
            text += f"  • ⚠️ 작업 주의 (근접): {caution}일 ({caution/total*100:.1f}%)\n"
            text += f"  • ❌ 작업 불가 (초과): {danger}일 ({danger/total*100:.1f}%)\n"
        
        self.recommendation_text.insert('1.0', text)
    
    def save_to_excel(self):
        """결과를 엑셀로 저장"""
        if self.results_df is None:
            messagebox.showwarning("경고", "저장할 분석 결과가 없습니다. 먼저 분석을 실행해주세요.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"KEPCO_휴전분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                # 결과 데이터프레임 준비
                output_df = self.results_df.copy()
                output_df.columns = ['날짜', '요일', '판정', '최대부하(MW)', '여유량(MW)', '비고', '주말여부', '작업가능']
                
                # 엑셀 저장
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    output_df.to_excel(writer, sheet_name='분석결과', index=False)
                
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
