"""
KEPCO 지능형 관제 시스템 - 배전선로 휴전 작업 가부 판정 엔진
작성: 한국전력 신입 지능형 시스템 엔지니어
분석 대상: 황정D/L 및 절체선로 (광촌D/L, 봉지D/L, 성환D/L)
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import os
warnings.filterwarnings('ignore')

print("=" * 100)
print("⚡ KEPCO 지능형 관제 시스템 v2.0 - 배전선로 휴전 작업 가부 판정 엔진")
print("=" * 100)
print(f"📅 분석 기준일: {datetime.now().strftime('%Y년 %m월 %d일 %H시')}")
print(f"🎯 분석 목표: 향후 1개월간 휴전 작업 가능일 판정")
print(f"⚙️  임계치: 14,000kW (14MW)")
print("=" * 100)
print()


class KEPCOSmartOutageSystem:
    """KEPCO 지능형 휴전 작업 판정 시스템"""
    
    def __init__(self, threshold_kw=14000):
        self.threshold_kw = threshold_kw  # 14,000kW = 14MW
        self.shutdown_data = None
        self.load_data = None
        self.combined_data = None
        self.analysis_results = None
        
    def detect_unit_and_convert(self, value, column_name=''):
        """
        단위 자동 감지 및 kW로 변환
        - MW 단위로 추정되는 경우 (값이 100 미만): 1000 곱하기
        - kW 단위로 추정되는 경우: 그대로 사용
        """
        if pd.isna(value):
            return np.nan
        
        # 숫자로 변환
        try:
            num_value = float(value)
        except:
            return np.nan
        
        # 단위 감지: 100 미만이면 MW로 간주, 1000 곱하기
        if num_value < 100:  # MW 단위로 추정
            return num_value * 1000
        else:  # kW 단위로 추정
            return num_value
    
    def load_excel_files(self, shutdown_file='shutdown_list.xlsx', load_file='load_data.xlsx'):
        """엑셀 파일 로드 및 데이터 정제"""
        print("📂 엑셀 파일 로딩 중...")
        
        try:
            # 부하 데이터 파일 로드 (헤더가 여러 줄)
            if os.path.exists(load_file):
                # 원본 데이터 읽기
                df_raw = pd.read_excel(load_file)
                
                # 헤더 찾기 (3번째 행이 S/S, D/L, 일자, MW...)
                # 4번째 행이 시간대 (1시, 2시, ...)
                # 실제 데이터는 5번째 행부터
                
                # 데이터 부분만 추출 (4번째 행부터)
                df_data = pd.read_excel(load_file, skiprows=3)
                
                # 칼럼명 정리
                new_columns = ['S/S', 'D/L', '일자']
                # 시간 칼럼 추가 (실제 칼럼 개수만큼)
                for i in range(1, len(df_data.columns) - 2):
                    new_columns.append(f'{i}시')
                
                df_data.columns = new_columns
                
                # NaN 행 제거 (일자가 없는 행)
                df_data = df_data.dropna(subset=['일자'])
                
                # 숫자가 아닌 일자 제거
                df_data = df_data[pd.to_numeric(df_data['일자'], errors='coerce').notna()]
                
                self.load_data = df_data
                print(f"✅ 부하 데이터 로드 완료: {len(self.load_data)}개 레코드")
                print(f"   선로: {self.load_data['D/L'].unique() if 'D/L' in self.load_data.columns else '알 수 없음'}")
            else:
                print(f"⚠️  {load_file} 파일을 찾을 수 없습니다.")
            
            # 휴전 선로 파일 로드
            if os.path.exists(shutdown_file):
                self.shutdown_data = pd.read_excel(shutdown_file)
                print(f"✅ 절체 선로 정보 로드 완료")
            else:
                print(f"⚠️  {shutdown_file} 파일을 찾을 수 없습니다.")
            
            print()
            
        except Exception as e:
            print(f"❌ 파일 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            print()
    
    def preprocess_data(self):
        """데이터 전처리 및 정제"""
        print("🔧 데이터 전처리 중...")
        
        if self.load_data is None:
            print("❌ 부하 데이터가 없습니다.")
            return
        
        # 데이터 변환: Wide format → Long format
        # 각 행이 (날짜, 선로, 시간대별 부하) 형태
        
        # 날짜 칼럼 처리
        self.load_data['일자'] = pd.to_datetime(self.load_data['일자'].astype(str), format='%Y%m%d')
        
        # 시간대 칼럼 (1시~24시)
        hour_columns = [f'{i}시' for i in range(1, 25)]
        
        # Long format으로 변환
        data_list = []
        for _, row in self.load_data.iterrows():
            date = row['일자']
            line = str(row['D/L']).strip()
            
            for hour_col in hour_columns:
                if hour_col in row:
                    hour = int(hour_col.replace('시', ''))
                    # hour가 24시면 다음날 0시로 처리
                    if hour == 24:
                        timestamp = date + timedelta(days=1, hours=0)
                    else:
                        timestamp = date + timedelta(hours=hour)
                    
                    load_mw = row[hour_col]
                    
                    # MW → kW 변환
                    if pd.notna(load_mw):
                        load_kw = self.detect_unit_and_convert(load_mw)
                        data_list.append({
                            'Timestamp': timestamp,
                            'Line': line,
                            'Load_kW': load_kw
                        })
        
        # Long format 데이터프레임 생성
        self.long_data = pd.DataFrame(data_list)
        
        print(f"✅ 데이터 변환 완료: {len(self.long_data):,}개 시간대 레코드")
        print(f"   선로: {sorted(self.long_data['Line'].unique())}")
        print(f"   기간: {self.long_data['Timestamp'].min()} ~ {self.long_data['Timestamp'].max()}")
        print()
    
    def simulate_load_distribution(self):
        """부하 분산 시뮬레이션: 휴전 선로 부하를 3개 절체 선로에 1/3씩 배분"""
        print("⚙️  부하 분산 시뮬레이션 중...")
        
        if not hasattr(self, 'long_data') or self.long_data is None:
            print("❌ 전처리된 부하 데이터가 없습니다.")
            return
        
        # 선로별로 데이터 피벗
        pivot_data = self.long_data.pivot_table(
            index='Timestamp', 
            columns='Line', 
            values='Load_kW',
            aggfunc='first'
        ).reset_index()
        
        print(f"✅ 피벗 데이터 생성: {len(pivot_data)}개 시간대")
        print(f"   선로: {[col for col in pivot_data.columns if col != 'Timestamp']}")
        
        # 선로명 찾기 (실제 데이터의 선로명 사용)
        available_lines = [col for col in pivot_data.columns if col != 'Timestamp']
        
        # 황정 선로와 절체 선로 구분
        shutdown_line = None
        transfer_lines = []
        
        for line in available_lines:
            if '황정' in str(line) or line == '1':
                shutdown_line = line
            else:
                transfer_lines.append(line)
        
        if shutdown_line is None:
            print("⚠️  휴전 대상 선로를 찾을 수 없습니다. 첫 번째 선로를 사용합니다.")
            shutdown_line = available_lines[0] if available_lines else None
            transfer_lines = available_lines[1:4] if len(available_lines) > 1 else []
        
        if not transfer_lines:
            print("⚠️  절체 선로를 찾을 수 없습니다. 가상 선로를 생성합니다.")
            transfer_lines = ['광촌D/L', '봉지D/L', '성환D/L']
            for line in transfer_lines:
                pivot_data[line] = 8000  # 가상 부하 8MW
        
        print(f"   - 휴전 선로: {shutdown_line}")
        print(f"   - 절체 선로: {', '.join(transfer_lines[:3])}")
        
        # 부하 분산 계산
        self.combined_data = pivot_data.copy()
        
        # 황정D/L 부하를 3등분
        if shutdown_line in self.combined_data.columns:
            shutdown_load = self.combined_data[shutdown_line].fillna(0)
        else:
            shutdown_load = pd.Series([12000] * len(self.combined_data))  # 기본값 12MW
        
        distributed_load = shutdown_load / 3
        
        # 각 절체 선로에 부하 추가
        for line in transfer_lines[:3]:  # 최대 3개 선로만
            if line in self.combined_data.columns:
                # 기존 부하 + 분산 부하
                original_load = self.combined_data[line].fillna(8000)
                self.combined_data[f'{line}_합산부하_kW'] = original_load + distributed_load
            else:
                # 해당 선로 데이터가 없으면 가상 부하 + 분산 부하
                self.combined_data[f'{line}_합산부하_kW'] = 8000 + distributed_load
        
        # 3개 선로 중 최대 부하 계산 (가장 위험한 선로 기준)
        max_load_cols = [f'{line}_합산부하_kW' for line in transfer_lines[:3] 
                         if f'{line}_합산부하_kW' in self.combined_data.columns]
        
        if max_load_cols:
            self.combined_data['최대합산부하_kW'] = self.combined_data[max_load_cols].max(axis=1)
        else:
            self.combined_data['최대합산부하_kW'] = 12000  # 기본값
        
        print(f"✅ 부하 분산 완료")
        print(f"   - 평균 합산 부하: {self.combined_data['최대합산부하_kW'].mean()/1000:.2f}MW")
        print(f"   - 최대 합산 부하: {self.combined_data['최대합산부하_kW'].max()/1000:.2f}MW")
        print()
    
    def analyze_outage_feasibility(self, days=30):
        """휴전 작업 가부 판정 (향후 N일, 09:00~14:00 시간대)"""
        print(f"🔍 휴전 작업 가부 판정 분석 중 (향후 {days}일)...")
        print(f"   - 작업 시간대: 09:00 ~ 14:00")
        print(f"   - 임계치: {self.threshold_kw:,}kW ({self.threshold_kw/1000:.1f}MW)")
        print()
        
        if self.combined_data is None:
            print("❌ 부하 데이터가 없습니다.")
            return
        
        # 시작일 (오늘)
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        # 향후 N일간의 날짜 생성
        future_dates = [start_date + timedelta(days=i) for i in range(days)]
        
        results = []
        
        for date in future_dates:
            # 해당 날짜의 데이터 필터링
            date_data = self.combined_data[
                (self.combined_data['Timestamp'].dt.date == date.date())
            ]
            
            # 09:00 ~ 14:00 시간대 필터링
            work_hours_data = date_data[
                (date_data['Timestamp'].dt.hour >= 9) &
                (date_data['Timestamp'].dt.hour < 14)
            ]
            
            if len(work_hours_data) == 0:
                # 데이터가 없으면 예측값 사용 (평균 부하 패턴)
                # 시간대별 평균 계산
                hour_avg = self.combined_data.groupby(
                    self.combined_data['Timestamp'].dt.hour
                )['최대합산부하_kW'].mean()
                
                work_hours_avg = hour_avg[(hour_avg.index >= 9) & (hour_avg.index < 14)]
                max_load_kw = work_hours_avg.max() if len(work_hours_avg) > 0 else 15000
                min_load_kw = work_hours_avg.min() if len(work_hours_avg) > 0 else 12000
                avg_load_kw = work_hours_avg.mean() if len(work_hours_avg) > 0 else 13500
            else:
                # 실제 데이터가 있으면 해당 데이터 사용
                max_load_kw = work_hours_data['최대합산부하_kW'].max()
                min_load_kw = work_hours_data['최대합산부하_kW'].min()
                avg_load_kw = work_hours_data['최대합산부하_kW'].mean()
            
            # 판정
            max_load_mw = max_load_kw / 1000
            threshold_mw = self.threshold_kw / 1000
            margin_mw = threshold_mw - max_load_mw
            
            if max_load_kw < 13000:  # 13MW 미만
                status = '✅'
                status_text = '가능 (안전)'
            elif max_load_kw < self.threshold_kw:  # 13MW ~ 14MW
                status = '⚠️'
                status_text = '주의 (임계치 근접)'
            else:  # 14MW 초과
                status = '❌'
                status_text = '불가 (부하 초과)'
            
            # 요일
            weekday_kr = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
            is_weekend = date.weekday() >= 5
            
            # 비고
            remarks = []
            if is_weekend:
                remarks.append('주말')
            if max_load_kw < 13000:
                remarks.append('안전마진 충분')
            elif max_load_kw >= self.threshold_kw:
                remarks.append(f'초과량 {(max_load_kw - self.threshold_kw)/1000:.2f}MW')
            else:
                remarks.append(f'여유 {margin_mw:.2f}MW')
            
            results.append({
                'Date': date.date(),
                'Weekday': weekday_kr,
                'Status': status,
                'StatusText': status_text,
                'MaxLoad_MW': max_load_mw,
                'MinLoad_MW': min_load_kw / 1000,
                'AvgLoad_MW': avg_load_kw / 1000,
                'Margin_MW': margin_mw,
                'Remarks': ', '.join(remarks),
                'IsWeekend': is_weekend,
                'IsFeasible': max_load_kw < self.threshold_kw
            })
        
        self.analysis_results = pd.DataFrame(results)
        
        print("✅ 분석 완료!")
        print()
    
    def print_results_table(self):
        """팀장님 보고용 테이블 출력"""
        if self.analysis_results is None:
            print("❌ 분석 결과가 없습니다.")
            return
        
        print("=" * 120)
        print("📋 휴전 작업 가부 판정 결과 - 팀장님 보고용")
        print("=" * 120)
        print()
        
        # 헤더
        print(f"{'날짜':<12} {'요일':<4} {'판정':<8} {'예상 최대부하(MW)':<18} {'여유량(MW)':<12} {'비고':<30}")
        print("-" * 120)
        
        # 데이터 행
        for _, row in self.analysis_results.iterrows():
            date_str = str(row['Date'])
            weekday = row['Weekday']
            status = row['Status']
            max_load = f"{row['MaxLoad_MW']:.2f}"
            margin = f"{row['Margin_MW']:+.2f}" if row['Margin_MW'] > 0 else f"{row['Margin_MW']:.2f}"
            remarks = row['Remarks']
            
            print(f"{date_str:<12} {weekday:<4} {status:<8} {max_load:<18} {margin:<12} {remarks:<30}")
        
        print("=" * 120)
        print()
        
        # 통계
        total_days = len(self.analysis_results)
        feasible_days = self.analysis_results['IsFeasible'].sum()
        safe_days = (self.analysis_results['MaxLoad_MW'] < 13).sum()
        caution_days = ((self.analysis_results['MaxLoad_MW'] >= 13) & 
                        (self.analysis_results['MaxLoad_MW'] < 14)).sum()
        impossible_days = (self.analysis_results['MaxLoad_MW'] >= 14).sum()
        
        print("📊 통계 요약")
        print(f"   - 총 분석 기간: {total_days}일")
        print(f"   - ✅ 작업 가능 (안전): {safe_days}일 ({safe_days/total_days*100:.1f}%)")
        print(f"   - ⚠️  작업 주의 (근접): {caution_days}일 ({caution_days/total_days*100:.1f}%)")
        print(f"   - ❌ 작업 불가 (초과): {impossible_days}일 ({impossible_days/total_days*100:.1f}%)")
        print()
    
    def generate_smart_report(self):
        """신입사원 한전 지능형 분석 리포트"""
        if self.analysis_results is None:
            print("❌ 분석 결과가 없습니다.")
            return
        
        print("=" * 120)
        print("💼 신입사원 한전 지능형 분석 리포트")
        print("=" * 120)
        print()
        
        # 평일만 필터링 (주말 제외)
        weekday_results = self.analysis_results[~self.analysis_results['IsWeekend']]
        
        # 작업 가능한 평일
        feasible_weekdays = weekday_results[weekday_results['IsFeasible']].copy()
        
        if len(feasible_weekdays) == 0:
            print("⚠️  향후 1개월간 작업 가능한 평일이 없습니다!")
            print()
            print("💡 대안 제시:")
            print("   1. 작업 시간을 야간(22:00~06:00)으로 변경")
            print("   2. 부하가 낮은 시간대(새벽 04:00~07:00) 검토")
            print("   3. 2개월 후 부하가 감소하는 시기로 연기")
            print()
        else:
            # 여유량 기준 정렬 (안전한 날 우선)
            feasible_weekdays = feasible_weekdays.sort_values('Margin_MW', ascending=False)
            
            # TOP 3 추천
            top3 = feasible_weekdays.head(3)
            
            print("🎯 최적의 평일 작업 추천 TOP 3")
            print("-" * 120)
            print()
            
            for idx, (_, row) in enumerate(top3.iterrows(), 1):
                date_str = str(row['Date'])
                weekday = row['Weekday']
                max_load = row['MaxLoad_MW']
                margin = row['Margin_MW']
                
                print(f"🥇 추천 {idx}순위: {date_str} ({weekday}요일)")
                print(f"   - 예상 최대 부하: {max_load:.2f}MW")
                print(f"   - 안전 여유량: {margin:.2f}MW ({margin/14*100:.1f}%)")
                print(f"   - 추천 사유: ", end='')
                
                if max_load < 11:
                    print("부하가 매우 낮아 안전한 작업 가능")
                elif max_load < 12:
                    print("부하가 낮아 안정적인 작업 환경")
                elif max_load < 13:
                    print("적정 부하로 안전하게 작업 가능")
                else:
                    print("작업 가능하나 실시간 모니터링 필요")
                print()
            
            print("-" * 120)
            print()
            
            # 지능형 분석 코멘트
            avg_margin = top3['Margin_MW'].mean()
            best_day = top3.iloc[0]
            
            print("💡 지능형 분석 인사이트")
            print()
            print(f"안녕하십니까, 한국전력 배전운영팀 신입 시스템 엔지니어입니다!")
            print()
            print(f"AI 기반 부하 분산 시뮬레이션을 통해 향후 1개월간 휴전 작업 가능일을 분석했습니다.")
            print(f"황정D/L 휴전 시 부하를 광촌D/L, 봉지D/L, 성환D/L 3개 선로에 1/3씩 균등 배분하여")
            print(f"가장 안전한 작업 시나리오를 도출했습니다.")
            print()
            
            if avg_margin > 2:
                print(f"🎊 희소식입니다! TOP 3 추천일의 평균 안전 여유량이 {avg_margin:.2f}MW로,")
                print(f"매우 안정적인 작업 환경이 확보됩니다. 특히 {best_day['Date']} ({best_day['Weekday']}요일)은")
                print(f"최대 부하 {best_day['MaxLoad_MW']:.2f}MW로 {best_day['Margin_MW']:.2f}MW의 충분한 여유가 있어")
                print(f"가장 안전한 작업일로 강력히 추천드립니다.")
            elif avg_margin > 1:
                print(f"📊 TOP 3 추천일의 평균 안전 여유량이 {avg_margin:.2f}MW로 적정 수준입니다.")
                print(f"{best_day['Date']} ({best_day['Weekday']}요일)에 작업하시면 안전하게 진행 가능하며,")
                print(f"실시간 부하 모니터링 체계를 갖추시면 더욱 완벽합니다.")
            else:
                print(f"⚠️  주의가 필요합니다. TOP 3 추천일도 평균 여유량이 {avg_margin:.2f}MW로 제한적입니다.")
                print(f"작업 시에는 반드시 실시간 부하 감시를 실시하고, 부하 급증 시 즉시 복구할 수 있는")
                print(f"비상 대응 체계를 갖추시기 바랍니다.")
            
            print()
            print("🔧 작업 시 체크리스트:")
            print("   ✓ 작업 전날 기상 예보 확인 (온도 급변 시 부하 변동)")
            print("   ✓ 작업 당일 오전 8시 실시간 부하 재확인")
            print("   ✓ 작업 중 15분 간격 부하 모니터링")
            print("   ✓ 임계치 90% 도달 시 즉시 작업 중단 및 복구 준비")
            print()
            print("데이터 기반 의사결정으로 안전하고 효율적인 배전 운영을 실현하겠습니다!")
            print("감사합니다! 💪")
        
        print()
        print("=" * 120)
        print()


def main():
    """메인 실행 함수"""
    # 1. 시스템 객체 생성
    system = KEPCOSmartOutageSystem(threshold_kw=14000)
    
    # 2. 엑셀 파일 로드
    system.load_excel_files('load_data.xlsx', 'shutdown_list.xlsx')
    
    # 3. 데이터 전처리
    system.preprocess_data()
    
    # 4. 부하 분산 시뮬레이션
    system.simulate_load_distribution()
    
    # 5. 휴전 가부 판정 (향후 30일)
    system.analyze_outage_feasibility(days=30)
    
    # 6. 결과 테이블 출력
    system.print_results_table()
    
    # 7. 지능형 분석 리포트
    system.generate_smart_report()
    
    print("🎉 모든 분석이 완료되었습니다!")
    print()


if __name__ == "__main__":
    main()
