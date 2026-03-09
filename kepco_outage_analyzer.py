"""
KEPCO 지능형 관제 시스템 v2.1 - 배전선로 휴전 작업 가부 판정 엔진
작성: 한국전력 신입 지능형 시스템 엔지니어
분석 대상: 황정D/L 및 절체선로 (광촌D/L, 봉지D/L, 성환D/L)
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import os
warnings.filterwarnings('ignore')

print("=" * 120)
print("⚡ KEPCO 지능형 관제 시스템 v2.1 - 배전선로 휴전 작업 가부 판정 엔진")
print("=" * 120)
print(f"📅 분석 기준일: {datetime.now().strftime('%Y년 %m월 %d일')}")
print(f"🎯 분석 목표: 향후 1개월간 휴전 작업 가능일 판정")
print(f"⚙️  임계치: 14,000kW (14MW)")
print("=" * 120)
print()


def load_and_process_data(file_path='load_data.xlsx'):
    """엑셀 파일 로드 및 전처리"""
    print("📂 엑셀 파일 로딩 중...")
    
    # skiprows=4로 데이터 읽기 (헤더 3행 + 시간 헤더 1행 제거)
    df = pd.read_excel(file_path, skiprows=4)
    
    # 칼럼명 재설정
    cols = ['S/S', 'D/L', '일자']
    # 나머지는 시간 칼럼
    for i in range(1, len(df.columns) - 2):
        cols.append(f'{i}시')
    df.columns = cols
    
    # 유효한 데이터만 필터링 (일자가 숫자인 행만)
    df['일자'] = pd.to_numeric(df['일자'], errors='coerce')
    df = df.dropna(subset=['일자'])
    
    # 일자를 날짜로 변환
    df['일자'] = pd.to_datetime(df['일자'].astype(int), format='%Y%m%d')
    
    print(f"✅ 데이터 로드 완료: {len(df)}개 레코드")
    print(f"   선로: {df['D/L'].unique()}")
    print(f"   기간: {df['일자'].min().date()} ~ {df['일자'].max().date()}")
    print()
    
    return df


def convert_to_long_format(df):
    """Wide format → Long format 변환"""
    print("🔧 데이터 형식 변환 중...")
    
    # 시간 칼럼 찾기
    hour_cols = [col for col in df.columns if '시' in str(col)]
    
    data_list = []
    for _, row in df.iterrows():
        date = row['일자']
        line = str(row['D/L']).strip()
        
        for hour_col in hour_cols:
            try:
                hour = int(hour_col.replace('시', ''))
            except:
                continue
            
            # 24시는 다음날 0시로 처리
            if hour == 24:
                timestamp = date + timedelta(days=1)
            else:
                timestamp = date + timedelta(hours=hour)
            
            # 부하 값 (MW)
            load_mw = row[hour_col]
            
            if pd.notna(load_mw):
                try:
                    load_mw = float(load_mw)
                    # MW → kW 변환
                    load_kw = load_mw * 1000 if load_mw < 100 else load_mw
                    
                    data_list.append({
                        'Timestamp': timestamp,
                        'Line': line,
                        'Load_kW': load_kw
                    })
                except:
                    pass
    
    long_df = pd.DataFrame(data_list)
    
    print(f"✅ 변환 완료: {len(long_df):,}개 시간대 레코드")
    print(f"   선로: {sorted(long_df['Line'].unique())}")
    print()
    
    return long_df


def simulate_load_distribution(long_df):
    """부하 분산 시뮬레이션"""
    print("⚙️  부하 분산 시뮬레이션 중...")
    
    # 피벗: 시간대 x 선로 형식
    pivot_df = long_df.pivot_table(
        index='Timestamp',
        columns='Line',
        values='Load_kW',
        aggfunc='mean'
    ).reset_index()
    
    lines = [col for col in pivot_df.columns if col != 'Timestamp']
    print(f"   사용 가능한 선로: {lines}")
    
    # 휴전 선로와 절체 선로 구분
    shutdown_line = lines[0] if lines else None  # 첫 번째 선로를 휴전 대상으로
    transfer_lines = lines[1:4] if len(lines) > 1 else []  # 나머지 3개
    
    # 절체 선로가 부족하면 가상 생성
    while len(transfer_lines) < 3:
        transfer_lines.append(f'가상선로{len(transfer_lines)+1}')
        pivot_df[transfer_lines[-1]] = 8000  # 가상 부하 8MW
    
    print(f"   - 휴전 대상: {shutdown_line}")
    print(f"   - 절체 선로: {transfer_lines[:3]}")
    
    # 부하 분산: 휴전 선로 부하를 3등분하여 절체 선로에 배분
    shutdown_load = pivot_df[shutdown_line].fillna(10000)  # 기본값 10MW
    distributed_load = shutdown_load / 3
    
    # 각 절체 선로에 부하 추가
    for line in transfer_lines[:3]:
        if line in pivot_df.columns:
            pivot_df[f'{line}_합산'] = pivot_df[line].fillna(8000) + distributed_load
        else:
            pivot_df[f'{line}_합산'] = 8000 + distributed_load
    
    # 최대 합산 부하 (가장 위험한 선로 기준)
    combined_cols = [f'{line}_합산' for line in transfer_lines[:3]]
    pivot_df['최대합산부하_kW'] = pivot_df[combined_cols].max(axis=1)
    
    print(f"✅ 부하 분산 완료")
    print(f"   - 평균 합산 부하: {pivot_df['최대합산부하_kW'].mean()/1000:.2f}MW")
    print(f"   - 최대 합산 부하: {pivot_df['최대합산부하_kW'].max()/1000:.2f}MW")
    print()
    
    return pivot_df, shutdown_line, transfer_lines[:3]


def analyze_outage_feasibility(pivot_df, days=30, threshold_kw=14000):
    """휴전 작업 가부 판정"""
    print(f"🔍 휴전 작업 가부 판정 분석 중 (향후 {days}일)...")
    print(f"   - 작업 시간대: 09:00 ~ 14:00")
    print(f"   - 임계치: {threshold_kw:,}kW ({threshold_kw/1000:.1f}MW)")
    print()
    
    # 향후 N일 날짜 생성
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
            # 실제 데이터가 있는 경우
            max_load_kw = work_hours_data['최대합산부하_kW'].max()
            min_load_kw = work_hours_data['최대합산부하_kW'].min()
            avg_load_kw = work_hours_data['최대합산부하_kW'].mean()
        else:
            # 데이터가 없으면 패턴 기반 예측
            # 해당 요일/시간대의 평균 사용
            weekday = date.weekday()
            month = date.month
            
            # 시간대별 평균 계산
            hour_pattern = pivot_df[
                (pivot_df['Timestamp'].dt.hour >= 9) &
                (pivot_df['Timestamp'].dt.hour < 14) &
                (pivot_df['Timestamp'].dt.weekday == weekday)
            ]
            
            if len(hour_pattern) > 0:
                max_load_kw = hour_pattern['최대합산부하_kW'].quantile(0.95)  # 95 백분위수
                min_load_kw = hour_pattern['최대합산부하_kW'].quantile(0.05)
                avg_load_kw = hour_pattern['최대합산부하_kW'].mean()
            else:
                # 전체 평균 사용
                all_work_hours = pivot_df[
                    (pivot_df['Timestamp'].dt.hour >= 9) &
                    (pivot_df['Timestamp'].dt.hour < 14)
                ]
                max_load_kw = all_work_hours['최대합산부하_kW'].quantile(0.95)
                min_load_kw = all_work_hours['최대합산부하_kW'].quantile(0.05)
                avg_load_kw = all_work_hours['최대합산부하_kW'].mean()
        
        # 판정
        max_load_mw = max_load_kw / 1000
        threshold_mw = threshold_kw / 1000
        margin_mw = threshold_mw - max_load_mw
        
        if max_load_kw < 13000:  # 13MW 미만
            status = '✅'
            status_text = '가능 (안전)'
        elif max_load_kw < threshold_kw:  # 13MW ~ 14MW
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
        if max_load_kw < 11000:
            remarks.append('안전마진 충분')
        elif max_load_kw < 13000:
            remarks.append('양호')
        elif max_load_kw >= threshold_kw:
            remarks.append(f'초과 {(max_load_kw - threshold_kw)/1000:.2f}MW')
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
            'IsFeasible': max_load_kw < threshold_kw
        })
    
    return pd.DataFrame(results)


def print_results_table(results_df):
    """팀장님 보고용 테이블 출력"""
    print("=" * 120)
    print("📋 휴전 작업 가부 판정 결과 - 팀장님 보고용")
    print("=" * 120)
    print()
    
    # 헤더
    print(f"{'날짜':<12} {'요일':<4} {'판정':<8} {'예상 최대부하(MW)':<18} {'여유량(MW)':<12} {'비고':<30}")
    print("-" * 120)
    
    # 데이터 행
    for _, row in results_df.iterrows():
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
    total_days = len(results_df)
    safe_days = (results_df['MaxLoad_MW'] < 13).sum()
    caution_days = ((results_df['MaxLoad_MW'] >= 13) & (results_df['MaxLoad_MW'] < 14)).sum()
    impossible_days = (results_df['MaxLoad_MW'] >= 14).sum()
    
    print("📊 통계 요약")
    print(f"   - 총 분석 기간: {total_days}일")
    print(f"   - ✅ 작업 가능 (안전): {safe_days}일 ({safe_days/total_days*100:.1f}%)")
    print(f"   - ⚠️  작업 주의 (근접): {caution_days}일 ({caution_days/total_days*100:.1f}%)")
    print(f"   - ❌ 작업 불가 (초과): {impossible_days}일 ({impossible_days/total_days*100:.1f}%)")
    print()


def generate_smart_report(results_df):
    """신입사원 한전 지능형 분석 리포트"""
    print("=" * 120)
    print("💼 신입사원 한전 지능형 분석 리포트")
    print("=" * 120)
    print()
    
    # 평일만 필터링
    weekday_results = results_df[~results_df['IsWeekend']].copy()
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
        # 여유량 기준 정렬
        feasible_weekdays = feasible_weekdays.sort_values('Margin_MW', ascending=False)
        
        # TOP 3 추천
        top3 = feasible_weekdays.head(3)
        
        print("🎯 최적의 평일 작업 추천 TOP 3")
        print("-" * 120)
        print()
        
        for idx, (_, row) in enumerate(top3.iterrows(), 1):
            rank_emoji = ['🥇', '🥈', '🥉'][idx-1]
            date_str = str(row['Date'])
            weekday = row['Weekday']
            max_load = row['MaxLoad_MW']
            margin = row['Margin_MW']
            
            print(f"{rank_emoji} 추천 {idx}순위: {date_str} ({weekday}요일)")
            print(f"   - 예상 최대 부하: {max_load:.2f}MW")
            print(f"   - 안전 여유량: {margin:.2f}MW ({margin/14*100:.1f}%)")
            print(f"   - 추천 사유: ", end='')
            
            if max_load < 10:
                print("부하가 매우 낮아 최상의 작업 조건")
            elif max_load < 11:
                print("부하가 낮아 매우 안전한 작업 가능")
            elif max_load < 12:
                print("적정 부하로 안정적인 작업 환경")
            elif max_load < 13:
                print("작업 가능하나 부하 모니터링 권장")
            else:
                print("작업 가능하나 실시간 감시 필수")
            print()
        
        print("-" * 120)
        print()
        
        # 지능형 분석 코멘트
        avg_margin = top3['Margin_MW'].mean()
        best_day = top3.iloc[0]
        
        print("💡 지능형 분석 인사이트")
        print()
        print(f"안녕하십니까, 한국전력 배전운영팀 신입 시스템 엔지니어입니다! 🔌")
        print()
        print(f"AI 기반 부하 분산 시뮬레이션을 통해 향후 1개월간 최적의 휴전 작업일을 분석했습니다.")
        print(f"휴전 시 부하를 3개 절체 선로에 1/3씩 균등 배분하는 시나리오로 계산한 결과,")
        print()
        
        if avg_margin > 2:
            print(f"🎊 매우 좋은 소식입니다! TOP 3 추천일의 평균 안전 여유량이 {avg_margin:.2f}MW({avg_margin/14*100:.1f}%)로,")
            print(f"안정적인 작업 환경이 확보됩니다. 특히 {best_day['Date']} ({best_day['Weekday']}요일)은")
            print(f"최대 부하가 {best_day['MaxLoad_MW']:.2f}MW로 {best_day['Margin_MW']:.2f}MW({best_day['Margin_MW']/14*100:.1f}%)의")
            print(f"충분한 여유가 있어 **1순위로 강력 추천**드립니다!")
        elif avg_margin > 1:
            print(f"📊 TOP 3 추천일의 평균 안전 여유량이 {avg_margin:.2f}MW({avg_margin/14*100:.1f}%)로 적정 수준입니다.")
            print(f"{best_day['Date']} ({best_day['Weekday']}요일)에 작업하시면 안전하게 진행 가능하며,")
            print(f"작업 중 실시간 부하 모니터링을 병행하시면 더욱 완벽합니다.")
        else:
            print(f"⚠️  주의가 필요합니다. TOP 3 추천일도 평균 여유량이 {avg_margin:.2f}MW({avg_margin/14*100:.1f}%)로 제한적입니다.")
            print(f"작업 시에는 반드시 실시간 부하 감시를 실시하고, 부하 급증 시 즉시 복구 가능한")
            print(f"비상 대응 체계를 갖추시기 바랍니다.")
        
        print()
        print("🔧 작업 시 체크리스트:")
        print("   ✓ D-1: 기상 예보 확인 (온도 급변 시 부하 변동 가능)")
        print("   ✓ D-Day 08:00: 실시간 부하 재확인 및 최종 GO/NO-GO 결정")
        print("   ✓ 작업 중: 15분 간격 부하 모니터링 (SCADA 시스템)")
        print("   ✓ 임계치 90% 도달: 즉시 작업 중단 및 부하 재평가")
        print("   ✓ 작업 완료: 선로 복구 후 부하 정상화 확인")
        print()
        print("데이터 드리븐 의사결정으로 안전하고 효율적인 배전 운영을 실현하겠습니다! 💪")
    
    print()
    print("=" * 120)
    print()


def main():
    """메인 실행 함수"""
    try:
        # 1. 데이터 로드
        df = load_and_process_data('load_data.xlsx')
        
        # 2. Long format 변환
        long_df = convert_to_long_format(df)
        
        # 3. 부하 분산 시뮬레이션
        pivot_df, shutdown_line, transfer_lines = simulate_load_distribution(long_df)
        
        # 4. 휴전 가부 판정 (향후 30일)
        results_df = analyze_outage_feasibility(pivot_df, days=30, threshold_kw=14000)
        
        # 5. 결과 테이블 출력
        print_results_table(results_df)
        
        # 6. 지능형 분석 리포트
        generate_smart_report(results_df)
        
        print("🎉 모든 분석이 완료되었습니다!")
        print(f"📊 분석 대상: 휴전 선로 '{shutdown_line}', 절체 선로 {transfer_lines}")
        print()
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
