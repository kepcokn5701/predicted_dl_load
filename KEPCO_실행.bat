@echo off
chcp 65001 > nul
title KEPCO 배전선로 휴전 작업 가부 판정 시스템 v5.5 Enhanced

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo ⚡ KEPCO 배전선로 휴전 작업 가부 판정 시스템 v5.5 Enhanced
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.
echo [Yearly-Week-Sync 알고리즘]
echo   - 364일(52주) 전 동일 주차 데이터 동기화
echo   - 요일(Day of Week) 완벽 매칭
echo   - 최근 4주 트렌드 보정 (Scaling Factor)
echo   - 평일/주말 별도 프로파일 지원
echo.
echo [주요 기능]
echo   - 24시간 예측 그래프 시각화
echo   - 변전소 A, B, C 3개 파일 통합 분석
echo   - XGBoost 피처 엔지니어링 지원
echo.
echo 프로그램을 실행합니다...
echo.

cd /d "%~dp0"

REM Python 가상환경 확인
if not exist "%~dp0.venv\Scripts\python.exe" (
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    echo ❌ Python 가상환경을 찾을 수 없습니다!
    echo    .venv 폴더를 확인해주세요.
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    pause
    exit /b 1
)

REM GUI 실행
"%~dp0.venv\Scripts\python.exe" "%~dp0kepco_gui_v5.py"

if errorlevel 1 (
    echo.
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    echo ❌ 오류가 발생했습니다!
    echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    echo.
    pause
)
