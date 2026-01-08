@echo off
chcp 65001 >nul
title 엑셀 측정 데이터 입력 시스템 (Streamlit)

echo ============================================================
echo      엑셀 측정 데이터 입력 시스템 (웹 버전) 시작
echo ============================================================
echo.

echo [1/2] 필요한 패키지 설치 확인 중...
python -m pip install streamlit openpyxl --quiet

if errorlevel 1 (
    echo ❌ 패키지 설치 실패!
    pause
    exit /b 1
)

echo ✓ 패키지 설치 완료
echo.

echo [2/2] 웹 브라우저에서 앱이 실행됩니다...
echo.
echo 💡 종료하려면 이 창에서 Ctrl+C를 누르세요
echo ============================================================
echo.

streamlit run streamlit_app.py

pause
