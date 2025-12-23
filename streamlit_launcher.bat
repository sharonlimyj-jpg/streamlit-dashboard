@echo off
REM Streamlit 빠른 실행 배치 파일
REM 작성일: 2025-12-22

echo ====================================
echo Streamlit 앱 실행 중...
echo ====================================
echo.

REM 1단계: Streamlit 설치 (이미 설치되어 있으면 스킵됨)
echo [1/3] Streamlit 패키지 확인 및 설치...
C:\Users\20021396\AppData\Local\Python\pythoncore-3.14-64\python.exe -m pip install streamlit --quiet
if errorlevel 1 (
    echo 오류: Streamlit 설치 실패
    pause
    exit /b 1
)
echo Streamlit 설치 완료!
echo.

REM 2단계: 작업 디렉토리로 이동 및 Streamlit 실행
echo [2/3] Streamlit 앱 시작...
cd /d C:\Users\20021396\Documents\6. 홈케어영업기획\AI\Streamlit_dash
if errorlevel 1 (
    echo 오류: 디렉토리를 찾을 수 없습니다
    pause
    exit /b 1
)

REM 3단계: 브라우저 자동 열기 및 Streamlit 실행
echo [3/3] 브라우저에서 앱 열기...
echo.
echo ====================================
echo 브라우저에서 http://localhost:8501 을 엽니다
echo 종료하려면 이 창에서 Ctrl+C를 누르세요
echo ====================================
echo.

REM 2초 후 브라우저 자동 실행
timeout /t 2 /nobreak >nul
start http://localhost:8501

REM Streamlit 앱 실행 (이 명령은 계속 실행됨)
C:\Users\20021396\AppData\Local\Python\pythoncore-3.14-64\python.exe -m streamlit run app.py

REM 앱이 종료되면 메시지 표시
echo.
echo Streamlit 앱이 종료되었습니다.
pause