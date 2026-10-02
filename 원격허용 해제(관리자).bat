@echo off
chcp 65001 >nul
net session >nul 2>&1
if errorlevel 1 (
  echo   마우스 오른쪽 클릭 - "관리자 권한으로 실행" 으로 실행해 주세요.
  pause
  exit /b 1
)
netsh advfirewall firewall delete rule name="ContractBuilder"
echo   방화벽 허용 규칙을 지웠습니다.
pause
