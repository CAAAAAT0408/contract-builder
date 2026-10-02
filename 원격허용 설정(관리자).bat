@echo off
chcp 65001 >nul
title 원격허용 설정 - 관리자 권한 필요
set PORT=8790

net session >nul 2>&1
if errorlevel 1 (
  echo.
  echo   이 파일은 관리자 권한이 필요합니다.
  echo   파일을 마우스 오른쪽 클릭 - "관리자 권한으로 실행" 을 눌러 주세요.
  echo.
  pause
  exit /b 1
)

echo.
echo   사내망의 다른 PC 가 이 PC 의 %PORT% 번 포트로 접속할 수 있게 방화벽을 엽니다.
echo   되돌리려면 "원격허용 해제(관리자).bat" 을 실행하세요.
echo.
pause
netsh advfirewall firewall delete rule name="ContractBuilder" >nul 2>&1
netsh advfirewall firewall add rule name="ContractBuilder" dir=in action=allow protocol=TCP localport=%PORT% profile=any
echo.
echo   완료했습니다. "서버 상태.bat" 을 실행하면 팀원에게 알려 줄 주소가 나옵니다.
echo.
pause
