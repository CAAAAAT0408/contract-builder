@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 계약서 자동작성 — 깃허브 올리기
cd /d "%~dp0"
echo ============================================================
echo   바뀐 파일을 커밋하고 GitHub(CAAAAAT0408/contract-builder)에 올립니다.
echo   사이트를 고쳤다면 먼저  python bump_version.py  로 제목 버전을 올리세요.
echo   처음 올릴 때는 GitHub 로그인 창(브라우저)이 열립니다.
echo   커밋 설명을 직접 쓰려면:  "깃허브 올리기.bat" 설명글
echo ============================================================
where git >nul 2>nul
if not errorlevel 1 goto gitok
echo git 이 설치되어 있지 않습니다.  winget install Git.Git  으로 설치한 뒤 다시 실행하세요.
pause
exit /b 1
:gitok
echo.
echo [바뀐 파일]
git status --short
git add -A
git diff --cached --quiet
if errorlevel 1 goto commit
echo   새로 커밋할 변경이 없습니다. 아직 올리지 않은 커밋이 있으면 그것만 올립니다.
goto push
:commit
set "MSG=%*"
if "!MSG!"=="" set "MSG=갱신 !date! !time:~0,5!"
git commit -m "!MSG!"
:push
echo.
echo [GitHub 에 올리기]
git push
if errorlevel 1 goto fail
echo.
echo 올렸습니다. 1~2분 뒤 https://caaaaat0408.github.io/contract-builder/ 를 새로고침하세요.
goto end
:fail
echo.
echo 올리기 실패. GitHub 로그인 창을 마쳤는지, 인터넷 연결을 확인한 뒤 다시 실행하세요.
echo 다른 로그인 방법:  gh auth login --web   (계정 CAAAAAT0408)
:end
echo.
pause