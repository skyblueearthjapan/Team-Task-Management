@echo off
REM TeamTaskMail EXE ビルドスクリプト（STEP 8）
REM
REM 実行前の準備:
REM   1. Python 3.11+ がインストールされていること
REM   2. 仮想環境を有効化すること（推奨）:
REM        python -m venv .venv
REM        .\.venv\Scripts\activate
REM   3. 環境変数を設定すること（ビルドマシンのみ。配布先では各自設定）:
REM        set TEAMTASK_WEBAPP_URL=https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
REM        set TEAMTASK_API_TOKEN=<EXE_API_TOKEN の値>
REM
REM 実行コマンド（exe\ ディレクトリで実行）:
REM   .\build.bat

setlocal

set VERSION=1.0.0
set EXE_NAME=TeamTaskMail
set ENTRY=mail_agent\main.py
set DIST_DIR=dist\v%VERSION%

echo === TeamTaskMail v%VERSION% ビルド開始 ===
echo.

REM 依存ライブラリのインストール
echo [1/3] 依存ライブラリをインストールします...
python -m pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo エラー: pip install に失敗しました。
    exit /b %ERRORLEVEL%
)
echo.

REM バージョン情報ファイルの生成
echo [2/3] バージョン情報ファイルを生成します...
(
echo VSVersionInfo^(
echo   ffi=FixedFileInfo^(
echo     filevers=^(1, 0, 0, 0^),
echo     prodvers=^(1, 0, 0, 0^),
echo     mask=0x3f,
echo     flags=0x0,
echo     OS=0x40004,
echo     fileType=0x1,
echo     subtype=0x0,
echo     date=^(0, 0^)
echo   ^),
echo   kids=[
echo     StringFileInfo^([
echo       StringTable^(u'040904B0', [
echo         StringStruct^(u'CompanyName',      u'機械設計技術部'^),
echo         StringStruct^(u'ProductName',      u'TeamTaskMail'^),
echo         StringStruct^(u'FileVersion',      u'%VERSION%'^),
echo         StringStruct^(u'ProductVersion',   u'%VERSION%'^),
echo         StringStruct^(u'FileDescription',  u'タスク管理 メールエージェント'^),
echo       ]^)
echo     ]^),
echo     VarFileInfo^([VarStruct^(u'Translation', [1033, 1200]^)^]^)
echo   ]
echo ^)
) > version_info.txt

REM PyInstaller でビルド
echo [3/3] PyInstaller でビルドします...
pyinstaller ^
  --onefile ^
  --noconsole ^
  --name %EXE_NAME% ^
  --version-file version_info.txt ^
  %ENTRY%

if %ERRORLEVEL% neq 0 (
    echo エラー: PyInstaller ビルドに失敗しました。
    if exist version_info.txt del version_info.txt
    exit /b %ERRORLEVEL%
)

REM 出力先ディレクトリへコピー
if not exist %DIST_DIR% mkdir %DIST_DIR%
copy dist\%EXE_NAME%.exe %DIST_DIR%\%EXE_NAME%.exe

echo.
echo === ビルド完了 ===
echo 出力先: %DIST_DIR%\%EXE_NAME%.exe

REM 一時ファイルのクリーンアップ
if exist version_info.txt del version_info.txt
if exist build rmdir /s /q build
if exist %EXE_NAME%.spec del %EXE_NAME%.spec

echo.
echo 配布方法:
echo   %DIST_DIR%\%EXE_NAME%.exe を共有フォルダへコピーしてください。
echo   各スタッフは共有フォルダから自分の PC にコピーして実行します。

endlocal
