@echo off
chcp 65001 >nul
setlocal

set SOURCE_DIR=D:\code\aimt-monorepo\packages\ainote-web\dist
set TARGET_DIR=%~dp0h5\ainote-share-mobile

echo 正在删除目标目录: %TARGET_DIR%
if exist "%TARGET_DIR%" (
    del /q /s "%TARGET_DIR%\*" >nul 2>&1
    for /d %%p in ("%TARGET_DIR%\*") do rd /q /s "%%p" >nul 2>&1
)

echo 正在从源目录复制文件: %SOURCE_DIR%
if exist "%SOURCE_DIR%" (
    xcopy "%SOURCE_DIR%\*" "%TARGET_DIR%\" /E /I /Y /Q
) else (
    echo 错误: 源目录不存在: %SOURCE_DIR%
    exit /b 1
)

echo 同步完成!
