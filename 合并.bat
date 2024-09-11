@echo off
chcp 65001 >nul
:menu
cls
echo 请选择一个选项:
echo 1. 无父体1.0合并
echo 2. 有父体1.0合并
echo 3. 无父体2.0合并
echo 4. 有父体2.0合并
echo.
set /p choice=请输入1到4之间的数字并按回车: 

if "%choice%"=="1" (
    echo 正在运行 'no parent1.0.py'...
    python "no parent1.0.py"
    goto end
)
if "%choice%"=="2" (
    echo 正在运行 'parent1.0.py'...
    python "parent1.0.py"
    goto end
)
if "%choice%"=="3" (
    echo 正在运行 'no parent2.0.py'...
    python "no parent2.0.py"
    goto end
)
if "%choice%"=="4" (
    echo 正在运行 'parent2.0.py'...
    python "parent2.0.py"
    goto end
)

echo 无效的选项，程序退出。
exit /b

:end
pause
