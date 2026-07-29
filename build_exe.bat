@echo off
chcp 65001 >nul
echo ============================================================
echo   BUILD EXE - Portal BAU
echo   Đóng gói app.py thành file .exe chạy độc lập
echo ============================================================
echo.

:: Kiểm tra PyInstaller đã cài chưa
where pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] PyInstaller chưa được cài. Đang cài đặt...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [X] Không thể cài PyInstaller. Kiểm tra kết nối mạng hoặc proxy.
        pause
        exit /b 1
    )
)

:: Kiểm tra các thư viện cần thiết
echo [1/3] Kiểm tra dependencies...
pip show flask >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Đang cài flask...
    pip install flask
)
pip show pandas >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Đang cài pandas...
    pip install pandas
)
pip show openpyxl >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Đang cài openpyxl...
    pip install openpyxl
)
pip show pyzipper >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Đang cài pyzipper...
    pip install pyzipper
)

:: Xóa build cũ
echo [2/3] Dọn dẹp build cũ...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

:: Build EXE
echo [3/3] Đang build EXE... (có thể mất 1-3 phút)
echo.
pyinstaller app.spec --noconfirm

if %errorlevel% neq 0 (
    echo.
    echo [X] BUILD THẤT BẠI! Kiểm tra lỗi phía trên.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   BUILD THÀNH CÔNG!
echo   File EXE nằm tại: dist\PortalBAU\PortalBAU.exe
echo.
echo   Cách chạy:
echo     1. Mở thư mục dist\PortalBAU
echo     2. Double-click PortalBAU.exe
echo     3. Mở trình duyệt: http://localhost:5000
echo ============================================================
echo.
pause
