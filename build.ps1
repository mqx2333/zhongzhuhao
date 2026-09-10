# 打包 Windows 可执行文件（PyInstaller）
# 用法：powershell -ExecutionPolicy Bypass -File build.ps1
# 产物输出到 dist\ 目录。

$ErrorActionPreference = "Stop"

Write-Host "==> 检查 PyInstaller" -ForegroundColor Cyan
python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple | Out-Null

Write-Host "==> 打包 GUI（无控制台窗口）" -ForegroundColor Cyan
python -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name restore-zhongzhuhao-gui gui.py

Write-Host "==> 打包 CLI（带控制台）" -ForegroundColor Cyan
python -m PyInstaller --noconfirm --onefile --console `
    --name restore-zhongzhuhao-cli entry_cli.py

Write-Host "==> 完成，产物：" -ForegroundColor Green
Get-ChildItem dist\*.exe | Select-Object Name, @{n="MB"; e={[math]::Round($_.Length/1MB,2)}}
