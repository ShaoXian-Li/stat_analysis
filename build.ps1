# 一键打包两个 exe 到 dist/ 目录
#
# 用法（在仓库根目录打开 PowerShell）：
#     .\build.ps1
#     .\build.ps1 -Python "E:\program\anaconda\anaconda\envs\stat\python.exe"
<#
.SYNOPSIS
    使用 PyInstaller 打包“就业质量报告生成器”与“就业变化趋势分析器”。
.DESCRIPTION
    在本文件所在目录（仓库根目录）依次执行两个 spec 文件，产物输出到 dist/。
    打包前会检查 PyInstaller 是否可用，缺少时提示安装命令。
.PARAMETER Python
    用于打包的 Python 解释器，默认为 PATH 中的 python。
#>
param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

Write-Host "== 检查环境 ==" -ForegroundColor Cyan
& $Python -c "import sys; print('Python', sys.version.split()[0])"
& $Python -c "import PyInstaller; print('PyInstaller', PyInstaller.__version__)"
if ($LASTEXITCODE -ne 0) {
    Write-Host "未检测到 PyInstaller，请先执行： $Python -m pip install pyinstaller" -ForegroundColor Yellow
    exit 1
}

$specs = @(
    "就业质量报告生成器.spec",
    "就业变化趋势分析器.spec"
)

foreach ($spec in $specs) {
    Write-Host "== 打包 $spec ==" -ForegroundColor Cyan
    & $Python -m PyInstaller $spec --noconfirm
    if ($LASTEXITCODE -ne 0) {
        Write-Host "打包失败：$spec" -ForegroundColor Red
        exit 1
    }
}

Write-Host "== 打包完成，产物如下 ==" -ForegroundColor Green
Get-ChildItem -Path .\dist -Filter *.exe |
    Select-Object Name, @{Name = "大小(MB)"; Expression = { [math]::Round($_.Length / 1MB, 1) } } |
    Format-Table -AutoSize
