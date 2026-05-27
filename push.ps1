# push.ps1 - Git 提交推送脚本
# 用法: .\push.ps1 "V20260527.01_改动描述"

param(
    [string]$msg = ""
)

if (-not $msg) {
    $date = Get-Date -Format "yyyyMMdd"
    $msg = "V$date.01_update"
    Write-Host "未指定提交信息，使用默认：$msg"
}

git add .
git commit -m $msg
git push origin main

Write-Host "推送完成：$msg"
