# NPOIPlus GitHub 设置配置脚本
# 运行此脚本来自动替换文档中的占位符

param(
    [Parameter(Mandatory=$true)]
    [string]$GitHubUsername,
    
    [Parameter(Mandatory=$true)]
    [string]$Email,
    
    [Parameter(Mandatory=$false)]
    [string]$AuthorName = ""
)

$ErrorActionPreference = 'Stop'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "NPOIPlus GitHub 设置配置" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "GitHub 用户名: $GitHubUsername" -ForegroundColor Yellow
Write-Host "邮箱地址: $Email" -ForegroundColor Yellow
if ($AuthorName) {
    Write-Host "作者名称: $AuthorName" -ForegroundColor Yellow
}
Write-Host ""

# 需要替换的文件列表
$files = @(
    ".github/dependabot.yml",
    "README.md",
    "CI_CD_SETUP.md",
    "CI_CD_SUMMARY.md",
    "CONTRIBUTING.md"
)

Write-Host "开始替换占位符..." -ForegroundColor Yellow
Write-Host ""

$replacementCount = 0

foreach ($file in $files) {
    $filePath = Join-Path $PSScriptRoot $file
    
    if (-not (Test-Path $filePath)) {
        Write-Host "  ⚠️  跳过: $file (文件不存在)" -ForegroundColor Yellow
        continue
    }
    
    Write-Host "  处理: $file" -ForegroundColor White
    
    try {
        $content = Get-Content $filePath -Raw -Encoding UTF8
        $originalContent = $content
        
        # 替换占位符
        $content = $content -replace 'your-github-username', $GitHubUsername
        $content = $content -replace 'your-username', $GitHubUsername
        $content = $content -replace 'your-email@example\.com', $Email
        
        if ($AuthorName) {
            $content = $content -replace '\[Your Name\]', $AuthorName
            $content = $content -replace 'your-org', $GitHubUsername
        }
        
        # 如果内容有变化，保存文件
        if ($content -ne $originalContent) {
            Set-Content $filePath $content -Encoding UTF8 -NoNewline
            Write-Host "    ✅ 已更新" -ForegroundColor Green
            $replacementCount++
        } else {
            Write-Host "    ℹ️  无需更新" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "    ❌ 错误: $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "替换完成！" -ForegroundColor Green
Write-Host "已更新 $replacementCount 个文件" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 显示下一步操作
Write-Host "✅ 下一步操作:" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. 初始化 Git 仓库（如果还没有）：" -ForegroundColor White
Write-Host "   git init" -ForegroundColor Cyan
Write-Host ""
Write-Host "2. 添加所有文件：" -ForegroundColor White
Write-Host "   git add ." -ForegroundColor Cyan
Write-Host ""
Write-Host "3. 提交：" -ForegroundColor White
Write-Host "   git commit -m `"feat: initial commit with CI/CD setup`"" -ForegroundColor Cyan
Write-Host ""
Write-Host "4. 添加远程仓库：" -ForegroundColor White
Write-Host "   git remote add origin https://github.com/$GitHubUsername/NPOIPlus.git" -ForegroundColor Cyan
Write-Host ""
Write-Host "5. 推送到 GitHub：" -ForegroundColor White
Write-Host "   git push -u origin main" -ForegroundColor Cyan
Write-Host ""
Write-Host "6. ⚠️  重要！设置 GitHub Secrets：" -ForegroundColor Yellow
Write-Host "   前往: https://github.com/$GitHubUsername/NPOIPlus/settings/secrets/actions" -ForegroundColor Cyan
Write-Host "   添加 Secret: NUGET_API_KEY" -ForegroundColor Cyan
Write-Host ""

