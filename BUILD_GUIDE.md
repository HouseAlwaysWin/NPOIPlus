# 本地構建指南 / Local Build Guide

## 使用構建腳本 / Using Build Script

### 前置準備 / Prerequisites

1. 安裝 .NET SDK 6.0 或更高版本
2. 安裝 PowerShell 7.0 或更高版本（推薦）

### PowerShell 執行策略設置 / PowerShell Execution Policy Setup

如果遇到「無法載入檔案，因為這個系統上已停用指令碼執行」錯誤：

**臨時解決方案（單次使用）：**

```powershell
# 方法 1：使用 -ExecutionPolicy 參數
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Task Build

# 方法 2：設置當前 Session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\build.ps1 -Task Build
```

**永久解決方案（僅限當前用戶）：**

```powershell
# 設置當前用戶的執行策略為 RemoteSigned
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# 確認設置
Get-ExecutionPolicy -List
```

## 構建腳本使用方法 / Build Script Usage

### 基本用法 / Basic Usage

```powershell
# 顯示幫助
Get-Help .\build.ps1 -Detailed

# 構建專案（默認 Release 配置）
.\build.ps1 -Task Build

# 運行測試
.\build.ps1 -Task Test

# 打包 NuGet 套件
.\build.ps1 -Task Pack -Version 1.0.0

# 清理構建產物
.\build.ps1 -Task Clean

# 執行完整流程（清理、構建、測試、打包）
.\build.ps1 -Task All
```

### 參數說明 / Parameters

| 參數 | 類型 | 默認值 | 說明 |
|------|------|--------|------|
| `-Task` | String | 'All' | 要執行的任務：Build, Test, Pack, Clean, All |
| `-Configuration` | String | 'Release' | 構建配置：Debug 或 Release |
| `-Version` | String | '1.0.0' | NuGet 套件版本號 |

### 範例 / Examples

```powershell
# Debug 模式構建
.\build.ps1 -Task Build -Configuration Debug

# 指定版本號打包
.\build.ps1 -Task Pack -Version 1.2.3

# 完整流程（清理、構建、測試、打包）
.\build.ps1 -Task All -Configuration Release -Version 1.0.0
```

## 手動構建 / Manual Build

如果不想使用腳本，可以直接使用 dotnet CLI：

### 1. 清理 / Clean

```bash
dotnet clean NPOIPlus/NPOIPlus.csproj
dotnet clean NPOIPlusUnitTest/NPOIPlusUnitTest.csproj
```

### 2. 恢復依賴 / Restore

```bash
dotnet restore NPOIPlus/NPOIPlus.csproj
```

### 3. 構建 / Build

```bash
# Debug 模式
dotnet build NPOIPlus/NPOIPlus.csproj --configuration Debug

# Release 模式
dotnet build NPOIPlus/NPOIPlus.csproj --configuration Release
```

### 4. 運行測試 / Run Tests

```bash
dotnet test NPOIPlusUnitTest/NPOIPlusUnitTest.csproj --configuration Release --verbosity normal
```

### 5. 測試覆蓋率 / Test Coverage

```bash
# 運行測試並收集覆蓋率
dotnet test NPOIPlusUnitTest/NPOIPlusUnitTest.csproj `
    --configuration Release `
    --collect:"XPlat Code Coverage" `
    --results-directory ./TestResults

# 安裝報告生成工具
dotnet tool install -g dotnet-reportgenerator-globaltool

# 生成 HTML 報告
reportgenerator `
    -reports:"./TestResults/**/coverage.cobertura.xml" `
    -targetdir:"./CoverageReport" `
    -reporttypes:"Html"

# 打開報告
start ./CoverageReport/index.html
```

### 6. 打包 NuGet / Pack NuGet

```bash
dotnet pack NPOIPlus/NPOIPlus.csproj `
    --configuration Release `
    --output ./artifacts `
    /p:PackageVersion=1.0.0
```

### 7. 發布到本地 NuGet / Publish to Local NuGet

```bash
# 添加本地 NuGet 源（僅需執行一次）
dotnet nuget add source "D:\LocalNuGet" --name LocalNuGet

# 推送套件到本地源
dotnet nuget push ./artifacts/NPOIPlus.1.0.0.nupkg --source LocalNuGet
```

## Visual Studio 構建 / Build in Visual Studio

1. 打開 `NPOIPlus.sln`
2. 選擇配置：Debug 或 Release
3. 按 `Ctrl + Shift + B` 構建整個解決方案
4. 按 `Ctrl + R, A` 運行所有測試

## 常見問題 / Common Issues

### 問題 1：找不到 .NET SDK

**錯誤：** `The command 'dotnet' was not found`

**解決方案：**
1. 從 [Microsoft 官網](https://dotnet.microsoft.com/download) 下載並安裝 .NET SDK
2. 重啟終端或 IDE
3. 驗證安裝：`dotnet --version`

### 問題 2：NuGet 恢復失敗

**錯誤：** `Unable to load the service index for source`

**解決方案：**
```bash
# 清除 NuGet 緩存
dotnet nuget locals all --clear

# 重新恢復
dotnet restore
```

### 問題 3：測試項目找不到

**錯誤：** `Could not find a project to run`

**解決方案：**
確保在項目根目錄執行命令，或提供完整路徑：
```bash
dotnet test NPOIPlusUnitTest/NPOIPlusUnitTest.csproj
```

### 問題 4：權限錯誤

**錯誤：** `Access to the path is denied`

**解決方案：**
1. 以管理員身份運行終端
2. 檢查文件權限
3. 關閉可能鎖定文件的程序（如 Excel）

## 性能優化建議 / Performance Tips

### 1. 並行構建

```bash
dotnet build --configuration Release /maxcpucount
```

### 2. 增量構建

不要每次都使用 `Clean`，只在必要時清理。

### 3. 使用本地 NuGet 緩存

確保 NuGet 緩存未被禁用：
```bash
dotnet nuget locals all --list
```

### 4. 跳過不必要的步驟

```bash
# 跳過恢復（已手動恢復）
dotnet build --no-restore

# 跳過構建（已手動構建）
dotnet test --no-build
```

## 持續集成測試 / CI Testing

在提交前本地運行完整 CI 流程：

```powershell
# 完整測試
.\build.ps1 -Task All

# 或手動執行
dotnet clean
dotnet restore
dotnet build --configuration Release
dotnet test --configuration Release --no-build
```

## Docker 構建（可選）/ Docker Build (Optional)

如果想使用 Docker 構建：

### Dockerfile 範例

```dockerfile
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src

# 複製項目文件
COPY ["NPOIPlus/NPOIPlus.csproj", "NPOIPlus/"]
COPY ["NPOIPlusUnitTest/NPOIPlusUnitTest.csproj", "NPOIPlusUnitTest/"]

# 恢復依賴
RUN dotnet restore "NPOIPlus/NPOIPlus.csproj"

# 複製所有文件並構建
COPY . .
WORKDIR "/src/NPOIPlus"
RUN dotnet build "NPOIPlus.csproj" -c Release -o /app/build

# 運行測試
WORKDIR "/src/NPOIPlusUnitTest"
RUN dotnet test "NPOIPlusUnitTest.csproj" -c Release --no-build

# 打包
WORKDIR "/src/NPOIPlus"
RUN dotnet pack "NPOIPlus.csproj" -c Release -o /app/publish
```

### 使用 Docker 構建

```bash
# 構建 Docker 映像
docker build -t npoiplus-build .

# 從容器複製產物
docker create --name temp npoiplus-build
docker cp temp:/app/publish ./artifacts
docker rm temp
```

## 故障排除 / Troubleshooting

### 啟用詳細日誌

```bash
dotnet build --verbosity diagnostic
dotnet test --verbosity diagnostic
```

### 檢查項目狀態

```bash
# 檢查 .NET 版本
dotnet --info

# 列出所有項目
dotnet sln list

# 檢查項目引用
dotnet list NPOIPlus/NPOIPlus.csproj reference
```

### 清理所有構建產物

```powershell
# PowerShell
Get-ChildItem -Path . -Include bin,obj,TestResults,artifacts,CoverageReport -Recurse | Remove-Item -Recurse -Force

# 或使用 Git
git clean -xdf
```

## 效能基準測試 / Performance Benchmarks

### 典型構建時間（參考）

| 任務 | 時間 | 備註 |
|------|------|------|
| Clean | ~2 秒 | |
| Restore | ~5-10 秒 | 首次較慢 |
| Build | ~10-15 秒 | |
| Test | ~5-10 秒 | 依測試數量 |
| Pack | ~2-3 秒 | |
| Coverage Report | ~5-10 秒 | |
| **Total** | ~30-50 秒 | 完整流程 |

## 資源 / Resources

- [.NET CLI 文檔](https://docs.microsoft.com/en-us/dotnet/core/tools/)
- [MSBuild 參考](https://docs.microsoft.com/en-us/visualstudio/msbuild/msbuild-reference)
- [NuGet 文檔](https://docs.microsoft.com/en-us/nuget/)
- [xUnit 文檔](https://xunit.net/)


