# 貢獻指南 / Contributing Guide

感謝您考慮為 NPOIPlus 做出貢獻！/ Thank you for considering contributing to NPOIPlus!

[繁體中文](#繁體中文) | [English](#english)

---

## 繁體中文

### 如何貢獻

#### 報告 Bug

如果您發現 Bug，請：

1. **檢查現有 Issues** - 確認該 Bug 尚未被報告
2. **創建新 Issue** - 使用 Bug 報告模板
3. **提供詳細資訊**：
   - NPOIPlus 版本
   - .NET 版本
   - 作業系統
   - 重現步驟
   - 預期行為 vs 實際行為
   - 相關程式碼或錯誤訊息

#### 建議新功能

1. **檢查現有 Issues** - 確認功能尚未被提出
2. **創建 Feature Request**
3. **描述使用情境** - 說明為什麼需要這個功能
4. **提供範例** - 展示理想的 API 設計

#### 提交 Pull Request

1. **Fork 專案**
2. **創建功能分支**
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **撰寫程式碼**
   - 遵循現有程式碼風格
   - 添加必要的註解
   - 更新相關文檔
4. **撰寫測試**
   - 為新功能添加單元測試
   - 確保所有測試通過
   ```bash
   dotnet test
   ```
5. **提交變更**
   ```bash
   git commit -m "feat: add awesome feature"
   ```
6. **推送到您的 Fork**
   ```bash
   git push origin feature/your-feature-name
   ```
7. **創建 Pull Request**

### 程式碼風格

#### C# 編碼規範

- 使用 **Tab 縮排**
- 大括號 `{` 另起一行
- 類別、方法使用 **PascalCase**
- 私有欄位使用 **_camelCase**（前綴底線）
- 公開屬性使用 **PascalCase**

```csharp
public class ExampleClass
{
	private string _privateField;

	public string PublicProperty { get; set; }

	public void ExampleMethod()
	{
		// 方法實作
	}
}
```

#### 註解規範

- 公開 API 必須有 XML 文檔註解
- 複雜邏輯添加內聯註解
- 使用繁體中文或英文

```csharp
/// <summary>
/// 設置單元格的值
/// </summary>
/// <param name="value">要設置的值</param>
/// <returns>FluentCell 實例以支援鏈式調用</returns>
public FluentCell SetValue<T>(T value)
{
	// 實作
}
```

### 測試要求

#### 單元測試

- 使用 xUnit 測試框架
- 測試方法名稱：`MethodName_Scenario_ExpectedResult`
- 使用 AAA 模式（Arrange, Act, Assert）

```csharp
[Fact]
public void SetValue_WithString_ShouldSetCellValue()
{
	// Arrange
	var workbook = new XSSFWorkbook();
	var fluent = new FluentWorkbook(workbook);

	// Act
	fluent.UseSheet("Test")
		.SetCellPosition(ExcelColumns.A, 1)
		.SetValue("Hello");

	// Assert
	var cell = workbook.GetSheet("Test").GetRow(0).GetCell(0);
	Assert.Equal("Hello", cell.StringCellValue);
}
```

#### 測試覆蓋率

- 新功能必須達到 80% 以上的覆蓋率
- 重要方法必須有完整測試
- 邊界條件測試

### 提交訊息規範

使用 [Conventional Commits](https://www.conventionalcommits.org/) 規範：

```
<type>(<scope>): <subject>

<body>

<footer>
```

#### Type
- `feat`: 新功能
- `fix`: Bug 修復
- `docs`: 文檔變更
- `style`: 程式碼格式（不影響功能）
- `refactor`: 重構（不是新功能也不是 Bug 修復）
- `test`: 添加測試
- `chore`: 構建過程或輔助工具的變更

#### 範例

```
feat(FluentSheet): add GetCellValue method

Add new method to read cell values with type conversion support.

Closes #123
```

### 分支策略

- `main` - 穩定版本
- `develop` - 開發分支
- `feature/xxx` - 新功能
- `fix/xxx` - Bug 修復
- `docs/xxx` - 文檔更新

### 發布流程

1. 從 `develop` 創建 `release/vX.Y.Z` 分支
2. 更新版本號和 CHANGELOG
3. 測試
4. 合併到 `main` 並標記版本
5. 合併回 `develop`

### 代碼審查

所有 Pull Request 都需要通過代碼審查：

- 至少一位維護者批准
- 所有測試通過
- 無合併衝突
- 符合編碼規範

### 開發環境設置

```bash
# 克隆專案
git clone https://github.com/your-org/NPOIPlus.git
cd NPOIPlus

# 還原套件
dotnet restore

# 建置專案
dotnet build

# 執行測試
dotnet test

# 執行範例
cd NPOIPlusConsoleExample
dotnet run
```

### 專案結構

```
NPOIPlus/
├── NPOIPlus/                  # 主要函式庫專案
│   ├── Base/                  # 基礎類別
│   ├── Stages/                # 流暢 API 階段類別
│   ├── Models/                # 資料模型
│   └── Helpers/               # 輔助類別和擴展方法
├── NPOIPlusConsoleExample/    # 控制台範例專案
├── NPOIPlusUnitTest/          # 單元測試專案
├── README.md                  # 專案說明
├── CHANGELOG.md               # 變更記錄
└── CONTRIBUTING.md            # 本文件
```

### 需要幫助的領域

我們特別歡迎以下方面的貢獻：

- 📝 文檔改進
- 🐛 Bug 修復
- ✨ 新功能實作
- 🎨 UI/UX 改進（如果適用）
- 🌐 多語言支援
- 📊 效能優化
- 🧪 測試覆蓋率提升

### 行為準則

- 尊重所有貢獻者
- 建設性的反饋
- 包容不同意見
- 專注於問題本身而非個人

### 聯繫方式

- GitHub Issues: [提出問題](../../issues)
- GitHub Discussions: [參與討論](../../discussions)
- Email: martinwang7963@gmail.com

---

## English

### How to Contribute

#### Reporting Bugs

If you find a bug:

1. **Check existing Issues** - Make sure it hasn't been reported
2. **Create a new Issue** - Use the Bug Report template
3. **Provide details**:
   - NPOIPlus version
   - .NET version
   - Operating system
   - Steps to reproduce
   - Expected vs actual behavior
   - Relevant code or error messages

#### Suggesting Features

1. **Check existing Issues** - Make sure it hasn't been suggested
2. **Create a Feature Request**
3. **Describe the use case**
4. **Provide examples** - Show ideal API design

#### Submitting Pull Requests

1. **Fork the repository**
2. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **Write code**
   - Follow existing code style
   - Add necessary comments
   - Update documentation
4. **Write tests**
   - Add unit tests for new features
   - Ensure all tests pass
   ```bash
   dotnet test
   ```
5. **Commit changes**
   ```bash
   git commit -m "feat: add awesome feature"
   ```
6. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```
7. **Create Pull Request**

### Code Style

#### C# Coding Standards

- Use **Tab indentation**
- Opening brace `{` on new line
- Classes, methods use **PascalCase**
- Private fields use **_camelCase** (underscore prefix)
- Public properties use **PascalCase**

```csharp
public class ExampleClass
{
	private string _privateField;

	public string PublicProperty { get; set; }

	public void ExampleMethod()
	{
		// Implementation
	}
}
```

### Development Setup

```bash
# Clone repository
git clone https://github.com/your-org/NPOIPlus.git
cd NPOIPlus

# Restore packages
dotnet restore

# Build project
dotnet build

# Run tests
dotnet test

# Run examples
cd NPOIPlusConsoleExample
dotnet run
```

### Getting Help

- GitHub Issues: [Create Issue](../../issues)
- GitHub Discussions: [Join Discussion](../../discussions)
- Email: martinwang7963@gmail.com

---

## 致謝 / Acknowledgments

感謝所有為 NPOIPlus 做出貢獻的開發者！

Thank you to all developers who contributed to NPOIPlus!


