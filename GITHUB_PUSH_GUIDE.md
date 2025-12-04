# GitHub 推送完整指南 / Complete GitHub Push Guide

## 📋 推送前检查清单

### ⚠️ 第一步：配置个人信息（必须！）

#### 选项 A：使用自动化脚本（推荐）

```powershell
# 设置执行策略（如果需要）
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 运行配置脚本
.\setup-config.ps1 -GitHubUsername "你的GitHub用户名" -Email "你的邮箱" -AuthorName "你的名字"

# 例如：
.\setup-config.ps1 -GitHubUsername "martinchen" -Email "martin@example.com" -AuthorName "Martin Chen"
```

#### 选项 B：手动替换

需要在以下文件中替换：

1. **`.github/dependabot.yml`** - 第 13 行和第 31 行

   - `your-github-username` → 你的 GitHub 用户名

2. **`README.md`** - 多处

   - `your-username` → 你的 GitHub 用户名

3. **`CI_CD_SETUP.md`** - 多处

   - `your-username` → 你的 GitHub 用户名

4. **`CI_CD_SUMMARY.md`** - 多处

   - `[Your Name]` → 你的名字
   - `your-username` → 你的 GitHub 用户名

5. **`CONTRIBUTING.md`** - 多处
   - `your-email@example.com` → 你的邮箱

---

## 🚀 推送步骤

### 第二步：初始化 Git 仓库

```bash
# 如果还没有初始化
cd D:\DotNetProjects\FluentNPOI
git init
```

### 第三步：配置 Git 用户信息

```bash
git config user.name "你的名字"
git config user.email "你的邮箱"
```

### 第四步：添加文件到暂存区

```bash
# 添加所有文件
git add .

# 查看状态
git status
```

### 第五步：提交

```bash
git commit -m "feat: initial commit with complete CI/CD setup

- Add FluentNPOI library with fluent API
- Add console example project
- Add comprehensive unit tests
- Add CI/CD workflows (GitHub Actions)
- Add complete documentation (Chinese & English)
- Add build scripts and tools
- Add issue and PR templates"
```

### 第六步：在 GitHub 上创建仓库

1. 前往 https://github.com/new
2. 仓库名称：`FluentNPOI`
3. 描述：`A fluent API wrapper for NPOI to simplify Excel operations in .NET`
4. 选择：`Public` 或 `Private`
5. **不要**勾选 "Add a README file"（我们已经有了）
6. **不要**勾选 "Add .gitignore"（我们已经有了）
7. **不要**选择 License（我们已经有了）
8. 点击 "Create repository"

### 第七步：添加远程仓库并推送

```bash
# 添加远程仓库（替换你的用户名）
git remote add origin https://github.com/你的用户名/FluentNPOI.git

# 检查远程仓库
git remote -v

# 推送到 main 分支
git branch -M main
git push -u origin main
```

---

## 🔐 推送后必做：设置 GitHub Secrets

### 第八步：获取 NuGet API Key

1. 前往 https://www.nuget.org/
2. 登录你的账号
3. 点击右上角头像 > **API Keys**
4. 点击 **Create**
5. 填写信息：
   - Key Name: `FluentNPOI GitHub Actions`
   - Scopes: 选择 `Push` 和 `Push new packages and package versions`
   - Glob Pattern: `FluentNPOI*`
   - Expires: 选择一个合适的过期时间（例如 365 天）
6. 点击 **Create**
7. **复制** 生成的 API Key（只显示一次！）

### 第九步：在 GitHub 添加 Secret

1. 前往你的仓库：`https://github.com/你的用户名/FluentNPOI`
2. 点击 **Settings** 标签
3. 左侧菜单选择 **Secrets and variables** > **Actions**
4. 点击 **New repository secret**
5. 填写：
   - Name: `NUGET_API_KEY`
   - Value: 粘贴刚才复制的 NuGet API Key
6. 点击 **Add secret**

---

## ✅ 验证推送成功

### 检查清单：

1. **查看仓库**

   - 访问 `https://github.com/你的用户名/FluentNPOI`
   - 确认所有文件已上传
   - 查看 README 是否正确显示

2. **查看 Actions**

   - 点击 **Actions** 标签
   - 应该看到第一次推送触发的 CI 工作流正在运行
   - 等待工作流完成（约 3-5 分钟）
   - 确认所有检查都通过（绿色勾）

3. **查看 Branches**
   - 点击仓库主页的分支下拉菜单
   - 应该看到 `main` 分支

---

## 🎨 推送后可选操作

### 选项 1：设置分支保护规则

1. 前往 **Settings** > **Branches**
2. 点击 **Add branch protection rule**
3. Branch name pattern: `main`
4. 启用以下选项：
   - ✅ Require a pull request before merging
   - ✅ Require approvals (1)
   - ✅ Require status checks to pass before merging
   - ✅ Require branches to be up to date before merging
   - 选择必须通过的检查：
     - ✅ Build and Test
     - ✅ Code Analysis
5. 点击 **Create**

### 选项 2：添加徽章到 README

在 `README.md` 顶部添加：

```markdown
# FluentNPOI

[![CI](https://github.com/你的用户名/FluentNPOI/workflows/CI/badge.svg)](https://github.com/你的用户名/FluentNPOI/actions)
[![Code Quality](https://github.com/你的用户名/FluentNPOI/workflows/Code%20Quality/badge.svg)](https://github.com/你的用户名/FluentNPOI/actions)
[![.NET Standard 2.0](https://img.shields.io/badge/.NET%20Standard-2.0-blue.svg)](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
```

然后：

```bash
git add README.md
git commit -m "docs: add status badges"
git push
```

### 选项 3：创建第一个 Release

1. 前往仓库的 **Releases** 页面
2. 点击 **Create a new release**
3. 填写信息：
   - Tag version: `v1.0.0`
   - Release title: `Version 1.0.0 - Initial Release`
   - Description: 填写发布说明（可参考 CHANGELOG.md）
4. 点击 **Publish release**
5. GitHub Actions 会自动：
   - 构建项目
   - 运行测试
   - 打包 NuGet
   - 发布到 NuGet.org

### 选项 4：启用 GitHub Discussions（可选）

1. 前往 **Settings** > **General**
2. 滚动到 **Features** 区域
3. 勾选 **Discussions**
4. 点击 **Set up discussions**

### 选项 5：添加 Topics（标签）

1. 在仓库主页，点击齿轮图标（在 About 区域）
2. 添加相关 topics：
   - `csharp`
   - `dotnet`
   - `excel`
   - `npoi`
   - `fluent-api`
   - `library`
   - `netstandard`
3. 点击 **Save changes**

---

## 🔄 后续开发流程

### 日常开发工作流

```bash
# 1. 创建功能分支
git checkout -b feature/your-feature-name

# 2. 进行开发并测试
.\build.ps1 -Task Test

# 3. 提交变更
git add .
git commit -m "feat: add your feature"

# 4. 推送到 GitHub
git push origin feature/your-feature-name

# 5. 在 GitHub 上创建 Pull Request
# 6. 等待 CI 检查通过
# 7. 代码审查
# 8. 合并到 main 分支
```

---

## ❓ 常见问题

### Q1: 推送时要求输入用户名和密码？

A: GitHub 已不再支持密码验证，需要使用：

**方法 1：Personal Access Token (推荐)**

1. GitHub 头像 > Settings > Developer settings > Personal access tokens > Tokens (classic)
2. Generate new token (classic)
3. 选择 scopes: `repo` (全选)
4. 生成并复制 token
5. 推送时使用 token 作为密码

**方法 2：SSH Key**

```bash
# 生成 SSH Key
ssh-keygen -t ed25519 -C "your_email@example.com"

# 添加到 ssh-agent
ssh-add ~/.ssh/id_ed25519

# 复制公钥到 GitHub
cat ~/.ssh/id_ed25519.pub

# 在 GitHub: Settings > SSH and GPG keys > New SSH key
# 粘贴公钥

# 更改远程 URL 为 SSH
git remote set-url origin git@github.com:你的用户名/FluentNPOI.git
```

### Q2: CI 工作流失败怎么办？

A:

1. 前往 Actions 标签查看详细日志
2. 常见原因：
   - 测试失败：本地运行 `.\build.ps1 -Task Test` 检查
   - 构建失败：检查项目文件和依赖
   - 权限问题：检查工作流权限设置

### Q3: 如何更新已推送的代码？

A:

```bash
# 修改代码后
git add .
git commit -m "fix: your fix description"
git push
```

### Q4: 如何回退错误的提交？

A:

```bash
# 撤销最后一次提交（保留修改）
git reset --soft HEAD~1

# 或者创建一个反向提交
git revert HEAD
git push
```

---

## 📞 获取帮助

如果遇到问题：

1. 查看文档：

   - `CI_CD_SETUP.md` - CI/CD 详细说明
   - `BUILD_GUIDE.md` - 构建指南
   - `CONTRIBUTING.md` - 贡献指南

2. 查看 GitHub 帮助：

   - https://docs.github.com/

3. 创建 Issue：
   - 在仓库中点击 **Issues** > **New issue**

---

## ✅ 完成！

恭喜！你的项目现在已经：

- ✅ 在 GitHub 上公开/私有托管
- ✅ 配置了完整的 CI/CD 流程
- ✅ 准备好接受贡献
- ✅ 可以自动发布到 NuGet

开始享受自动化开发流程吧！🎉
