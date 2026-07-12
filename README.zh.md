<div align="center">

# ⌨️ WordVim

**Microsoft Word 的 Vim 键绑定**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)]()
[![Word](https://img.shields.io/badge/Word-365%20Desktop-orange.svg)]()
[![.NET](https://img.shields.io/badge/.NET-Framework%204.7.2-purple.svg)]()
[![Status](https://img.shields.io/badge/Status-Proof%20of%20Concept-brightgreen.svg)]()
[![Release](https://img.shields.io/github/v/release/spidychoipro/wordvim.svg)](https://github.com/spidychoipro/wordvim/releases)

[English](README.md) · [한국어](README.ko.md) · [日本語](README.ja.md) · **中文**

---

**WordVim** 为 Microsoft Word 带来 Vim 的模态编辑体验。

按 `Escape` 进入 Normal 模式，按 `i` 返回 Insert 模式 — 就像 Vim 一样。

**无需 Visual Studio。** 使用 `dotnet build` 构建。一键安装。

[**⬇️ 下载 v0.1.0**](https://github.com/spidychoipro/wordvim/releases/tag/v0.1.0) · [**📖 文档**](#-文档) · [**🚀 快速开始**](#-快速开始)

---

</div>

## 🎯 为什么选择 WordVim

Microsoft Word 缺少 Vim 强大的模态编辑功能。**Office.js 无法拦截按键** — 这是一个无法绕过的[架构限制](https://github.com/nickebbitt/office-dev-journey/issues/6454)。

WordVim 使用 **原生 COM 加载项 + Win32 键盘钩子** 解决了这个问题 — 这是在 Word 中实现真正 Vim 模态编辑的唯一方法。

<div align="center">

```
┌─────────────────────────────────────────────────────────┐
│  Word: -- NORMAL --                                     │
│                                                         │
│  vim → Word 集成                                       │
│                                                         │
│  ✅ 模态编辑 (Insert/Normal/Visual)                     │
│  ✅ 通过 Win32 钩子拦截键盘输入                         │
│  ✅ 通过 Word COM API 移动光标                          │
│  ✅ 一键安装 — 无需 Visual Studio                       │
└─────────────────────────────────────────────────────────┘
```

</div>

## ✨ 功能特性

| 功能 | 状态 |
|------|------|
| Insert / Normal 模式切换 | ✅ 已完成 |
| 模式指示器（标题栏） | ✅ 已完成 |
| Normal 模式下按键抑制 | ✅ 已完成 |
| `h`/`j`/`k`/`l` 光标移动 | 🔜 下一步 |
| `w`/`b`/`e` 单词移动 | 🔜 计划中 |
| `x`/`dd`/`yy` 编辑命令 | 🔜 计划中 |
| Visual 模式 | 🔜 计划中 |
| 命令行模式 (`:`) | 🔜 计划中 |

## 🚀 快速开始

### 方式一：下载安装（推荐）

> **步骤 1：** 下载 [`wordvim-v0.1.0.zip`](https://github.com/spidychoipro/wordvim/releases/download/v0.1.0/wordvim-v0.1.0.zip)

> **步骤 2：** 将 ZIP 文件解压到任意位置

> **步骤 3：** 右键点击 `install.bat` → **以管理员身份运行**

> **步骤 4：** 打开 Microsoft Word → 在标题栏看到 `-- NORMAL --` 🎉

```
📁 wordvim-v0.1.0.zip
   └── 📁 wordvim
       ├── 📄 install.bat      ← 双击此文件（管理员权限）
       ├── 📄 uninstall.bat
       ├── 📄 register.ps1
       ├── 📄 unregister.ps1
       ├── 📄 PoC.dll
       └── 📄 PoC.pdb
```

### 方式二：从源码构建

```bash
# 前提条件：.NET SDK 8.0+ 和 Git
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim/poc
dotnet build

# 注册（以管理员身份运行 PowerShell）
.\register.ps1
```

## 📖 文档

| 文档 | 说明 |
|------|------|
| [架构分析](docs/architecture.md) | 为什么选择 COM 加载项 + Win32 钩子，而不是 VSTO 或 Office.js |
| [API 研究](docs/research.md) | Word COM API 的功能和限制 |
| [MVP 规格](docs/mvp.md) | 最小可行产品包含的内容 |
| [里程碑](docs/milestones.md) | 30 天开发路线图 |
| [PoC 计划](docs/poc-plan.md) | 概念验证计划 |
| [环境指南](docs/environment.md) | 开发环境设置 |
| [贡献指南](CONTRIBUTING.md) | 贡献者指南和代码规范 |

## 🏗️ 架构

```
┌──────────────────────────────────────────────────────────────┐
│                     WordVim 架构                             │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐     │
│  │   Win32     │    │   Word      │    │   Vim       │     │
│  │   钩子      │───▶│   COM API   │───▶│   状态      │     │
│  │  (键盘)     │    │  (光标)     │    │   机        │     │
│  └─────────────┘    └─────────────┘    └─────────────┘     │
│         │                  │                   │             │
│         ▼                  ▼                   ▼             │
│  ┌─────────────────────────────────────────────────────┐    │
│  │            COM 加载项 (IDTExtensibility2)           │    │
│  │               WordVim.Connect                       │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

| 组件 | 技术 | 用途 |
|------|------|------|
| **加载项** | COM + IDTExtensibility2 | 加载到 Word 进程 |
| **键盘** | WH_KEYBOARD 钩子 | 在 Word 之前拦截按键 |
| **光标** | Word COM API | 移动光标、编辑文本 |
| **状态** | C# 状态机 | 跟踪 Insert/Normal 模式 |

## 🎮 使用方法

### 按键绑定

| 按键 | 模式 | 动作 |
|------|------|------|
| `Escape` | Insert → Normal | 切换到 Normal 模式 |
| `i` | Normal → Insert | 切换到 Insert 模式 |
| `h`/`j`/`k`/`l` | Normal | 光标移动 *(计划中)* |
| `w`/`b`/`e` | Normal | 单词移动 *(计划中)* |
| `x` | Normal | 删除字符 *(计划中)* |
| `dd` | Normal | 删除行 *(计划中)* |
| `yy` | Normal | 复制行 *(计划中)* |

### 模式指示器

窗口标题栏显示当前模式：

```
-- NORMAL --  ← 除 'i' 外的所有按键都被抑制
-- INSERT --  ← 普通打字模式
```

## ❓ 常见问题

<details>
<summary><b>在 Word Online 上能用吗？</b></summary>
<br>
不能。WordVim 需要 Windows 桌面版 Word 365。Word Online 使用不支持 COM 加载项的不同架构。
</details>

<details>
<summary><b>需要 Visual Studio 吗？</b></summary>
<br>
不需要。WordVim 使用 `dotnet build` 构建 — 只需安装 .NET SDK。
</details>

<details>
<summary><b>安全吗？</b></summary>
<br>
安全。WordVim 是开源的（MIT 许可证）。安装程序只注册 COM 加载项 — 不修改系统文件。
</details>

<details>
<summary><b>如何卸载？</b></summary>
<br>
右键点击 `uninstall.bat` → 以管理员身份运行。完成。
</details>

## 🤝 贡献

请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解贡献者指南。

```bash
# 提交规范
feat(scope): 描述
fix(scope): 描述
docs: 描述

# 示例
feat(core): 添加 w/b/e 单词移动
fix(hooks): 防止委托 GC 收集
docs: 更新第三阶段里程碑
```

## 📊 项目状态

```
✅ 研究 & 架构              ████████████ 100%
✅ PoC: 加载项加载          ████████████ 100%
✅ PoC: 键盘钩子            ████████████ 100%
✅ PoC: 模式切换            ████████████ 100%
🔜 阶段 0: 基础            ░░░░░░░░░░░░   0%
🔜 阶段 1: 核心引擎        ░░░░░░░░░░░░   0%
🔜 阶段 2: Word 集成       ░░░░░░░░░░░░   0%
🔜 阶段 3: 命令            ░░░░░░░░░░░░   0%
🔜 阶段 4: Visual 模式     ░░░░░░░░░░░░   0%
🔜 阶段 5: v1.0 发布       ░░░░░░░░░░░░   0%
```

## 📄 许可证

[MIT 许可证](LICENSE) — 免费使用、修改和分发。

## 🙏 致谢

- [Vim](https://www.vim.org/) — 启发此项目的传奇文本编辑器
- [Word Object Model](https://learn.microsoft.com/en-us/office/vba/api/word.word) — Microsoft 的 Word 自动化 COM API
- [WH_KEYBOARD Hook](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw) — 用于键盘拦截的 Win32 API

---

<div align="center">

**为 Vim 用户用 ❤️ 打造**

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-ffdd00?style=flat&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/spidychoi)

[⬆️ 回到顶部](#-wordvim)

</div>
