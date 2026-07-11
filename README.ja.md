<div align="center">

# ⌨️ WordVim

**Microsoft Word 用 Vim キーバインド**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)]()
[![Word](https://img.shields.io/badge/Word-365%20Desktop-orange.svg)]()
[![.NET](https://img.shields.io/badge/.NET-Framework%204.7.2-purple.svg)]()
[![Status](https://img.shields.io/badge/Status-Proof%20of%20Concept-brightgreen.svg)]()
[![Release](https://img.shields.io/github/v/release/spidychoipro/wordvim.svg)](https://github.com/spidychoipro/wordvim/releases)

[English](README.md) · [한국어](README.ko.md) · **日本語** · [中文](README.zh.md)

---

**WordVim** は、Microsoft Word で Vim のモーダル編集体験を提供します。

`Escape` で Normal モードに切り替え、`i` で Insert モードに戻ります — Vim と同じです。

**Visual Studio は不要です。** `dotnet build` でビルドします。ワンクリックインストール。

[**⬇️ v0.1.0 ダウンロード**](https://github.com/spidychoipro/wordvim/releases/tag/v0.1.0) · [**📖 ドキュメント**](#-ドキュメント) · [**🚀 クイックスタート**](#-クイックスタート)

---

</div>

## 🎯 WordVim を作る理由

Microsoft Word には Vim の強力なモーダル編集がありません。**Office.js はキーストロークを傍受できません** — これは[アーキテクチャ上の制限](https://github.com/nickebbitt/office-dev-journey/issues/6454)であり、回避することはできません。

WordVim は **ネイティブ COM アドイン + Win32 キーボードフック** を使用してこの問題を解決します — Word で本物の Vim モーダル編集を可能にする唯一のアプローチです。

<div align="center">

```
┌─────────────────────────────────────────────────────────┐
│  Word: -- NORMAL --                                     │
│                                                         │
│  vim → Word 統合                                       │
│                                                         │
│  ✅ モーダル編集 (Insert/Normal/Visual)                  │
│  ✅ Win32 フックによるキーボード傍受                     │
│  ✅ Word COM API によるカーソル移動                     │
│  ✅ ワンクリックインストール — Visual Studio 不要       │
└─────────────────────────────────────────────────────────┘
```

</div>

## ✨ 機能

| 機能 | ステータス |
|------|-----------|
| Insert / Normal モード切替 | ✅ 完了 |
| モードインジケーター (タイトルバー) | ✅ 完了 |
| Normal モードでのキー抑制 | ✅ 完了 |
| `h`/`j`/`k`/`l` カーソル移動 | 🔜 次 |
| `w`/`b`/`e` ワードモーション | 🔜 予定 |
| `x`/`dd`/`yy` 編集コマンド | 🔜 予定 |
| Visual モード | 🔜 予定 |
| コマンドラインモード (`:`) | 🔜 予定 |

## 🚀 クイックスタート

### オプション 1: ダウンロード (推奨)

> **ステップ 1:** [`wordvim-v0.1.0.zip`](https://github.com/spidychoipro/wordvim/releases/download/v0.1.0/wordvim-v0.1.0.zip) をダウンロード

> **ステップ 2:** ZIP ファイルを任意の場所に展開

> **ステップ 3:** `install.bat` を右クリック → **管理者として実行**

> **ステップ 4:** Microsoft Word を開く → タイトルバーに `-- NORMAL --` が表示 🎉

```
📁 wordvim-v0.1.0.zip
   └── 📁 wordvim
       ├── 📄 install.bat      ← このファイルをダブルクリック (管理者権限)
       ├── 📄 uninstall.bat
       ├── 📄 register.ps1
       ├── 📄 unregister.ps1
       ├── 📄 PoC.dll
       └── 📄 PoC.pdb
```

### オプション 2: ソースからビルド

```bash
# 前提条件: .NET SDK 8.0+ と Git
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim/poc
dotnet build

# 登録 (PowerShell を管理者として実行)
.\register.ps1
```

## 📖 ドキュメント

| ドキュメント | 説明 |
|-------------|------|
| [アーキテクチャ分析](docs/architecture.md) | COM アドイン + Win32 フックを選択した理由、VSTO/Office.js でない理由 |
| [API 研究](docs/research.md) | Word COM API の機能と制限 |
| [MVP 仕様](docs/mvp.md) | 最小限の製品に含まれるもの |
| [マイルストーン](docs/milestones.md) | 30日間の開発ロードマップ |
| [PoC 計画](docs/poc-plan.md) | 概念実証の検証計画 |
| [環境ガイド](docs/environment.md) | 開発環境のセットアップ |
| [コントリビューション](AGENTS.md) | コントリビューターガイドラインとコード規約 |

## 🏗️ アーキテクチャ

```
┌──────────────────────────────────────────────────────────────┐
│                   WordVim アーキテクチャ                      │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐     │
│  │   Win32     │    │   Word      │    │   Vim       │     │
│  │   フック    │───▶│   COM API   │───▶│   ステート  │     │
│  │  (キーボード)│    │  (カーソル) │    │   マシン    │     │
│  └─────────────┘    └─────────────┘    └─────────────┘     │
│         │                  │                   │             │
│         ▼                  ▼                   ▼             │
│  ┌─────────────────────────────────────────────────────┐    │
│  │            COM アドイン (IDTExtensibility2)          │    │
│  │               WordVim.Connect                       │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

| コンポーネント | 技術 | 目的 |
|---------------|------|------|
| **アドイン** | COM + IDTExtensibility2 | Word プロセスにロード |
| **キーボード** | WH_KEYBOARD フック | Word より先にキーストロークを傍受 |
| **カーソル** | Word COM API | カーソル移動、テキスト編集 |
| **ステート** | C# ステートマシン | Insert/Normal モードを追跡 |

## 🎮 使用方法

### キーバインド

| キー | モード | アクション |
|------|-------|-----------|
| `Escape` | Insert → Normal | Normal モードに切替 |
| `i` | Normal → Insert | Insert モードに切替 |
| `h`/`j`/`k`/`l` | Normal | カーソル移動 *(予定)* |
| `w`/`b`/`e` | Normal | ワードモーション *(予定)* |
| `x` | Normal | 文字削除 *(予定)* |
| `dd` | Normal | 行削除 *(予定)* |
| `yy` | Normal | 行コピー *(予定)* |

### モードインジケーター

ウィンドウタイトルバーに現在のモードが表示されます:

```
-- NORMAL --  ← 'i' を除くすべてのキーが抑制される
-- INSERT --  ← 通常のタイピングモード
```

## ❓ よくある質問

<details>
<summary><b>Word Online で動作しますか?</b></summary>
<br>
いいえ。WordVim は Windows のデスクトップ版 Word 365 が必要です。Word Online は COM アドインをサポートしない別のアーキテクチャを使用しています。
</details>

<details>
<summary><b>Visual Studio が必要ですか?</b></summary>
<br>
いいえ。WordVim は `dotnet build` でビルドされます — .NET SDK をインストールするだけです。
</details>

<details>
<summary><b>安全ですか?</b></summary>
<br>
はい。WordVim はオープンソース (MIT ライセンス) です。インストーラーは COM アドインを登録するだけで、システムファイルは変更されません。
</details>

<details>
<summary><b>アンインストール方法は?</b></summary>
<br>
`uninstall.bat` を右クリック → 管理者として実行。完了。
</details>

## 🤝 コントリビューション

[AGENTS.md](AGENTS.md) でコントリビューターガイドラインをご確認ください。

```bash
# コミット規約
feat(scope): 説明
fix(scope): 説明
docs: 説明

# 例
feat(core): w/b/e ワードモーション追加
fix(hooks): デリゲート GC 収集を防止
docs: フェーズ 3 でマイルストーンを更新
```

## 📊 プロジェクトステータス

```
✅ 研究 & アーキテクチャ        ████████████ 100%
✅ PoC: アドイン読み込み        ████████████ 100%
✅ PoC: キーボードフック        ████████████ 100%
✅ PoC: モード切替              ████████████ 100%
🔜 フェーズ 0: 基盤            ░░░░░░░░░░░░   0%
🔜 フェーズ 1: コアエンジン    ░░░░░░░░░░░░   0%
🔜 フェーズ 2: Word 統合       ░░░░░░░░░░░░   0%
🔜 フェーズ 3: コマンド        ░░░░░░░░░░░░   0%
🔜 フェーズ 4: Visual モード   ░░░░░░░░░░░░   0%
🔜 フェーズ 5: v1.0 リリース   ░░░░░░░░░░░░   0%
```

## 📄 ライセンス

[MIT ライセンス](LICENSE) — 無料で使用、変更、配布可能。

## 🙏 謝辞

- [Vim](https://www.vim.org/) — このプロジェクトにインスピレーションを与えた伝説的なテキストエディタ
- [Word Object Model](https://learn.microsoft.com/en-us/office/vba/api/word.word) — Word 自動化のための Microsoft の COM API
- [WH_KEYBOARD Hook](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw) — キーボード傍受のための Win32 API

---

<div align="center">

**Vim ユーザーのために ❤️ でビルド**

[⬆️ トップに戻る](#-wordvim)

</div>
