<div align="center">

# ⌨️ WordVim

**Microsoft Word용 바인딩 키 바인딩**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)]()
[![Word](https://img.shields.io/badge/Word-365%20Desktop-orange.svg)]()
[![.NET](https://img.shields.io/badge/.NET-Framework%204.7.2-purple.svg)]()
[![Status](https://img.shields.io/badge/Status-Proof%20of%20Concept-brightgreen.svg)]()
[![Release](https://img.shields.io/github/v/release/spidychoipro/wordvim.svg)](https://github.com/spidychoipro/wordvim/releases)

[English](README.md) · **한국어** · [日本語](README.ja.md) · [中文](README.zh.md)

---

**WordVim**은 Microsoft Word에서 Vim의 모달 편집 경험을 제공합니다.

`Escape`를 눌러 Normal 모드로 전환하고, `i`를 눌러 Insert 모드로 돌아갑니다 — Vim과 같습니다.

**Visual Studio가 필요 없습니다.** `dotnet build`로 빌드합니다. 원클릭 설치.

[**⬇️ v0.1.0 다운로드**](https://github.com/spidychoipro/wordvim/releases/tag/v0.1.0) · [**📖 문서**](#-문서) · [**🚀 빠른 시작**](#-빠른-시작)

---

</div>

## 🎯 WordVim을 만드는 이유

Microsoft Word는 Vim의 강력한 모달 편집이 부족합니다. **Office.js는 키보드 입력을 가로채지 못합니다** — 이는 극복할 수 없는 [아키텍처적 한계](https://github.com/nickebbitt/office-dev-journey/issues/6454)입니다.

WordVim은 **네이티브 COM 애드인 + Win32 키보드 훅**을 사용하여 이 문제를 해결합니다 — Word에서 진정한 Vim 모달 편집을 가능하게 하는 유일한 접근 방식입니다.

<div align="center">

```
┌─────────────────────────────────────────────────────────┐
│  Word: -- NORMAL --                                     │
│                                                         │
│  vim → Word 통합                                       │
│                                                         │
│  ✅ 모달 편집 (Insert/Normal/Visual)                    │
│  ✅ Win32 훅을 통한 키보드 가로채기                     │
│  ✅ Word COM API를 통한 커서 이동                      │
│  ✅ 원클릭 설치 — Visual Studio 불필요                 │
└─────────────────────────────────────────────────────────┘
```

</div>

## ✨ 기능

| 기능 | 상태 |
|------|------|
| Insert / Normal 모드 전환 | ✅ 완료 |
| 모드 표시기 (제목 표시줄) | ✅ 완료 |
| Normal 모드에서 키 억제 | ✅ 완료 |
| `h`/`j`/`k`/`l` 커서 이동 | 🔜 다음 |
| `w`/`b`/`e` 단어 이동 | 🔜 계획 |
| `x`/`dd`/`yy` 편집 명령 | 🔜 계획 |
| Visual 모드 | 🔜 계획 |
| 명령줄 모드 (`:`) | 🔜 계획 |

## 🚀 빠른 시작

### 옵션 1: 다운로드 (권장)

> **1단계:** [`wordvim-v0.1.0.zip`](https://github.com/spidychoipro/wordvim/releases/download/v0.1.0/wordvim-v0.1.0.zip) 다운로드

> **2단계:** ZIP 파일을 아무 곳에나 압축 해제

> **3단계:** `install.bat` 우클릭 → **관리자 권한으로 실행**

> **4단계:** Microsoft Word 열기 → 제목 표시줄에 `-- NORMAL --` 확인 🎉

```
📁 wordvim-v0.1.0.zip
   └── 📁 wordvim
       ├── 📄 install.bat      ← 이 파일을 더블클릭 (관리자 권한)
       ├── 📄 uninstall.bat
       ├── 📄 register.ps1
       ├── 📄 unregister.ps1
       ├── 📄 PoC.dll
       └── 📄 PoC.pdb
```

### 옵션 2: 소스에서 빌드

```bash
# 필수: .NET SDK 8.0+ 및 Git
git clone https://github.com/spidychoipro/wordvim.git
cd wordvim/poc
dotnet build

# 등록 (PowerShell을 관리자 권한으로 실행)
.\register.ps1
```

## 📖 문서

| 문서 | 설명 |
|------|------|
| [아키텍처 분석](docs/architecture.md) | COM 애드인 + Win32 훅을 선택한 이유, VSTO/Office.js가 아닌 이유 |
| [API 연구](docs/research.md) | Word COM API 기능 및 한계 |
| [MVP 사양](docs/mvp.md) | 최소 기능 제품에 포함된 내용 |
| [마일스톤](docs/milestones.md) | 30일 개발 로드맵 |
| [PoC 계획](docs/poc-plan.md) | 개념 증명 검증 계획 |
| [환경 가이드](docs/environment.md) | 개발 환경 설정 |
| [기여 가이드](AGENTS.md) | 기여자 가이드라인 및 코드 규칙 |

## 🏗️ 아키텍처

```
┌──────────────────────────────────────────────────────────────┐
│                    WordVim 아키텍처                           │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐     │
│  │   Win32     │    │   Word      │    │   Vim       │     │
│  │   훅       │───▶│   COM API   │───▶│   상태      │     │
│  │  (키보드)   │    │  (커서)     │    │   머신     │     │
│  └─────────────┘    └─────────────┘    └─────────────┘     │
│         │                  │                   │             │
│         ▼                  ▼                   ▼             │
│  ┌─────────────────────────────────────────────────────┐    │
│  │           COM 애드인 (IDTExtensibility2)            │    │
│  │                WordVim.Connect                      │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

| 구성 요소 | 기술 | 용도 |
|-----------|------|------|
| **애드인** | COM + IDTExtensibility2 | Word 프로세스에 로드 |
| **키보드** | WH_KEYBOARD 훅 | Word보다 먼저 키보드 입력 가로채기 |
| **커서** | Word COM API | 커서 이동, 텍스트 편집 |
| **상태** | C# 상태 머신 | Insert/Normal 모드 추적 |

## 🎮 사용법

### 키 바인딩

| 키 | 모드 | 동작 |
|----|------|------|
| `Escape` | Insert → Normal | Normal 모드로 전환 |
| `i` | Normal → Insert | Insert 모드로 전환 |
| `h`/`j`/`k`/`l` | Normal | 커서 이동 *(계획)* |
| `w`/`b`/`e` | Normal | 단어 이동 *(계획)* |
| `x` | Normal | 문자 삭제 *(계획)* |
| `dd` | Normal | 줄 삭제 *(계획)* |
| `yy` | Normal | 줄 복사 *(계획)* |

### 모드 표시기

윈도우 제목 표시줄에 현재 모드가 표시됩니다:

```
-- NORMAL --  ← 'i'를 제외한 모든 키가 억제됨
-- INSERT --  ← 일반 타이핑 모드
```

## ❓ 자주 묻는 질문

<details>
<summary><b>Word Online에서 작동하나요?</b></summary>
<br>
아니요. WordVim은 Windows의 데스크톱 버전 Word 365가 필요합니다. Word Online은 COM 애드인을 지원하지 않는 다른 아키텍처를 사용합니다.
</details>

<details>
<summary><b>Visual Studio가 필요한가요?</b></summary>
<br>
아니요. WordVim은 `dotnet build`로 빌드됩니다 — .NET SDK만 설치하면 됩니다.
</details>

<details>
<summary><b>안전한가요?</b></summary>
<br>
예. WordVim은 오픈 소스(MIT 라이선스)입니다. 설치 프로그램은 COM 애드인만 등록합니다 — 시스템 파일은 수정되지 않습니다.
</details>

<details>
<summary><b>제거하려면 어떻게 하나요?</b></summary>
<br>
`uninstall.bat` 우클릭 → 관리자 권한으로 실행. 끝.
</details>

## 🤝 기여

[AGENTS.md](AGENTS.md)에서 기여 가이드라인을 확인하세요.

```bash
# 커밋 규칙
feat(scope): 설명
fix(scope): 설명
docs: 설명

# 예시
feat(core): w/b/e 단어 이동 추가
fix(hooks): 대리자 GC 수집 방지
docs: 3단계로 마일스톤 업데이트
```

## 📊 프로젝트 상태

```
✅ 연구 및 아키텍처          ████████████ 100%
✅ PoC: 애드인 로딩          ████████████ 100%
✅ PoC: 키보드 훅            ████████████ 100%
✅ PoC: 모드 전환            ████████████ 100%
🔜 0단계: 기반              ░░░░░░░░░░░░   0%
🔜 1단계: 코어 엔진         ░░░░░░░░░░░░   0%
🔜 2단계: Word 통합         ░░░░░░░░░░░░   0%
🔜 3단계: 명령              ░░░░░░░░░░░░   0%
🔜 4단계: Visual 모드       ░░░░░░░░░░░░   0%
🔜 5단계: v1.0 릴리스       ░░░░░░░░░░░░   0%
```

## 📄 라이선스

[MIT 라이선스](LICENSE) — 무료로 사용, 수정, 배포 가능.

## 🙏 감사의 글

- [Vim](https://www.vim.org/) — 이 프로젝트에 영감을 준 전설적인 텍스트 에디터
- [Word Object Model](https://learn.microsoft.com/en-us/office/vba/api/word.word) — Word 자동화를 위한 Microsoft의 COM API
- [WH_KEYBOARD Hook](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowshookexw) — 키보드 가로채기를 위한 Win32 API

---

<div align="center">

**Vim 사용자를 위해 ❤️로 빌드됨**

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-ffdd00?style=flat&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/spidychoi)

[⬆️ 맨 위로](#-wordvim)

</div>
