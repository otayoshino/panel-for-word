---
applyTo: "src/**"
description: |
  panel-for-word モック実装ガイド。
  Office.js Word Add-in（React + TypeScript + Fluent UI）のコーディング規約・実装方針を定める。
  Use when: implementing Word panel features, creating React components, using Office.js API, Fluent UI components.
---

# Word パネル モック 実装ガイド

## プロジェクト概要

VBA 製 Word カスタム作業ウィンドウを Office.js Add-in（React + TypeScript + Vite）として作り替えるプロジェクト。

---

## 機能対応状況（VBA_vs_OfficeJS_比較表.xlsx より）

凡例: ○ 対応可 / △ 部分対応 / × 非対応（廃止）

### 〈基本設定〉ページ設定

| No | 機能 | 判定 | Office.js API / 備考 |
|----|------|------|----------------------|
| 1 | 現在のドキュメントの設定値 | ○ | `context.document.body.style` 等で取得可 |
| 2 | 用紙サイズパレット | △ | ダイアログ呼び出し不可。値設定はAPI経由で可 |
| 3 | 用紙サイズ（プルダウン選択） | ○ | `body.style.pageWidth/Height` |
| 4 | 横組み・縦組み | △ | 段落レベルの `textDirection` は可。セクション縦組みは制限あり |
| 5 | 基本日本語フォントパレット | △ | フォントウィンドウ呼び出し不可。フォント名の取得・設定はAPI経由可 |
| 6 | 基本文字サイズ | ○ | `font.size` |
| 7 | 段組み | △ | `columnCount/columnSpacing` 設定可。GUIパレット表示は不可 |
| 8 | 文字数（字送り変更） | △ | `characterSpacing` で近似可。ページ設定の文字数指定は直接APIなし |
| 9 | 行数（行送り変更） | △ | `lineSpacing` で近似可。行数指定はAPI非対応 |
| 10 | ページ設定パレット | × | Wordダイアログ直接呼び出し不可 |
| 11 | ページ余白（ミリ）表示 | ○ | `pageSetup.topMargin` 等で取得・表示可 |
| 12 | 実行（余白設定適用） | ○ | `topMargin/bottomMargin/leftMargin/rightMargin` |
| 13 | 余白パレット | × | Wordダイアログ直接呼び出し不可 |
| 14 | Wordのオプション | × | Wordアプリ設定ダイアログは呼び出し不可 |
| 15 | オートコレクト | × | オートコレクトダイアログはAPI非公開 |
| 16 | 入力フォーマット | × | オートコレクトダイアログはAPI非公開 |

### 〈文字組み〉インデント

| No | 機能 | 判定 | Office.js API |
|----|------|------|---------------|
| 17 | 左インデント | ○ | `paragraph.leftIndent` |
| 18 | 最初の行インデント | ○ | `paragraph.firstLineIndent` |
| 19 | 右インデント | ○ | `paragraph.rightIndent` |
| 20 | 全て0（リセット） | ○ | 各プロパティを0にセット |
| 21 | 段落パレットを表示 | × | Wordダイアログ直接呼び出し不可 |

### 〈文字組み〉行間

| No | 機能 | 判定 | Office.js API |
|----|------|------|---------------|
| 22 | 倍数 | ○ | `paragraph.lineSpacing / lineSpacingRule` |
| 23 | 固定値 | ○ | `paragraph.lineSpacingRule = Fixed` |

### 〈文字組み〉タブとリーダー

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 24 | タブパレット表示 | × | ダイアログ呼び出し不可。`tabStops` APIで個別設定は可能 |

### 〈文字組み〉表

| No | 機能 | 判定 | Office.js API / 備考 |
|----|------|------|----------------------|
| 25 | 表の挿入 | ○ | `context.document.body.insertTable()` |
| 26 | 表パレット表示 | × | 表のプロパティダイアログ呼び出し不可。API経由でプロパティ取得・設定は可 |
| 27 | 罫線 | × | 罫線ダイアログ呼び出し不可。`border` APIで線種・色設定は可 |
| 28 | 編みかけ | × | ダイアログ呼び出し不可。`shadingColor/pattern` のAPI設定は可 |

### 〈文字組み〉ドキュメント使用フォント一覧・置換

| No | 機能 | 判定 | Office.js API / 備考 |
|----|------|------|----------------------|
| 29 | フォントリスト取得 | △ | 全テキストを走査してフォント名を収集する実装が必要（直接APIなし） |
| 30 | フォント置換 | ○ | `body.search()` + `font.name` 変更 |

### 〈文字組み〉ルビ機能

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 31 | 自動ルビ | × | Office.js にルビ（phonetic guide）APIなし |
| 32 | ルビ解除 | × | 同上 |
| 33 | ルビ入力（任意） | × | Word の phonetic guide は Office.js 非公開API |

### 〈枠〉画像・オブジェクト挿入

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 34 | 画像・オブジェクトの挿入 | △ | `insertInlinePictureFromBase64` で画像挿入可。図形挿入APIは限定的 |

### 〈枠〉テキスト・画像枠作成

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 35 | ①文字数指定（字・行） | △ | テキストボックス直接挿入APIなし。`ContentControl` で代替可能（制約あり） |
| 36 | ②サイズ指定（横・縦ミリ） | △ | 同上 |
| 37 | 枠の種類 / テキスト枠 | △ | `ContentControl` を代替として使用可。レイアウト自由度は制限 |
| 38 | 枠の種類 / 画像枠 | × | 空の画像枠（プレースホルダー）作成APIなし |
| 39 | 枠の種類 / 図形 | × | 図形挿入リボン表示はAPI経由不可。図形自体の挿入APIも未対応 |
| 40 | 文字列の折り返しパレット | × | レイアウトダイアログ呼び出し不可 |
| 41 | サイズパレット | × | 同上 |
| 42 | テキスト余白 | △ | InlinePicture/Shape の margin APIは限定的 |
| 43 | 設定を表示 | △ | 取得可能なオブジェクトは限定的 |
| 44 | 全て0（リセット） | △ | 取得・設定できる場合のみ |
| 45 | 図形書式設定 | × | リボンタブの切り替えAPIなし |

### 〈枠〉重ね順・枠揃え

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 46–49 | 重ね順（最前面〜最背面） | × | Shape の zOrder は Office.js では未対応（Excelのみ一部対応） |
| 50–52 | 枠揃え（左・中央・右） | × | Shape の位置設定API（Word）非対応 |

### 〈数式〉数式入力

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 53–64 | 数式エディタパレット各種 | × | 数式リボン個別コマンドの呼び出しAPIなし |
| 65–72 | 記号パレット各種 / インク数式 | × | 同上（ギリシャ文字テキスト挿入のみ可） |
| 73 | 選択範囲を変換（テキスト→数式） | × | OfficeMath の挿入APIは Office.js 未対応 |
| 74–78 | # $ % & @ の文字入力 | ○ | `range.insertText('#')` 等で可 |

### 〈定型文〉定型文入力

| No | 機能 | 判定 | Office.js API |
|----|------|------|---------------|
| 78 | 定型文を入力（1〜5） | ○ | `Office.context.document.settings` でテキスト保存・`insertText` で挿入 |
| 79 | 文章よりコピー登録 | ○ | `selection.getTextRanges()` で取得し `Settings` に保存 |
| 80 | 実行（定型文挿入） | ○ | 保存テキストを `insertText` で挿入 |
| 81 | 記号入力 | ○ | UI + `Settings` で実現可 |
| 82 | 記号変更 | ○ | `Settings` の更新で対応可 |
| 83 | 実行（記号順次入力） | ○ | インデックス管理 + `insertText` で実現可 |
| 84 | 記号リセット | ○ | `Settings` をデフォルト値に戻す |

### 〈ウィンドウ下部〉メニュー

| No | 機能 | 判定 | 備考 |
|----|------|------|------|
| 85 | 閉じる | ○ | `Office.context.ui` / taskpane の非表示 |
| 86 | ズーム | × | WordのズームAPIは Office.js 非対応 |
| 87 | ミニツールバー | × | Word組み込みミニツールバーの制御API非対応 |

---

## 実装対象サマリー（○ 対応可 および △ 部分対応）

### ○ 対応可（完全実装）

| カテゴリ | 機能 | Office.js API |
|----------|------|---------------|
| ページ設定 | 現在ドキュメント設定値取得 | `document.body.style` |
| ページ設定 | 用紙サイズ設定 | `body.style.pageWidth/Height` |
| ページ設定 | 基本文字サイズ | `font.size` |
| ページ設定 | ページ余白表示・設定 | `pageSetup.topMargin` 等 |
| インデント | 左・右・最初の行・全て0 | `paragraph.leftIndent/rightIndent/firstLineIndent` |
| 行間 | 倍数・固定値 | `paragraph.lineSpacing/lineSpacingRule` |
| 表 | 表の挿入 | `body.insertTable()` |
| フォント置換 | フォント名の検索・置換 | `body.search()` + `font.name` |
| 文字挿入 | # $ % & @ 入力 | `range.insertText()` |
| 定型文 | 定型文保存・挿入・記号管理 | `Office.context.document.settings` + `insertText` |
| メニュー | パネルを閉じる | `Office.context.ui` |

### △ 部分対応（制約付き実装）

| カテゴリ | 機能 | 代替手段 |
|----------|------|---------|
| ページ設定 | 用紙サイズパレット表示 | 値の設定のみ実装 |
| ページ設定 | 横組み・縦組み | `paragraph.textDirection`（段落レベルのみ） |
| ページ設定 | 基本日本語フォント | フォント名選択UIとAPI設定 |
| ページ設定 | 段組み | `columnCount/columnSpacing` |
| ページ設定 | 文字数・行数 | `characterSpacing/lineSpacing` で近似 |
| フォント一覧 | ドキュメント使用フォント取得 | 全テキスト走査で収集 |
| 画像挿入 | 画像の挿入 | `insertInlinePictureFromBase64` |
| テキスト枠 | テキスト枠作成 | `ContentControl` で代替 |

---

## 非対応機能（廃止）

以下は Office.js Word API で実現不可のため廃止:

- **ルビ（モノルビ）**: phonetic guide API なし
- **Wordダイアログ直接呼び出し**: ページ設定・段落・タブ・表プロパティ・罫線等のダイアログ全般
- **数式リボン操作**: 数式エディタパレット・記号パレット全般、OfficeMath挿入
- **図形挿入・操作**: 図形・画像枠の作成、重ね順、枠揃え（Shape API Word非対応）
- **ズーム・ミニツールバー**: Word UI 制御系 API 非対応
- **Wordアプリ設定**: オプション・オートコレクトダイアログ

---

## コーディング規約

### Office.js の使い方

```typescript
// ✅ 正しいパターン: Word.run + context.sync
const runWord = async (action: (context: Word.RequestContext) => Promise<void>) => {
  try {
    await Word.run(async (context) => {
      await action(context)
    })
  } catch (e) {
    setStatus({ type: 'error', message: `エラー: ${e instanceof Error ? e.message : String(e)}` })
  }
}

// ✅ プロパティ読み込みパターン
const getSelection = () =>
  runWord(async (context) => {
    const range = context.document.getSelection()
    range.load('text')         // 読み取るプロパティを宣言
    await context.sync()       // サーバーと同期
    console.log(range.text)    // sync 後にプロパティにアクセス
  })

// ✅ 書き込みパターン（load 不要）
const applyBold = () =>
  runWord(async (context) => {
    const range = context.document.getSelection()
    range.font.bold = true
    await context.sync()
  })
```

### ステータス管理

```typescript
type Status = { type: 'success' | 'error' | 'warning'; message: string }

// 成功・エラー・警告を統一的に管理
const [status, setStatus] = useState<Status | null>(null)
```

### UIコンポーネント

- **必須**: `@fluentui/react-components` を使用する（HTML ネイティブ要素を直接使わない）
- **スタイル**: `makeStyles` + `tokens` を使用（インラインスタイル禁止）
- **アイコン**: `@fluentui/react-icons` から選択

### コンポーネント分割方針

VBA アプリのタブ構成（基本設定 / 文字組 / 枠 / 数式 / 定型文）に対応させる。

```
src/
  App.tsx                        # ルートコンポーネント（TabList + レイアウト）
  components/
    tabs/
      BasicSettingsTab.tsx       # 基本設定タブ（ページ設定・用紙・余白・文字サイズ等）
      CharCompositionTab.tsx     # 文字組タブ（インデント・行間・表・フォント置換）
      FrameTab.tsx               # 枠タブ（画像挿入・ContentControl枠作成）
      FormulaTab.tsx             # 数式タブ（# $ % & @ 入力のみ実装）
      TemplateTextTab.tsx        # 定型文タブ（定型文・記号管理）
    shared/
      StatusBar.tsx              # ステータス表示（MessageBar）
      SectionHeader.tsx          # セクションヘッダー（Divider + Text）
  hooks/
    useWordRun.ts                # Word.run ラッパーフック
```

### useWordRun フック

```typescript
// hooks/useWordRun.ts
import { useState } from 'react'

type Status = { type: 'success' | 'error'; message: string }

export function useWordRun() {
  const [status, setStatus] = useState<Status | null>(null)

  const runWord = async (action: (context: Word.RequestContext) => Promise<void>) => {
    try {
      await Word.run(async (context) => {
        await action(context)
      })
    } catch (e) {
      setStatus({
        type: 'error',
        message: `エラー: ${e instanceof Error ? e.message : String(e)}`,
      })
    }
  }

  return { runWord, status, setStatus }
}
```

---

## UIレイアウト方針

- パネルは縦スクロール可能な単一カラム
- **最上部**: タブ（基本設定 / 文字組 / 枠 / 数式 / 定型文）
- タブ内の各機能セクションは `<Divider>` で区切る
- セクションヘッダーは `【セクション名】` 形式の `<Text weight="semibold">`（VBIツールボックスに準拠）
- ボタン行は `flexWrap: 'wrap'` でレスポンシブに
- パネル幅: 最小 280px、パディング `tokens.spacingHorizontalM`
- **最下部**: ステータスバー（MessageBar 固定表示）

### タブ別セクション構成

**基本設定タブ**
1. 【ページ設定】— ドキュメント設定値取得、用紙サイズ、横組方向
2. 【文字サイズ・フォント】— 基本文字サイズ（SpinButton）
3. 【段組み】— 段数（SpinButton）
4. 【余白設定（ミリ）】— 上・下・左・右（Input + 実行ボタン）

**文字組タブ**
1. 【インデント】— 左・右・最初の行（SpinButton）+ 全て0
2. 【行間】— 倍数・固定値（SpinButton + RadioGroup）
3. 【表】— 表の挿入ボタン
4. 【ドキュメント使用フォント一覧・置換】— フォントリスト取得 + 置換 Dropdown + 実行

**枠タブ**
1. 【画像挿入】— ファイル選択 + 挿入ボタン
2. 【テキスト枠作成（ContentControl）】— サイズ指定 + 作成ボタン

**数式タブ**
1. 【記号入力】— # $ % & @ ボタン（`range.insertText` で挿入）
2. ※数式リボン系は非対応のため廃止（旨をUI上に表示）

**定型文タブ**
1. 【定型文入力】— 定型文1〜5（Input + ラジオ + 文章からコピー登録 + 実行）
2. 【記号入力】— 記号表示・変更・実行・リセット

---

## Fluent UI コンポーネント対応表

| 用途 | コンポーネント |
|------|-------------|
| アクションボタン | `<Button appearance="secondary">` / `<Button appearance="primary">` |
| 数値入力 | `<SpinButton>` |
| テキスト入力 | `<Input>` |
| ドロップダウン | `<Select>` または `<Dropdown>` |
| ラベル付き入力 | `<Field label="..."><Input /></Field>` |
| セクション区切り | `<Divider>` |
| ステータス表示 | `<MessageBar intent="success|error|warning">` |
| ツールチップ | `<Tooltip content="...">` |
| アコーディオン | `<Accordion>` （セクション折りたたみに使用可） |

---

## TypeScript 規約

- `strict: true` 必須
- `any` 型禁止（`unknown` → 型ガードで絞る）
- Office.js 型は `@types/office-js` から参照（グローバル `Word` 名前空間）
- コンポーネント Props は `interface` で定義

---

## 禁止事項

- `document.getElementById` など DOM 直接操作
- `console.log` を製品コードに残す（デバッグ後削除）
- インラインスタイル（`style={{ ... }}`）
- Office API 呼び出しを `Word.run` 外で行う
