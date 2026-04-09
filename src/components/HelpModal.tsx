// src/components/HelpModal.tsx
// ヘルプモーダル — 使い方・FAQ・用語集・トラブルシューティング・お問い合わせ

import { useState } from 'react'
import { createPortal } from 'react-dom'
import { Text, makeStyles, tokens } from '@fluentui/react-components'

const useStyles = makeStyles({
  overlay: {
    position: 'fixed',
    inset: '0',
    backgroundColor: 'rgba(0,0,0,0.45)',
    zIndex: 10000,
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'center',
    paddingTop: '40px',
  },
  modal: {
    backgroundColor: '#ffffff',
    borderRadius: '8px',
    boxShadow: '0 8px 32px rgba(0,0,0,0.24)',
    width: '340px',
    maxWidth: '95vw',
    maxHeight: '80vh',
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  modalHeader: {
    backgroundColor: '#1e4d8c',
    padding: '10px 14px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexShrink: 0,
  },
  modalTitle: {
    color: '#ffffff',
    fontSize: '13px',
    fontWeight: '600',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  closeBtn: {
    background: 'transparent',
    border: 'none',
    color: '#ffffff',
    cursor: 'pointer',
    fontSize: '16px',
    lineHeight: '1',
    padding: '2px 6px',
    borderRadius: '4px',
    ':hover': { backgroundColor: 'rgba(255,255,255,0.2)' },
  },
  tabBar: {
    display: 'flex',
    borderBottom: '1px solid #c5dcf5',
    backgroundColor: '#f0f4fa',
    flexShrink: 0,
    overflowX: 'auto',
  },
  tab: {
    padding: '7px 10px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    background: 'transparent',
    border: 'none',
    borderBottom: '2px solid transparent',
    cursor: 'pointer',
    whiteSpace: 'nowrap',
    color: '#4a7cb5',
    ':hover': { color: '#1e4d8c', backgroundColor: '#e0ebf8' },
  },
  tabActive: {
    padding: '7px 10px',
    fontSize: '11px',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    background: 'transparent',
    border: 'none',
    borderBottom: '2px solid #1e4d8c',
    cursor: 'pointer',
    whiteSpace: 'nowrap',
    color: '#1e4d8c',
    fontWeight: '600',
  },
  body: {
    padding: '14px',
    overflowY: 'auto',
    flexGrow: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  sectionTitle: {
    fontSize: '12px',
    fontWeight: '700',
    color: '#0c51a0',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    borderLeft: '3px solid #1e4d8c',
    paddingLeft: '6px',
  },
  p: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.7',
  },
  step: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.7',
    paddingLeft: '4px',
  },
  faqQ: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#1e4d8c',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  faqA: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.7',
    paddingLeft: '8px',
  },
  termRow: {
    display: 'grid',
    gridTemplateColumns: '90px 1fr',
    gap: '4px',
    alignItems: 'baseline',
  },
  termKey: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#0c51a0',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  termVal: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.6',
  },
  errBox: {
    backgroundColor: '#fff8f0',
    border: '1px solid #f5d0a0',
    borderRadius: '6px',
    padding: '8px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  errLabel: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#b85c00',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  errText: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.6',
  },
  contactRow: {
    display: 'flex',
    gap: '8px',
    alignItems: 'flex-start',
  },
  contactIcon: {
    fontSize: '16px',
    flexShrink: 0,
    marginTop: '1px',
  },
  contactInfo: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  contactLabel: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#0c51a0',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
  },
  contactText: {
    fontSize: '11px',
    color: '#333',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.6',
  },
  divider: {
    borderTop: '1px solid #e0ebf8',
    margin: '4px 0',
  },
})

// ── タブコンテンツ定義 ────────────────────────────────────────────────────────

function GuideTab() {
  const styles = useStyles()
  return (
    <>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>主要機能の概要</Text>
        <Text className={styles.p}>
          かんたんツールボックスは、Wordの文書作成をサポートするアドインです。
          基本設定・文字組・数式入力などの機能をタブで切り替えて使用できます。
        </Text>
      </div>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>基本的な使い方</Text>
        <Text className={styles.step}>① 上部タブから操作したいカテゴリを選択</Text>
        <Text className={styles.step}>② カードをクリックして機能の設定画面を開く</Text>
        <Text className={styles.step}>③ 値を設定して「実行」ボタンをクリック</Text>
        <Text className={styles.step}>④「← 戻る」で一覧へ戻る</Text>
      </div>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>お気に入り機能</Text>
        <Text className={styles.p}>
          各カードの設定画面右上の ★ ボタンでお気に入り登録できます。
          ★タブでお気に入りのカードのみを表示でき、ドラッグ＆ドロップで並び替えが可能です。
        </Text>
      </div>
    </>
  )
}

function FaqTab() {
  const styles = useStyles()
  const faqs = [
    {
      q: 'アドインが動作しない場合は？',
      a: 'Wordを再起動し、アドインを再度サイドロードしてください。問題が続く場合はブラウザのキャッシュをクリアしてください。',
    },
    {
      q: '段組みを設定してもエラーが出る',
      a: '段组みの設定はWordApiDesktop 1.3が必要です。Microsoft 365 Appsの最新版にアップデートしてください。',
    },
    {
      q: 'ルビが正しく付かない',
      a: '自動ルビには形態素解析辞書の読み込みが必要です。初回使用時に少し時間がかかることがあります。',
    },
    {
      q: 'お気に入りの並び順がリセットされた',
      a: 'お気に入りはブラウザのlocalStorageに保存されています。ブラウザのデータをクリアすると初期化されます。',
    },
    {
      q: '数式が挿入されない',
      a: '数式はカーソルが文書内にある状態で実行してください。ヘッダー・フッター内では動作しない場合があります。',
    },
  ]
  return (
    <>
      {faqs.map((faq, i) => (
        <div key={i} className={styles.section}>
          <Text className={styles.faqQ}>Q. {faq.q}</Text>
          <Text className={styles.faqA}>A. {faq.a}</Text>
          {i < faqs.length - 1 && <div className={styles.divider} />}
        </div>
      ))}
    </>
  )
}

function GlossaryTab() {
  const styles = useStyles()
  const terms = [
    { key: 'サイドロード', val: 'マニフェストファイルを使ってアドインをWordに手動で読み込む方法。開発・テスト時に使用します。' },
    { key: 'ContentControl', val: 'Wordの「コンテンツコントロール」。テキスト枠を作成する際に使用します。' },
    { key: 'OOXML', val: 'Office Open XMLの略。Word文書の内部フォーマットで、数式挿入などに使用します。' },
    { key: 'textColumns', val: 'Wordの段組み設定を制御するAPIオブジェクト。段数・間隔を変更できます。' },
    { key: 'モノルビ', val: '漢字1文字ずつにルビを付ける方式。熟語全体にまとめて付ける「グループルビ」と対比されます。' },
    { key: '★アイコン', val: 'お気に入り登録ボタン。黄色で塗りつぶされている場合は登録済みです。' },
  ]
  return (
    <>
      {terms.map((term, i) => (
        <div key={i} className={styles.section}>
          <div className={styles.termRow}>
            <Text className={styles.termKey}>{term.key}</Text>
            <Text className={styles.termVal}>{term.val}</Text>
          </div>
          {i < terms.length - 1 && <div className={styles.divider} />}
        </div>
      ))}
    </>
  )
}

function TroubleTab() {
  const styles = useStyles()
  const items = [
    {
      label: 'ホワイトアウト（真っ白になる）',
      text: '開発サーバーのURLとマニフェストのURLが一致しているか確認してください。manifest-dev.xml が https://localhost:3000 を指定しているか確認してください。',
    },
    {
      label: '「値が有効範囲を超えています」エラー',
      text: 'Wordに設定できる値の範囲を超えています。文字数は1〜200、行数は1〜200の範囲で入力してください。',
    },
    {
      label: '段組み設定でエラーが表示される',
      text: 'WordApiDesktop 1.3が必要です。Microsoft 365 Appsを最新版にアップデートすることで解消されます。',
    },
    {
      label: 'ルビ辞書の読み込みが遅い',
      text: 'kuromoji辞書ファイルの読み込みに時間がかかることがあります。しばらく待って再試行してください。',
    },
    {
      label: 'フォント一覧が空になる',
      text: 'ドキュメントが空の場合、使用フォントが取得できません。テキストが入力された状態で実行してください。',
    },
  ]
  return (
    <>
      {items.map((item, i) => (
        <div key={i} className={styles.errBox} style={{ marginBottom: i < items.length - 1 ? '8px' : 0 }}>
          <Text className={styles.errLabel}>⚠ {item.label}</Text>
          <Text className={styles.errText}>{item.text}</Text>
        </div>
      ))}
    </>
  )
}

function ContactTab() {
  const styles = useStyles()
  return (
    <>
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>サポートへのお問い合わせ</Text>
        <Text className={styles.p}>
          ご不明な点・ご要望は以下の方法でお問い合わせください。サンプル情報のため、実際の連絡先は後日更新されます。
        </Text>
      </div>
      <div className={styles.section}>
        <div className={styles.contactRow}>
          <span className={styles.contactIcon}>✉</span>
          <div className={styles.contactInfo}>
            <Text className={styles.contactLabel}>メール</Text>
            <Text className={styles.contactText}>support@example.com</Text>
            <Text className={styles.contactText}>（受付時間：平日 9:00〜18:00）</Text>
          </div>
        </div>
        <div className={styles.divider} />
        <div className={styles.contactRow}>
          <span className={styles.contactIcon}>💬</span>
          <div className={styles.contactInfo}>
            <Text className={styles.contactLabel}>チャットサポート</Text>
            <Text className={styles.contactText}>公式サイトのチャットウィジェットからお問い合わせください。</Text>
            <Text className={styles.contactText}>（受付時間：平日 10:00〜17:00）</Text>
          </div>
        </div>
        <div className={styles.divider} />
        <div className={styles.contactRow}>
          <span className={styles.contactIcon}>📞</span>
          <div className={styles.contactInfo}>
            <Text className={styles.contactLabel}>電話</Text>
            <Text className={styles.contactText}>00-0000-0000</Text>
            <Text className={styles.contactText}>（受付時間：平日 9:00〜17:00）</Text>
          </div>
        </div>
      </div>
    </>
  )
}

// ── メインコンポーネント ──────────────────────────────────────────────────────

type HelpTabId = 'guide' | 'faq' | 'glossary' | 'trouble' | 'contact'

const HELP_TABS: { id: HelpTabId; label: string }[] = [
  { id: 'guide',    label: '使い方' },
  { id: 'faq',      label: 'FAQ' },
  { id: 'glossary', label: '用語集' },
  { id: 'trouble',  label: 'トラブル' },
  { id: 'contact',  label: 'お問い合わせ' },
]

interface HelpModalProps {
  onClose: () => void
}

export function HelpModal({ onClose }: HelpModalProps) {
  const styles = useStyles()
  const [activeTab, setActiveTab] = useState<HelpTabId>('guide')

  const renderContent = () => {
    switch (activeTab) {
      case 'guide':    return <GuideTab />
      case 'faq':      return <FaqTab />
      case 'glossary': return <GlossaryTab />
      case 'trouble':  return <TroubleTab />
      case 'contact':  return <ContactTab />
    }
  }

  return createPortal(
    <div
      className={styles.overlay}
      onClick={(e) => { if (e.target === e.currentTarget) onClose() }}
    >
      <div className={styles.modal}>
        <div className={styles.modalHeader}>
          <Text className={styles.modalTitle}>ヘルプ</Text>
          <button className={styles.closeBtn} onClick={onClose} aria-label="閉じる">✕</button>
        </div>
        <div className={styles.tabBar}>
          {HELP_TABS.map((tab) => (
            <button
              key={tab.id}
              className={activeTab === tab.id ? styles.tabActive : styles.tab}
              onClick={() => setActiveTab(tab.id)}
            >
              {tab.label}
            </button>
          ))}
        </div>
        <div className={styles.body}>
          {renderContent()}
        </div>
      </div>
    </div>,
    document.body
  )
}
