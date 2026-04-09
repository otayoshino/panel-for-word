// src/components/features/typography/RubyFeature.tsx
import { useEffect, useState } from 'react'
import { Button, Text, makeStyles, tokens, Spinner } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { getTokenizer, textToRubyPairs } from '../../../utils/rubyKuromoji'
import { buildRubyOoxml, containsKanji } from '../../../utils/rubyOoxml'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
  note: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.6',
  },
  noteWarn: {
    fontSize: '11px',
    color: '#b85c00',
    lineHeight: '1.6',
    backgroundColor: '#fff8f0',
    border: '1px solid #f5d0a0',
    borderRadius: '6px',
    padding: '6px 8px',
  },
  statusRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center',
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
  },
})

export function RubyFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [dictLoading, setDictLoading] = useState(false)
  const [dictReady, setDictReady] = useState(false)

  /** コンポーネントマウント時にバックグラウンドで辞書をロード開始 */
  useEffect(() => {
    setDictLoading(true)
    getTokenizer()
      .then(() => {
        setDictReady(true)
      })
      .catch(() => {
        // ロード失敗時は applyRuby 内でエラーを表示する
      })
      .finally(() => {
        setDictLoading(false)
      })
  }, [])

  /** 選択テキストにルビを振って Word に書き戻す */
  const applyRuby = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()

      const text = range.text
      if (!text || text.trim() === '') {
        setStatus({ type: 'warning', message: 'テキストを選択してから実行してください' })
        return
      }
      if (!containsKanji(text)) {
        setStatus({ type: 'warning', message: '選択範囲に漢字が含まれていません' })
        return
      }

      let pairs
      try {
        pairs = await textToRubyPairs(text)
        setDictReady(true)
      } catch (e) {
        setStatus({ type: 'error', message: `辞書読み込みエラー: ${e instanceof Error ? e.message : String(e)}` })
        return
      }

      const ooxml = buildRubyOoxml(pairs)
      range.insertOoxml(ooxml, Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="自動ルビ" />

      <Text className={styles.note}>
        選択したテキストの漢字にルビ（ふりがな）を自動で振ります。
      </Text>

      {!dictReady && (
        <Text className={styles.noteWarn}>
          ⚠ 初回実行時は辞書ファイルの読み込みに20〜30秒かかります。
          読み込み完了後にルビが適用されます。
        </Text>
      )}

      {dictLoading && (
        <div className={styles.statusRow}>
          <Spinner size="tiny" />
          <span>辞書を読み込んでいます...</span>
        </div>
      )}

      <Button
        appearance="primary"
        className={styles.btnFull}
        onClick={applyRuby}
      >
        実行（選択範囲にルビを振る）
      </Button>

      <StatusBar status={status} />
    </div>
  )
}
