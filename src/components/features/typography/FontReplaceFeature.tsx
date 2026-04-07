// src/components/features/typography/FontReplaceFeature.tsx
import { useState } from 'react'
import { Button, Field, Input, Text, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  fontListRow: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'flex-start', width: '100%' },
  fontList: {
    flex: 1,
    minHeight: '80px',
    maxHeight: '120px',
    overflowY: 'auto',
    overflowX: 'hidden',
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalS,
  },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function FontReplaceFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [fontList, setFontList] = useState<string[]>([])
  const [fromFont, setFromFont] = useState('')
  const [toFont, setToFont] = useState('')

  const collectFonts = () =>
    runWord(async (context) => {
      const body = context.document.body
      const paragraphs = body.paragraphs
      paragraphs.load('items')
      await context.sync()
      const tasks = paragraphs.items.map((p) => { p.load('font/name'); return p })
      await context.sync()
      const fonts = new Set<string>()
      tasks.forEach((p) => { if (p.font.name) fonts.add(p.font.name) })
      setFontList(Array.from(fonts).sort())
    })

  const replaceFont = () =>
    runWord(async (context) => {
      if (!fromFont || !toFont) {
        setStatus({ type: 'warning', message: '変換元と変換先のフォント名を入力してください' })
        return
      }
      const results = context.document.body.search('*', { matchWildcards: true })
      results.load('items')
      await context.sync()
      results.items.forEach((r) => r.load('font/name'))
      await context.sync()
      results.items.forEach((r) => { if (r.font.name === fromFont) r.font.name = toFont })
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="ドキュメント使用フォント一覧・置換" />
      <div className={styles.fontListRow}>
        <div className={styles.fontList}>
          {fontList.map((f) => <Text key={f} size={200} block>{f}</Text>)}
        </div>
        <Button appearance="secondary" onClick={collectFonts}>取得</Button>
      </div>
      <Field label="変換元フォント">
        <Input
          value={fromFont}
          onChange={(_, d) => setFromFont(d.value)}
          placeholder="例: MS 明朝"
          list="font-list-datalist"
        />
        {fontList.length > 0 && (
          <datalist id="font-list-datalist">
            {fontList.map((f) => <option key={f} value={f} />)}
          </datalist>
        )}
      </Field>
      <Field label="変換先フォント">
        <Input value={toFont} onChange={(_, d) => setToFont(d.value)} placeholder="例: 游明朝" />
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={replaceFont}>
        フォント置換
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
