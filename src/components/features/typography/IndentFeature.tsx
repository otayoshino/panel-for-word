// src/components/features/typography/IndentFeature.tsx
import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { SectionHeader } from '../../shared/SectionHeader'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: tokens.spacingHorizontalS, width: '100%' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function IndentFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [indentLeft, setIndentLeft] = useState(0)
  const [indentRight, setIndentRight] = useState(0)
  const [indentFirstLine, setIndentFirstLine] = useState(0)

  const applyIndent = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => p.load('font/size'))
      await context.sync()

      const items = paragraphs.items.map((p) => {
        const range = p.getRange('Whole')
        const ooxmlResult = range.getOoxml()
        return { para: p, range, ooxmlResult }
      })
      await context.sync()

      items.forEach(({ para, range, ooxmlResult }) => {
        const charPt = para.font.size || 10.5
        const toTwip = (ch: number) => Math.round(ch * charPt * 20)
        const toCh100 = (ch: number) => Math.round(ch * 100)

        let indTag: string
        if (indentFirstLine >= 0) {
          indTag = [
            `<w:ind`,
            ` w:left="${toTwip(indentLeft)}" w:leftChars="${toCh100(indentLeft)}"`,
            ` w:right="${toTwip(indentRight)}" w:rightChars="${toCh100(indentRight)}"`,
            ` w:firstLine="${toTwip(indentFirstLine)}" w:firstLineChars="${toCh100(indentFirstLine)}"`,
            `/>`,
          ].join('')
        } else {
          const h = -indentFirstLine
          indTag = [
            `<w:ind`,
            ` w:left="${toTwip(indentLeft)}" w:leftChars="${toCh100(indentLeft)}"`,
            ` w:right="${toTwip(indentRight)}" w:rightChars="${toCh100(indentRight)}"`,
            ` w:hanging="${toTwip(h)}" w:hangingChars="${toCh100(h)}"`,
            `/>`,
          ].join('')
        }

        let xml = ooxmlResult.value
        if (/<w:ind[^>]*\/>/s.test(xml)) {
          xml = xml.replace(/<w:ind[^>]*\/>/s, indTag)
        } else if (/<\/w:pPr>/.test(xml)) {
          xml = xml.replace('<\/w:pPr>', indTag + '<\/w:pPr>')
        } else if (/<w:pPr\s*\/>/s.test(xml)) {
          xml = xml.replace(/<w:pPr\s*\/>/s, `<w:pPr>${indTag}<\/w:pPr>`)
        } else {
          xml = xml.replace(/(<w:p(?:\s[^>]*)?>)/s, `$1<w:pPr>${indTag}<\/w:pPr>`)
        }
        range.insertOoxml(xml, 'Replace')
      })
      await context.sync()
    })

  const resetIndent = () =>
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      const items = paragraphs.items.map((p) => {
        const range = p.getRange('Whole')
        const ooxmlResult = range.getOoxml()
        return { range, ooxmlResult }
      })
      await context.sync()
      items.forEach(({ range, ooxmlResult }) => {
        const xml = ooxmlResult.value.replace(/<w:ind[^>]*\/>/s, '')
        range.insertOoxml(xml, 'Replace')
      })
      await context.sync()
      setIndentLeft(0)
      setIndentRight(0)
      setIndentFirstLine(0)
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="インデント" />
      <div className={styles.grid}>
        <Field label="左 (字)">
          <SpinButton value={indentLeft} min={0} max={30} step={0.5} onChange={(_, d) => setIndentLeft(d.value ?? 0)} />
        </Field>
        <Field label="最初の行 (字)">
          <SpinButton value={indentFirstLine} min={-10} max={30} step={0.5} onChange={(_, d) => setIndentFirstLine(d.value ?? 0)} />
        </Field>
        <Field label="右 (字)">
          <SpinButton value={indentRight} min={0} max={30} step={0.5} onChange={(_, d) => setIndentRight(d.value ?? 0)} />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={applyIndent}>選択範囲に適用</Button>
      <Button appearance="secondary" className={styles.btnFull} onClick={resetIndent}>リセット</Button>
      <StatusBar status={status} />
    </div>
  )
}
