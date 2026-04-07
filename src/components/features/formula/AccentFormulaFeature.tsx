// src/components/features/formula/AccentFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type AccentType = 'vec' | 'overlineABC' | 'overlineXOR'
const ACCENT_TYPES: { value: AccentType; label: string }[] = [
  { value: 'vec',         label: 'ベクトル A' },
  { value: 'overlineABC', label: 'オーバーライン付き ABC' },
  { value: 'overlineXOR', label: 'オーバーライン付き x XOR y' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function AccentFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [accentType, setAccentType] = useState<AccentType>('vec')

  const insertAccent = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      let mathContent = ''
      switch (accentType) {
        case 'vec':
          mathContent =
            `<m:acc><m:accPr><m:chr m:val="\u20d7"/></m:accPr>` +
            `<m:e><m:r><m:t></m:t></m:r></m:e></m:acc>`
          break
        case 'overlineABC':
          mathContent =
            `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>` +
            `<m:e><m:r><m:t>ABC</m:t></m:r></m:e></m:bar>`
          break
        case 'overlineXOR':
          mathContent =
            `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>` +
            `<m:e><m:r><m:t>x</m:t></m:r><m:r><m:t>\u2295</m:t></m:r><m:r><m:t>y</m:t></m:r></m:e></m:bar>`
          break
      }
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={accentType} onChange={(_, d) => setAccentType(d.value as AccentType)}>
          {ACCENT_TYPES.map((a) => <option key={a.value} value={a.value}>{a.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertAccent}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
