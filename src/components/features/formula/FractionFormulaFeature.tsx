// src/components/features/formula/FractionFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type FracType = 'bar' | 'skw' | 'lin' | 'noBar'
const FRAC_TYPES: { value: FracType; label: string }[] = [
  { value: 'bar',   label: '縦積み（横線あり）' },
  { value: 'skw',   label: '斜め分数' },
  { value: 'lin',   label: '線形（a/b）' },
  { value: 'noBar', label: '分数（小）' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function FractionFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [fracType, setFracType] = useState<FracType>('bar')

  const insertFraction = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const fPr = fracType !== 'bar' ? `<m:fPr><m:type m:val="${fracType}"/></m:fPr>` : ''
      const mathContent = `<m:f>${fPr}<m:num><m:r><m:t></m:t></m:r></m:num><m:den><m:r><m:t></m:t></m:r></m:den></m:f>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="分数タイプ">
        <Select value={fracType} onChange={(_, d) => setFracType(d.value as FracType)}>
          {FRAC_TYPES.map((f) => <option key={f.value} value={f.value}>{f.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertFraction}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
