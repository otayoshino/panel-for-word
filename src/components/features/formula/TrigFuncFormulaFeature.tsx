// src/components/features/formula/TrigFuncFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type TrigFuncType = 'sin' | 'cos' | 'tan'
const TRIG_FUNC_TYPES: { value: TrigFuncType; label: string }[] = [
  { value: 'sin', label: 'sin\u03b8' },
  { value: 'cos', label: 'Cos\u00a02x' },
  { value: 'tan', label: '正接式' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function TrigFuncFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [trigFuncType, setTrigFuncType] = useState<TrigFuncType>('sin')

  const insertTrigFunc = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      type FuncSpec = { name: string; arg: string }
      const specs: Record<TrigFuncType, FuncSpec> = {
        sin: { name: 'sin', arg: '\u03b8' },
        cos: { name: 'Cos', arg: '2x' },
        tan: { name: 'tan', arg: '' },
      }
      const s = specs[trigFuncType]
      const argContent = s.arg ? `<m:r><m:t>${s.arg}</m:t></m:r>` : empty
      const mathContent =
        `<m:func>` +
        `<m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${s.name}</m:t></m:r></m:fName>` +
        `<m:e>${argContent}</m:e>` +
        `</m:func>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={trigFuncType} onChange={(_, d) => setTrigFuncType(d.value as TrigFuncType)}>
          {TRIG_FUNC_TYPES.map((f) => <option key={f.value} value={f.value}>{f.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertTrigFunc}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
