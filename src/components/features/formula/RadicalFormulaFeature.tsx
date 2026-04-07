// src/components/features/formula/RadicalFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type RadicalType = 'sqrt' | 'nthRoot' | 'sqrtWithDeg' | 'cbrt'
const RADICAL_TYPES: { value: RadicalType; label: string }[] = [
  { value: 'sqrt',        label: '平方根' },
  { value: 'nthRoot',     label: '次数付きべき乗根' },
  { value: 'sqrtWithDeg', label: '次数付き平方根' },
  { value: 'cbrt',        label: '立方根' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function RadicalFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [radicalType, setRadicalType] = useState<RadicalType>('sqrt')

  const insertRadical = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      let radPr = ''
      let deg = ''
      switch (radicalType) {
        case 'sqrt':        radPr = '<m:radPr><m:degHide m:val="1"/></m:radPr>'; deg = `<m:deg></m:deg>`; break
        case 'nthRoot':     deg = `<m:deg>${empty}</m:deg>`; break
        case 'sqrtWithDeg': deg = `<m:deg><m:r><m:t>2</m:t></m:r></m:deg>`; break
        case 'cbrt':        deg = `<m:deg><m:r><m:t>3</m:t></m:r></m:deg>`; break
      }
      const mathContent = `<m:rad>${radPr}${deg}<m:e>${empty}</m:e></m:rad>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={radicalType} onChange={(_, d) => setRadicalType(d.value as RadicalType)}>
          {RADICAL_TYPES.map((r) => <option key={r.value} value={r.value}>{r.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertRadical}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
