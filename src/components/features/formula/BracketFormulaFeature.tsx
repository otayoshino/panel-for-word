// src/components/features/formula/BracketFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type BracketType = 'cases' | 'binom' | 'binomAngle'
const BRACKET_TYPES: { value: BracketType; label: string }[] = [
  { value: 'cases',      label: '場合分けを使う数式の例' },
  { value: 'binom',      label: '２項係数' },
  { value: 'binomAngle', label: '二項係数（山かっこ付き）' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function BracketFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [bracketType, setBracketType] = useState<BracketType>('cases')

  const insertBracket = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      const frac = `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${empty}</m:num><m:den>${empty}</m:den></m:f>`

      let mathContent = ''
      switch (bracketType) {
        case 'cases': {
          const row = `<m:e>${empty}</m:e>`
          mathContent = `<m:d><m:dPr><m:begChr m:val="{"/><m:sepChr m:val=""/><m:endChr m:val=""/></m:dPr><m:e><m:eqArr>${row}${row}</m:eqArr></m:e></m:d>`
          break
        }
        case 'binom':
          mathContent = `<m:d><m:e>${frac}</m:e></m:d>`
          break
        case 'binomAngle':
          mathContent = `<m:d><m:dPr><m:begChr m:val="\u27e8"/><m:endChr m:val="\u27e9"/></m:dPr><m:e>${frac}</m:e></m:d>`
          break
      }
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={bracketType} onChange={(_, d) => setBracketType(d.value as BracketType)}>
          {BRACKET_TYPES.map((b) => <option key={b.value} value={b.value}>{b.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertBracket}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
