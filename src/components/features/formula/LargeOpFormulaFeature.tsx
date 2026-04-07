// src/components/features/formula/LargeOpFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type LargeOpType = 'sumCondition' | 'sumFromTo' | 'sumTwoSub' | 'prod' | 'union'
const LARGE_OP_TYPES: { value: LargeOpType; label: string }[] = [
  { value: 'sumCondition', label: 'nからkを選ぶ場合のkの総和' },
  { value: 'sumFromTo',    label: '総和（i=0からnまで）' },
  { value: 'sumTwoSub',    label: '添え字２個を使う総和の例' },
  { value: 'prod',         label: '積の例' },
  { value: 'union',        label: '和集合の例' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function LargeOpFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [largeOpType, setLargeOpType] = useState<LargeOpType>('sumFromTo')

  const insertLargeOp = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      const sub  = `<m:sub>${empty}</m:sub>`
      const sup  = `<m:sup>${empty}</m:sup>`
      const e    = `<m:e>${empty}</m:e>`
      const twoSub = `<m:sub>${empty}<m:r><m:rPr><m:nor/></m:rPr><m:t>,\u00a0</m:t></m:r>${empty}</m:sub>`

      type NarySpec = { chr: string; naryPr: string; sub: string; sup: string }
      const specs: Record<LargeOpType, NarySpec> = {
        sumCondition: { chr: '\u2211', naryPr: '<m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>', sub, sup: '<m:sup/>' },
        sumFromTo:    { chr: '\u2211', naryPr: '<m:limLoc m:val="undOvr"/>', sub, sup },
        sumTwoSub:    { chr: '\u2211', naryPr: '<m:limLoc m:val="undOvr"/><m:supHide m:val="1"/>', sub: twoSub, sup: '<m:sup/>' },
        prod:         { chr: '\u220F', naryPr: '<m:limLoc m:val="undOvr"/>', sub, sup },
        union:        { chr: '\u22C3', naryPr: '<m:limLoc m:val="undOvr"/>', sub, sup },
      }
      const s = specs[largeOpType]
      const mathContent = `<m:nary><m:naryPr><m:chr m:val="${s.chr}"/>${s.naryPr}</m:naryPr>${s.sub}${s.sup}${e}</m:nary>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={largeOpType} onChange={(_, d) => setLargeOpType(d.value as LargeOpType)}>
          {LARGE_OP_TYPES.map((op) => <option key={op.value} value={op.value}>{op.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertLargeOp}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
