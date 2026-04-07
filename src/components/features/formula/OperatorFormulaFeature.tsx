// src/components/features/formula/OperatorFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type OperatorType = 'colonEq' | 'doubleEq' | 'plusEq' | 'minusEq' | 'defEq' | 'measure' | 'deltaEq'
const OPERATOR_TYPES: { value: OperatorType; label: string; sym: string }[] = [
  { value: 'colonEq',  label: 'コロン付き等号',       sym: '\u2254' },
  { value: 'doubleEq', label: '二重等号',              sym: '==' },
  { value: 'plusEq',   label: 'プラス付き等号',        sym: '+=' },
  { value: 'minusEq',  label: 'マイナス付き等号',      sym: '\u2212=' },
  { value: 'defEq',    label: '定義により等しい',      sym: '\u225d' },
  { value: 'measure',  label: '測度',                  sym: '\u2250' },
  { value: 'deltaEq',  label: 'デルタ付き等号',        sym: '\u225c' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function OperatorFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [operatorType, setOperatorType] = useState<OperatorType>('colonEq')

  const insertOperator = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const entry = OPERATOR_TYPES.find((o) => o.value === operatorType)
      if (!entry) return
      const mathContent = `<m:r><m:t>${entry.sym}</m:t></m:r>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={operatorType} onChange={(_, d) => setOperatorType(d.value as OperatorType)}>
          {OPERATOR_TYPES.map((o) => (
            <option key={o.value} value={o.value}>{o.label}　{o.sym}</option>
          ))}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertOperator}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
