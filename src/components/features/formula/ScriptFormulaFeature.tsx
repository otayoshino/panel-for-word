// src/components/features/formula/ScriptFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type ScriptType = 'sup' | 'sub' | 'subSup' | 'leftSubSup'
const SCRIPT_TYPES: { value: ScriptType; label: string }[] = [
  { value: 'sup',        label: '上付き文字' },
  { value: 'sub',        label: '下付き文字' },
  { value: 'subSup',     label: '下付き文字-上付き文字' },
  { value: 'leftSubSup', label: '左下付き文字-上付き文字' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function ScriptFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [scriptType, setScriptType] = useState<ScriptType>('sup')

  const insertScript = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const e   = '<m:e><m:r><m:t></m:t></m:r></m:e>'
      const sub = '<m:sub><m:r><m:t></m:t></m:r></m:sub>'
      const sup = '<m:sup><m:r><m:t></m:t></m:r></m:sup>'

      let mathContent = ''
      switch (scriptType) {
        case 'sup':        mathContent = `<m:sSup>${e}${sup}</m:sSup>`; break
        case 'sub':        mathContent = `<m:sSub>${e}${sub}</m:sSub>`; break
        case 'subSup':     mathContent = `<m:sSubSup>${e}${sub}${sup}</m:sSubSup>`; break
        case 'leftSubSup': mathContent = `<m:sSup><m:e><m:sPre>${sub}${sup}${e}</m:sPre></m:e>${sup}</m:sSup>`; break
      }
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="種類">
        <Select value={scriptType} onChange={(_, d) => setScriptType(d.value as ScriptType)}>
          {SCRIPT_TYPES.map((s) => <option key={s.value} value={s.value}>{s.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertScript}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
