// src/components/features/formula/IntegralFormulaFeature.tsx
import { useState } from 'react'
import { Button, Field, Select, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type IntegralType =
  | 'int' | 'intLim' | 'intStack'
  | 'iint' | 'iintLim' | 'iintStack'
  | 'iiint' | 'iiintLim' | 'iiintStack'
const INTEGRAL_TYPES: { value: IntegralType; label: string }[] = [
  { value: 'int',        label: '積分' },
  { value: 'intLim',     label: '積分（上下端値あり）' },
  { value: 'intStack',   label: '積分（上下端値を上下に配置）' },
  { value: 'iint',       label: '２重積分' },
  { value: 'iintLim',    label: '二重積分（上下端値あり）' },
  { value: 'iintStack',  label: '二重積分（上下端値を上下に配置）' },
  { value: 'iiint',      label: '３重積分' },
  { value: 'iiintLim',   label: '三重積分（上下端値あり）' },
  { value: 'iiintStack', label: '三重積分（上下端値を上下に配置）' },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function IntegralFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [integralType, setIntegralType] = useState<IntegralType>('int')

  const insertIntegral = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const empty = '<m:r><m:t></m:t></m:r>'
      const CHR_MAP: Record<IntegralType, string> = {
        int: '\u222B', intLim: '\u222B', intStack: '\u222B',
        iint: '\u222C', iintLim: '\u222C', iintStack: '\u222C',
        iiint: '\u222D', iiintLim: '\u222D', iiintStack: '\u222D',
      }
      const showLimits = integralType.endsWith('Lim') || integralType.endsWith('Stack')
      const limLoc = integralType.endsWith('Stack') ? 'undOvr' : 'subSup'
      const chr = CHR_MAP[integralType]
      const hidePr = showLimits ? '' : '<m:subHide m:val="1"/><m:supHide m:val="1"/>'
      const naryPr = `<m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="${limLoc}"/>${hidePr}</m:naryPr>`
      const sub = showLimits ? `<m:sub>${empty}</m:sub>` : '<m:sub/>'
      const sup = showLimits ? `<m:sup>${empty}</m:sup>` : '<m:sup/>'
      const mathContent = `<m:nary>${naryPr}${sub}${sup}<m:e>${empty}</m:e></m:nary>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="積分タイプ">
        <Select value={integralType} onChange={(_, d) => setIntegralType(d.value as IntegralType)}>
          {INTEGRAL_TYPES.map((it) => <option key={it.value} value={it.value}>{it.label}</option>)}
        </Select>
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertIntegral}>挿入</Button>
      <StatusBar status={status} />
    </div>
  )
}
