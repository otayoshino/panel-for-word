// src/components/features/formula/MatrixFormulaFeature.tsx
import { Button, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function MatrixFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()

  const insertMatrix = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const e = '<m:e><m:r><m:t></m:t></m:r></m:e>'
      const mathContent = `<m:m>
  <m:mPr>
    <m:mcs>
      <m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc>
    </m:mcs>
  </m:mPr>
  <m:mr>${e}${e}</m:mr>
  <m:mr>${e}${e}</m:mr>
</m:m>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Button appearance="primary" className={styles.btnFull} onClick={insertMatrix}>
        行列（2×2）を挿入
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
