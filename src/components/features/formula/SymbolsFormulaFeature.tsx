// src/components/features/formula/SymbolsFormulaFeature.tsx
import { Button, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const SYMBOLS = ['#', '$', '%', '&', '@'] as const

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  buttonRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(5, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
  },
  symbolButton: { minWidth: 'unset', width: '100%', fontFamily: 'monospace', fontSize: '16px' },
})

export function SymbolsFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()

  const insertSymbol = (sym: string) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.insertText(sym, Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <div className={styles.buttonRow}>
        {SYMBOLS.map((sym) => (
          <Button key={sym} appearance="secondary" className={styles.symbolButton} onClick={() => insertSymbol(sym)}>
            {sym}
          </Button>
        ))}
      </div>
      <StatusBar status={status} />
    </div>
  )
}
