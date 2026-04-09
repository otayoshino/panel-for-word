// src/components/features/basic/ColumnLayoutFeature.tsx
// 段組み（段数）の設定 — pageSetup.textColumns API（WordApiDesktop 1.3）を使用

import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function ColumnLayoutFeature() {
  const styles = useStyles()
  const { runRaw, status } = useWordRun()
  const [colCount, setColCount] = useState(1)

  const applyColumns = () =>
    runRaw(async () => {
      // Run 1: 段数設定のみ（setCount後プロキシが無効になるため単独）
      await Word.run(async (context) => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()
        sections.items[0].pageSetup.textColumns.setCount(colCount)
        await context.sync()
      })
    })

  return (
    <div className={styles.root}>
      <Field label="段数">
        <SpinButton
          value={colCount}
          min={1}
          max={10}
          step={1}
          onChange={(_, d) => setColCount(d.value ?? 1)}
        />
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={applyColumns}>
        実行
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
