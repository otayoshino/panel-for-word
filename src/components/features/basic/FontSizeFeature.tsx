// src/components/features/basic/FontSizeFeature.tsx
// 基本文字サイズの設定 — 選択範囲のフォントサイズを変更する

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

export function FontSizeFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [fontSize, setFontSize] = useState(10.5)

  const applyFontSize = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.font.size = fontSize
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <Field label="文字サイズ (pt)">
        <SpinButton
          value={fontSize}
          min={6}
          max={72}
          step={0.5}
          onChange={(_, d) => setFontSize(d.value ?? 10.5)}
        />
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={applyFontSize}>
        選択範囲に適用
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
