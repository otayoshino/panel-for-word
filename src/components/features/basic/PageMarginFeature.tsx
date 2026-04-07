// src/components/features/basic/PageMarginFeature.tsx
// ページ余白（上下左右、mm単位）の設定

import { useState } from 'react'
import { Button, Field, Input, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const mm2pt = (mm: number) => mm * 2.8346

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  marginGrid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    boxSizing: 'border-box',
  },
  marginField: {
    minWidth: 0,
    width: '100%',
    '& input': {
      minWidth: 0,
      width: '100%',
      boxSizing: 'border-box',
    },
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function PageMarginFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [marginTop, setMarginTop] = useState('')
  const [marginBottom, setMarginBottom] = useState('')
  const [marginLeft, setMarginLeft] = useState('')
  const [marginRight, setMarginRight] = useState('')

  const applyMargins = () =>
    runWord(async (context) => {
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      if (marginTop !== '')    ps.topMargin    = mm2pt(parseFloat(marginTop))
      if (marginBottom !== '') ps.bottomMargin = mm2pt(parseFloat(marginBottom))
      if (marginLeft !== '')   ps.leftMargin   = mm2pt(parseFloat(marginLeft))
      if (marginRight !== '')  ps.rightMargin  = mm2pt(parseFloat(marginRight))
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <div className={styles.marginGrid}>
        <Field label="①上（天）" className={styles.marginField}>
          <Input type="number" value={marginTop} onChange={(_, d) => setMarginTop(d.value)} placeholder="mm" />
        </Field>
        <Field label="②下（地）" className={styles.marginField}>
          <Input type="number" value={marginBottom} onChange={(_, d) => setMarginBottom(d.value)} placeholder="mm" />
        </Field>
        <Field label="③左" className={styles.marginField}>
          <Input type="number" value={marginLeft} onChange={(_, d) => setMarginLeft(d.value)} placeholder="mm" />
        </Field>
        <Field label="④右" className={styles.marginField}>
          <Input type="number" value={marginRight} onChange={(_, d) => setMarginRight(d.value)} placeholder="mm" />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={applyMargins}>
        実行
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
