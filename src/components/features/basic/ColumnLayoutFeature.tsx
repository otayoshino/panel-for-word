// src/components/features/basic/ColumnLayoutFeature.tsx
// 段組み（段数・列間隔）の設定 — OOXML の w:cols 要素を直接編集して適用する

import { useState } from 'react'
import { Button, Field, SpinButton, makeStyles, tokens } from '@fluentui/react-components'
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
  grid: {
    display: 'grid',
    gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)',
    gap: tokens.spacingHorizontalS,
    width: '100%',
    boxSizing: 'border-box',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
})

export function ColumnLayoutFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [colCount, setColCount] = useState(1)
  const [colSpacing, setColSpacing] = useState(10)

  const applyColumns = () =>
    runWord(async (context) => {
      const body = context.document.body
      const ooxmlResult = body.getOoxml()
      await context.sync()

      // w:space は twip（1/20pt）単位
      const spaceTwips = Math.round(mm2pt(colSpacing) * 20)
      const colsTag = `<w:cols w:equalWidth="1" w:num="${colCount}" w:space="${spaceTwips}"/>`
      let xml: string = ooxmlResult.value

      const SECT_CLOSE = '</w:sectPr>'
      const lastClose = xml.lastIndexOf(SECT_CLOSE)
      if (lastClose === -1) return

      const lastOpen = xml.lastIndexOf('<w:sectPr', lastClose)
      if (lastOpen === -1) return

      const sectPrXml = xml.slice(lastOpen, lastClose + SECT_CLOSE.length)
      let newSectPrXml: string

      if (/\bw:cols\b/.test(sectPrXml)) {
        // 既存の w:cols を置換（自己終了タグ → 子要素ありの順で試行）
        newSectPrXml = sectPrXml.replace(/<w:cols[^>]*\/>/, colsTag)
        if (newSectPrXml === sectPrXml) {
          newSectPrXml = sectPrXml.replace(/<w:cols[\s\S]*?<\/w:cols>/, colsTag)
        }
      } else {
        // w:cols なし → </w:sectPr> 直前に挿入
        newSectPrXml = sectPrXml.replace(SECT_CLOSE, colsTag + SECT_CLOSE)
      }

      xml = xml.slice(0, lastOpen) + newSectPrXml + xml.slice(lastClose + SECT_CLOSE.length)
      body.insertOoxml(xml, 'Replace')
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <div className={styles.grid}>
        <Field label="段数">
          <SpinButton
            value={colCount}
            min={1}
            max={10}
            step={1}
            onChange={(_, d) => setColCount(d.value ?? 1)}
          />
        </Field>
        <Field label="列間隔 (mm)">
          <SpinButton
            value={colSpacing}
            min={0}
            max={100}
            step={1}
            onChange={(_, d) => setColSpacing(d.value ?? 10)}
          />
        </Field>
      </div>
      <Button appearance="primary" className={styles.btnFull} onClick={applyColumns}>
        実行
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
