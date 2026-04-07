// src/components/features/basic/PaperSizeFeature.tsx
// 用紙サイズと組方向（横組み/縦組み）の設定

import { useState } from 'react'
import {
  Button,
  Field,
  Select,
  Radio,
  RadioGroup,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

// 用紙サイズ定義（単位: pt, 1mm = 2.8346pt / JIS B系採用）
const PAPER_SIZES: Record<string, { width: number; height: number }> = {
  'A3縦': { width: 841.89, height: 1190.55 },
  'A3横': { width: 1190.55, height: 841.89 },
  'A4縦': { width: 595.28, height: 841.89 },
  'A4横': { width: 841.89, height: 595.28 },
  'A5縦': { width: 419.53, height: 595.28 },
  'A5横': { width: 595.28, height: 419.53 },
  'A6縦': { width: 297.64, height: 419.53 },
  'A6横': { width: 419.53, height: 297.64 },
  'B4縦': { width: 728.50, height: 1031.81 },
  'B4横': { width: 1031.81, height: 728.50 },
  'B5縦': { width: 515.91, height: 728.50 },
  'B5横': { width: 728.50, height: 515.91 },
  'B6縦': { width: 362.83, height: 515.91 },
  'B6横': { width: 515.91, height: 362.83 },
  'レター縦': { width: 612, height: 792 },
  'レター横': { width: 792, height: 612 },
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  row: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'flex-end',
    flexWrap: 'wrap',
    width: '100%',
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: '10px',
  },
})

export function PaperSizeFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [paperSize, setPaperSize] = useState('A4縦')
  const [textDir, setTextDir] = useState<'horizontal' | 'vertical'>('horizontal')

  const applyPaperSize = () =>
    runWord(async (context) => {
      const size = PAPER_SIZES[paperSize]
      if (!size) return
      const sections = context.document.sections
      sections.load('items')
      await context.sync()
      const ps = sections.items[0].pageSetup
      ps.pageWidth = size.width
      ps.pageHeight = size.height
      await context.sync()
    })

  const applyTextDirection = (dir: 'horizontal' | 'vertical') => {
    setTextDir(dir)
    runWord(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs
      paragraphs.load('items')
      await context.sync()
      paragraphs.items.forEach((p) => {
        // textDirection は @types/office-js 未定義だが Word API では有効
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        ;(p as unknown as Record<string, unknown>)['textDirection'] =
          dir === 'horizontal' ? 'LeftToRight' : 'TopToBottom'
      })
      await context.sync()
    })
  }

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Field label="用紙サイズ">
          <Select value={paperSize} onChange={(_, d) => setPaperSize(d.value)}>
            {Object.keys(PAPER_SIZES).map((k) => (
              <option key={k} value={k}>{k}</option>
            ))}
          </Select>
        </Field>
        <Button appearance="primary" onClick={applyPaperSize}>設定</Button>
      </div>
      <RadioGroup
        layout="horizontal"
        value={textDir}
        onChange={(_, d) => applyTextDirection(d.value as 'horizontal' | 'vertical')}
      >
        <Radio value="horizontal" label="横組み" />
        <Radio value="vertical" label="縦組み" />
      </RadioGroup>
      <Text size={100} className={styles.hint}>
        組方向は選択中の段落に適用されます。セクション全体への縦組みはAPI制限のため完全には適用できません。
      </Text>
      <StatusBar status={status} />
    </div>
  )
}
