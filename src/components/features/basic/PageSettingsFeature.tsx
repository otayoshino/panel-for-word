// src/components/features/basic/PageSettingsFeature.tsx
// ページ設定情報の確認 — 現在のドキュメントの用紙・余白・文字サイズ情報を表示する

import { useState } from 'react'
import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    gap: tokens.spacingVerticalS,
  },
  infoBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: '8px',
    wordBreak: 'break-all',
    minHeight: '84px',
    lineHeight: '1.8',
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    boxSizing: 'border-box',
  },
  btnFull: {
    width: '100%',
    fontSize: '11px',
  },
  preWrap: {
    whiteSpace: 'pre-line',
  },
})

export function PageSettingsFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [docInfo, setDocInfo] = useState<string | null>(null)

  const getDocSettings = () =>
    runWord(async (context) => {
      const body = context.document.body
      body.load('style')
      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      const sec = sections.items[0]
      sec.load('body/style')
      const ps = sec.body.getRange('Whole')
      ps.load('paragraphs')
      await context.sync()

      const firstPara = ps.paragraphs.getFirst()
      firstPara.load('font/size,font/name,lineSpacing')
      await context.sync()

      const pageSetup = sec.pageSetup
      pageSetup.load('pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin')
      await context.sync()

      const toMm = (pt: number) => (pt / 2.8346).toFixed(1)
      setDocInfo(
        `用紙: ${toMm(pageSetup.pageWidth)}×${toMm(pageSetup.pageHeight)}mm\n` +
        `余白 上:${toMm(pageSetup.topMargin)} 下:${toMm(pageSetup.bottomMargin)}\n` +
        `      左:${toMm(pageSetup.leftMargin)} 右:${toMm(pageSetup.rightMargin)}mm\n` +
        `文字: ${firstPara.font.size}pt / ${firstPara.font.name}`
      )
      setStatus(null)
    })

  return (
    <div className={styles.root}>
      <Button appearance="secondary" className={styles.btnFull} onClick={getDocSettings}>
        現在のドキュメントの設定値を取得
      </Button>
      <div className={styles.infoBox}>
        <Text size={200} className={styles.preWrap}>
          {docInfo ?? 'ボタンを押すとここに表示されます'}
        </Text>
      </div>
      <StatusBar status={status} />
    </div>
  )
}
