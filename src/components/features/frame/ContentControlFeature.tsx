// src/components/features/frame/ContentControlFeature.tsx
import { useState } from 'react'
import { Button, Field, Input, MessageBar, MessageBarBody, Text, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  noticeBox: {
    backgroundColor: '#dce8f7',
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalS,
  },
  preWrap: { whiteSpace: 'pre-line' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function ContentControlFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const [ccTitle, setCcTitle] = useState('')

  const insertContentControl = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const cc = range.insertContentControl()
      cc.tag = 'textframe'
      cc.title = ccTitle || 'テキスト枠'
      cc.appearance = Word.ContentControlAppearance.boundingBox
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <MessageBar intent="warning">
        <MessageBarBody>
          Office.jsではテキストボックスの直接作成はできません。ContentControl（コンテンツコントロール）で代替します。
        </MessageBarBody>
      </MessageBar>
      <Field label="枠タイトル（任意）">
        <Input value={ccTitle} onChange={(_, d) => setCcTitle(d.value)} placeholder="例: 図キャプション" />
      </Field>
      <Button appearance="primary" className={styles.btnFull} onClick={insertContentControl}>
        テキスト枠（ContentControl）を挿入
      </Button>
      <div className={styles.noticeBox}>
        <Text size={200} block className={styles.preWrap}>
          {'非対応機能（廃止）:\n• 画像枠（プレースホルダー）作成\n• 図形挿入\n• 重ね順・枠揃え\n• 文字列の折り返し設定\n• サイズパレット表示'}
        </Text>
      </div>
      <StatusBar status={status} />
    </div>
  )
}
