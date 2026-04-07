// src/components/features/frame/ImageInsertFeature.tsx
import { useRef } from 'react'
import { Button, Text, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  hiddenInput: { display: 'none' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function ImageInsertFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()
  const fileInputRef = useRef<HTMLInputElement>(null)

  const insertImage = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const base64 = (reader.result as string).split(',')[1]
      runWord(async (context) => {
        const range = context.document.getSelection()
        range.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace)
        await context.sync()
      })
    }
    reader.readAsDataURL(file)
    e.target.value = ''
  }

  return (
    <div className={styles.root}>
      <Text size={200}>画像・図形を文章内に挿入します。カーソル位置に挿入されます。</Text>
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        className={styles.hiddenInput}
        onChange={insertImage}
      />
      <Button appearance="primary" className={styles.btnFull} onClick={() => fileInputRef.current?.click()}>
        画像・オブジェクトの挿入
      </Button>
      <StatusBar status={status} />
    </div>
  )
}
