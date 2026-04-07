// src/components/features/template/TemplateTextFeature.tsx
import { useState } from 'react'
import { Button, Input, Radio, RadioGroup, Text, makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { useWordRun } from '../../../hooks/useWordRun'

const TEMPLATE_COUNT = 5
const SETTINGS_KEY_TEMPLATE = 'templateTexts'

const getSettings = () => Office.context.document.settings
const loadSetting = <T,>(key: string, fallback: T): T => {
  const val = getSettings().get(key)
  return val !== null && val !== undefined ? (val as T) : fallback
}
const saveSetting = (key: string, value: unknown) => {
  getSettings().set(key, value)
  getSettings().saveAsync()
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  templateRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'center',
    minWidth: 0,
    width: '100%',
    marginBottom: tokens.spacingVerticalS,
  },
  radioLabel: { flexShrink: 0, whiteSpace: 'nowrap', minWidth: '64px' },
  templateInput: {
    flex: 1,
    minWidth: 0,
    width: '100%',
    '& input': { minWidth: 0, width: '100%', boxSizing: 'border-box' },
  },
  row: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: tokens.spacingHorizontalS, width: '100%' },
  hint: { color: tokens.colorNeutralForeground3, fontSize: '10px' },
  btnFull: { width: '100%', fontSize: '11px', whiteSpace: 'nowrap' },
})

export function TemplateTextFeature() {
  const styles = useStyles()
  const { runWord, status, setStatus } = useWordRun()
  const [templates, setTemplates] = useState<string[]>(() =>
    loadSetting<string[]>(SETTINGS_KEY_TEMPLATE, Array(TEMPLATE_COUNT).fill(''))
  )
  const [selectedTemplate, setSelectedTemplate] = useState('0')

  const updateTemplate = (idx: number, value: string) => {
    const next = [...templates]
    next[idx] = value
    setTemplates(next)
    saveSetting(SETTINGS_KEY_TEMPLATE, next)
  }

  const copyFromSelection = () =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      range.load('text')
      await context.sync()
      if (!range.text) {
        setStatus({ type: 'warning', message: 'テキストを選択してから実行してください' })
        return
      }
      const idx = parseInt(selectedTemplate, 10)
      updateTemplate(idx, range.text)
    })

  const insertTemplate = () =>
    runWord(async (context) => {
      const idx = parseInt(selectedTemplate, 10)
      const text = templates[idx]
      if (!text) {
        setStatus({ type: 'warning', message: `定型文 ${idx + 1} が登録されていません` })
        return
      }
      const range = context.document.getSelection()
      range.insertText(text, Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <RadioGroup value={selectedTemplate} onChange={(_, d) => setSelectedTemplate(d.value)}>
        {templates.map((t, i) => (
          <div key={i} className={styles.templateRow}>
            <div className={styles.radioLabel}>
              <Radio value={String(i)} label={`定型文 ${i + 1}`} />
            </div>
            <Input
              className={styles.templateInput}
              value={t}
              placeholder={`（定型文 ${i + 1}）`}
              onChange={(_, d) => updateTemplate(i, d.value)}
            />
          </div>
        ))}
      </RadioGroup>
      <div className={styles.row}>
        <Button appearance="secondary" className={styles.btnFull} onClick={copyFromSelection}>
          文章よりコピー登録
        </Button>
        <Button appearance="primary" className={styles.btnFull} onClick={insertTemplate}>
          実行
        </Button>
      </div>
      <Text size={100} className={styles.hint}>
        ラジオボタンで定型文を選択してから「実行」を押すと挿入されます。
      </Text>
      <StatusBar status={status} />
    </div>
  )
}
