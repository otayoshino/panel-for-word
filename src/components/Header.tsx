// src/components/Header.tsx
// 共通ヘッダー — 画面状態に応じてアドイン名またはナビゲーション戻るボタンを表示

import { Button, Text, makeStyles } from '@fluentui/react-components'
import type { FeatureItem } from '../types/feature'

interface HeaderProps {
  currentFeature: FeatureItem | null
  onBack: () => void
}

const useStyles = makeStyles({
  header: {
    backgroundColor: '#1e4d8c',
    padding: '8px 12px',
    width: '100%',
    boxSizing: 'border-box',
    display: 'flex',
    alignItems: 'center',
    minHeight: '40px',
    flexShrink: 0,
  },
  title: {
    color: '#ffffff',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '14px',
    fontWeight: '600',
  },
  backButton: {
    color: '#ffffff',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    fontSize: '13px',
    minWidth: 0,
    paddingLeft: '0',
    paddingRight: '4px',
    // ホバー時も白文字を維持
    ':hover': {
      color: '#ffffff',
      backgroundColor: 'rgba(255,255,255,0.15)',
    },
    ':active': {
      color: '#ffffff',
    },
  },
})

export function Header({ currentFeature, onBack }: HeaderProps) {
  const styles = useStyles()

  return (
    <div className={styles.header}>
      {currentFeature === null ? (
        // メイン画面：アドイン名を表示
        <Text className={styles.title}>かんたんツールボックス</Text>
      ) : (
        // 設定画面：「← 機能名」の戻るボタンを表示
        <Button
          appearance="subtle"
          className={styles.backButton}
          onClick={onBack}
          aria-label={`${currentFeature.label}から戻る`}
        >
          ← {currentFeature.label}
        </Button>
      )}
    </div>
  )
}
