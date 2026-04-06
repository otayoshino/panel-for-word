import { useState } from 'react'
import {
  FluentProvider,
  webLightTheme,
  Text,
  makeStyles,
  type Theme,
} from '@fluentui/react-components'
import { BasicSettingsTab } from './components/tabs/BasicSettingsTab'
import { CharCompositionTab } from './components/tabs/CharCompositionTab'
import { FrameTab } from './components/tabs/FrameTab'
import { FormulaTab } from './components/tabs/FormulaTab'
import { TemplateTextTab } from './components/tabs/TemplateTextTab'

type TabId = 'basic' | 'char' | 'frame' | 'formula' | 'template'

const TABS: { id: TabId; label: string }[] = [
  { id: 'basic', label: '基本設定' },
  { id: 'char', label: '文字組' },
  { id: 'frame', label: '枠' },
  { id: 'formula', label: '数式' },
  { id: 'template', label: '定型文' },
]

// B案カラーパレットに合わせたカスタムテーマ
const meiyushaTheme: Theme = {
  ...webLightTheme,
  colorBrandBackground: '#0c51a0',
  colorBrandBackgroundHover: '#185fa5',
  colorBrandBackgroundPressed: '#0c3370',
  colorBrandForeground1: '#0c51a0',
  colorBrandForeground2: '#185fa5',
  colorNeutralBackground1: '#f5f9ff',
  colorNeutralBackground2: '#dce8f7',
  colorNeutralBackground3: '#dce8f7',
  colorNeutralStroke1: '#c5dcf5',
  colorNeutralStroke2: '#c5dcf5',
  colorNeutralForeground1: '#0c3370',
  colorNeutralForeground2: '#4a7cb5',
  colorNeutralForeground3: '#7fb5e8',
  fontFamilyBase: "'Noto Sans JP', 'Segoe UI', sans-serif",
  fontFamilyNumeric: "'Sora', 'Segoe UI', sans-serif",
  fontSizeBase300: '11px',
}

const useStyles = makeStyles({
  provider: {
    width: '100%',
    maxWidth: '100%',
    boxSizing: 'border-box',
  },
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    maxWidth: '100%',
    height: '100vh',
    overflow: 'hidden',
    boxSizing: 'border-box',
  },
  header: {
    backgroundColor: '#0c51a0',
    padding: '12px 12px 0',
    flexShrink: 0,
    boxSizing: 'border-box',
    width: '100%',
  },
  brandName: {
    display: 'block',
    color: '#7fb5e8',
    fontSize: '8px',
    letterSpacing: '0.25em',
    marginBottom: '3px',
    fontFamily: "'Sora', 'Noto Sans JP', sans-serif",
  },
  titleLight: {
    color: '#ffffff',
    fontSize: '17px',
    fontWeight: '300',
    lineHeight: '1.2',
    display: 'block',
    fontFamily: "'Noto Sans JP', sans-serif",
  },
  titleBold: {
    color: '#ffffff',
    fontSize: '17px',
    fontWeight: '500',
    lineHeight: '1.2',
    display: 'block',
    marginBottom: '12px',
    fontFamily: "'Noto Sans JP', sans-serif",
  },
  tabBar: {
    display: 'flex',
    alignItems: 'flex-end',
    width: '100%',
  },
  tabItem: {
    flex: 1,
    padding: '6px 4px',
    fontSize: '9px',
    color: '#7fb5e8',
    backgroundColor: 'transparent',
    border: 'none',
    borderRadius: '6px 6px 0 0',
    cursor: 'pointer',
    fontFamily: "'Noto Sans JP', sans-serif",
    fontWeight: '400',
    appearance: 'none',
    textAlign: 'center',
    whiteSpace: 'nowrap',
    ':hover': {
      backgroundColor: 'rgba(255,255,255,0.15)',
      color: '#b8d4f0',
    },
  },
  tabSelected: {
    flex: 1,
    padding: '6px 4px',
    fontSize: '9px',
    color: '#0c51a0',
    fontWeight: '500',
    backgroundColor: '#f5f9ff',
    border: 'none',
    borderRadius: '6px 6px 0 0',
    cursor: 'pointer',
    fontFamily: "'Noto Sans JP', sans-serif",
    appearance: 'none',
    textAlign: 'center',
    whiteSpace: 'nowrap',
  },
  body: {
    flex: 1,
    backgroundColor: '#f5f9ff',
    padding: '12px',
    overflowY: 'scroll',
    overflowX: 'hidden',
    boxSizing: 'border-box',
    width: '100%',
  },
})

export default function App() {
  const styles = useStyles()
  const [activeTab, setActiveTab] = useState<TabId>('basic')

  return (
    <FluentProvider theme={meiyushaTheme} className={styles.provider}>
      <div className={styles.root}>

        {/* ── ヘッダー ── */}
        <div className={styles.header}>
          <Text block className={styles.brandName}>MEIYUSHA</Text>
          <Text block className={styles.titleLight}>かんたん</Text>
          <Text block className={styles.titleBold}>ツールボックス</Text>

          {/* タブバー */}
          <div className={styles.tabBar} role="tablist">
            {TABS.map((tab) => (
              <button
                key={tab.id}
                role="tab"
                aria-selected={activeTab === tab.id}
                className={activeTab === tab.id ? styles.tabSelected : styles.tabItem}
                onClick={() => setActiveTab(tab.id)}
              >
                {tab.label}
              </button>
            ))}
          </div>
        </div>

        {/* ── ボディ ── */}
        <div className={styles.body} role="tabpanel">
          {activeTab === 'basic' && <BasicSettingsTab />}
          {activeTab === 'char' && <CharCompositionTab />}
          {activeTab === 'frame' && <FrameTab />}
          {activeTab === 'formula' && <FormulaTab />}
          {activeTab === 'template' && <TemplateTextTab />}
        </div>

      </div>
    </FluentProvider>
  )
}

