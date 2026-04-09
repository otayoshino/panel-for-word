// src/components/features/formula/FractionFormulaFeature.tsx
import { makeStyles, tokens } from '@fluentui/react-components'
import { StatusBar } from '../../shared/StatusBar'
import { SectionHeader } from '../../shared/SectionHeader'
import { useWordRun } from '../../../hooks/useWordRun'
import { makeOoxmlMath } from './_ooxmlMath'

type FracType = 'bar' | 'skw' | 'lin' | 'noBar'

const FRAC_ITEMS: {
  value: FracType
  label: string
  icon: React.ReactNode
}[] = [
  {
    value: 'bar',
    label: '縦積み',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '14px', fontWeight: '600', fontFamily: 'serif' }}>
        <span style={{ borderBottom: '1.5px solid currentColor', paddingBottom: '1px', lineHeight: '1.1', minWidth: '12px', textAlign: 'center' }}>x</span>
        <span style={{ lineHeight: '1.1', minWidth: '12px', textAlign: 'center' }}>y</span>
      </span>
    ),
  },
  {
    value: 'skw',
    label: '斜め',
    icon: (
      <span style={{ fontSize: '15px', fontWeight: '600', fontFamily: 'serif', fontStyle: 'italic' }}>
        x/y
      </span>
    ),
  },
  {
    value: 'lin',
    label: '線形',
    icon: (
      <span style={{ fontSize: '13px', fontWeight: '600', fontFamily: 'serif' }}>
        a/b
      </span>
    ),
  },
  {
    value: 'noBar',
    label: '分数（小）',
    icon: (
      <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: '1', gap: '1px', fontSize: '11px', fontWeight: '600', fontFamily: 'serif' }}>
        <span style={{ lineHeight: '1.1', minWidth: '10px', textAlign: 'center' }}>x</span>
        <span style={{ lineHeight: '1.1', minWidth: '10px', textAlign: 'center' }}>y</span>
      </span>
    ),
  },
]

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', width: '100%', gap: tokens.spacingVerticalS },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: '8px',
    width: '100%',
  },
  card: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '64px',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    gap: '5px',
    border: '1px solid #c5dcf5',
    backgroundColor: '#ffffff',
    transitionProperty: 'background-color, transform, box-shadow',
    transitionDuration: '0.15s',
    transitionTimingFunction: 'ease',
    outline: 'none',
    userSelect: 'none',
    color: '#1e4d8c',
    ':hover': {
      backgroundColor: '#e8f0fb',
      transform: 'scale(1.04)',
      boxShadow: '0 2px 8px rgba(30,77,140,0.15)',
    },
    ':focus-visible': {
      outline: '2px solid #1e4d8c',
      outlineOffset: '2px',
    },
    ':active': {
      transform: 'scale(0.97)',
    },
  },
  cardLabel: {
    fontSize: '10px',
    textAlign: 'center',
    color: '#0c3370',
    fontFamily: "'Yu Gothic', 'Meiryo', sans-serif",
    lineHeight: '1.2',
  },
})

export function FractionFormulaFeature() {
  const styles = useStyles()
  const { runWord, status } = useWordRun()

  const insertFraction = (fracType: FracType) =>
    runWord(async (context) => {
      const range = context.document.getSelection()
      const fPr = fracType !== 'bar' ? `<m:fPr><m:type m:val="${fracType}"/></m:fPr>` : ''
      const mathContent = `<m:f>${fPr}<m:num><m:r><m:t></m:t></m:r></m:num><m:den><m:r><m:t></m:t></m:r></m:den></m:f>`
      range.insertOoxml(makeOoxmlMath(mathContent), Word.InsertLocation.replace)
      await context.sync()
    })

  return (
    <div className={styles.root}>
      <SectionHeader title="分数" />
      <div className={styles.grid}>
        {FRAC_ITEMS.map((item) => (
          <button
            key={item.value}
            className={styles.card}
            onClick={() => insertFraction(item.value)}
            title={item.label}
          >
            <span>{item.icon}</span>
            <span className={styles.cardLabel}>{item.label}</span>
          </button>
        ))}
      </div>
      <StatusBar status={status} />
    </div>
  )
}

