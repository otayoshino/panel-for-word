import { makeStyles, Text } from '@fluentui/react-components'

const useStyles = makeStyles({
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginTop: '10px',
    marginBottom: '7px',
  },
  label: {
    color: '#0c51a0',
    fontSize: '8px',
    fontWeight: '500',
    letterSpacing: '0.12em',
    whiteSpace: 'nowrap',
    fontFamily: "'Noto Sans JP', sans-serif",
    textTransform: 'uppercase',
  },
  line: {
    flex: 1,
    height: '1px',
    backgroundColor: '#c5dcf5',
  },
})

interface SectionHeaderProps {
  title: string
}

export function SectionHeader({ title }: SectionHeaderProps) {
  const styles = useStyles()
  return (
    <div className={styles.header}>
      <Text className={styles.label}>{title}</Text>
      <div className={styles.line} />
    </div>
  )
}
