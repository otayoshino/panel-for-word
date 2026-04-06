import { MessageBar, MessageBarBody } from '@fluentui/react-components'
import type { Status } from '../../hooks/useWordRun'

interface StatusBarProps {
  status: Status | null
}

export function StatusBar({ status }: StatusBarProps) {
  if (!status) return null

  return (
    <MessageBar intent={status.type === 'warning' ? 'warning' : status.type}>
      <MessageBarBody>{status.message}</MessageBarBody>
    </MessageBar>
  )
}
