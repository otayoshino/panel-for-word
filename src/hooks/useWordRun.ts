import { useState } from 'react'

export type Status = { type: 'success' | 'error' | 'warning'; message: string }

export function useWordRun() {
  const [status, setStatus] = useState<Status | null>(null)

  const runWord = async (action: (context: Word.RequestContext) => Promise<void>) => {
    try {
      await Word.run(async (context) => {
        await action(context)
      })
    } catch (e) {
      setStatus({
        type: 'error',
        message: `エラー: ${e instanceof Error ? e.message : String(e)}`,
      })
    }
  }

  return { runWord, status, setStatus }
}
