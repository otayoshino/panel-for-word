import { useState, useRef } from 'react'

export type Status = { type: 'success' | 'error' | 'warning'; message: string }

const AUTO_CLEAR_MS = 4000

export function useWordRun() {
  const [status, setStatusRaw] = useState<Status | null>(null)
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const setStatus = (s: Status | null) => {
    if (timerRef.current) clearTimeout(timerRef.current)
    setStatusRaw(s)
    if (s !== null) {
      timerRef.current = setTimeout(() => setStatusRaw(null), AUTO_CLEAR_MS)
    }
  }

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

  // Word.runをラップせずtry-catchのみ（内部で自前のWord.runを複数呼ぶ場合に使用）
  const runRaw = async (action: () => Promise<void>) => {
    try {
      await action()
    } catch (e) {
      setStatus({
        type: 'error',
        message: `エラー: ${e instanceof Error ? e.message : String(e)}`,
      })
    }
  }

  return { runWord, runRaw, status, setStatus }
}
