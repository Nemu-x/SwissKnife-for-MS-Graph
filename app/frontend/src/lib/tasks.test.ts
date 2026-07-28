import { describe, it, expect } from 'vitest'
import { TASKS, TASK_GROUPS } from './tasks'
import en from '../locales/en'
import ru from '../locales/ru'

// The palette renders every task by its i18n key, so a task without a label is
// an invisible dead entry. Ids also double as those keys — they must be unique.
describe('task registry', () => {
  it('every task has a label in both locales', () => {
    const missing = TASKS.flatMap((t) => [
      (en.tasks as Record<string, string>)[t.id] ? [] : [`en.tasks.${t.id}`],
      (ru.tasks as Record<string, string>)[t.id] ? [] : [`ru.tasks.${t.id}`],
    ].flat())
    expect(missing).toEqual([])
  })

  it('task ids are unique', () => {
    const seen = new Set<string>()
    const dupes = TASKS.filter((t) => (seen.has(t.id) ? true : (seen.add(t.id), false))).map((t) => t.id)
    expect(dupes).toEqual([])
  })

  it('every task belongs to a rendered group and carries search terms', () => {
    expect(TASKS.filter((t) => !TASK_GROUPS.includes(t.group)).map((t) => t.id)).toEqual([])
    expect(TASKS.filter((t) => !t.keywords.trim()).map((t) => t.id)).toEqual([])
  })
})
