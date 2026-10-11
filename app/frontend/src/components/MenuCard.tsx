import { useTranslation } from 'react-i18next'
import { Card, Button } from './ui'
import { useStore } from '../lib/store'
import { NAV_GROUPS, NAV_PINNED } from './Layout'
import { ALWAYS_SHOWN, ESSENTIAL_PAGES } from '../lib/navprefs'
import { pageEnabled } from '../lib/workspaces'
import type { PageId } from '../pages/registry'

// Which pages the sidebar shows. Hidden pages stay one Ctrl+K away, so a
// short menu costs no features.
export function MenuCard() {
  const { t } = useTranslation()
  const { hiddenPages, setHiddenPages, workspaces } = useStore()
  // The preset covers every page, also those of a part switched off now.
  const everyPage = [...NAV_PINNED, ...NAV_GROUPS.flatMap((g) => g.items)].map((it) => it.id)
  const toggle = (id: PageId, show: boolean) =>
    setHiddenPages(show ? hiddenPages.filter((p) => p !== id) : [...hiddenPages, id])
  const sections = [{ key: '', items: NAV_PINNED }, ...NAV_GROUPS]

  return (
    <Card title={t('settings.menu.title')}>
      <p className="text-xs text-[var(--text-faint)]">{t('settings.menu.hint')}</p>
      <div className="mt-3 flex flex-wrap gap-2">
        <Button variant="subtle" onClick={() => setHiddenPages(everyPage.filter((id) => !ESSENTIAL_PAGES.includes(id)))}>
          {t('settings.menu.essentials')}
        </Button>
        <Button variant="ghost" onClick={() => setHiddenPages([])}>{t('settings.menu.everything')}</Button>
      </div>
      <div className="mt-3 grid grid-cols-1 gap-x-4 gap-y-3 sm:grid-cols-2">
        {sections.map((g) => {
          const items = g.items.filter((it) => pageEnabled(it.id, workspaces))
          if (items.length === 0) return null
          return (
            <div key={g.key || 'pinned'}>
              {g.key && <p className="mb-1 text-[11px] font-semibold uppercase tracking-wider text-[var(--text-faint)]">{t(g.key)}</p>}
              {items.map((it) => (
                <label key={it.id} className="flex items-center gap-2 py-0.5 text-sm">
                  <input type="checkbox" checked={!hiddenPages.includes(it.id)} disabled={ALWAYS_SHOWN.includes(it.id)}
                    onChange={(e) => toggle(it.id, e.target.checked)} />
                  <span className="text-[var(--text-faint)]">{it.icon}</span>
                  {t(it.key)}
                </label>
              ))}
            </div>
          )
        })}
      </div>
    </Card>
  )
}
