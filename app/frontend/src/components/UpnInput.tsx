import { useId } from 'react'
import { Input } from './ui'
import { useStore } from '../lib/store'

// UPN field with a datalist of tenant domains: after typing "name@" the user
// gets the tenant's domains as suggestions (prefilled on connect).
export function UpnInput({
  value,
  onChange,
  placeholder,
}: {
  value: string
  onChange: (v: string) => void
  placeholder?: string
}) {
  const { domains } = useStore()
  const listId = useId()
  const local = value.includes('@') ? value.slice(0, value.indexOf('@')) : value
  const options = domains.map((d) => `${local}@${d}`)

  return (
    <>
      <Input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder || 'user@contoso.com'}
        list={domains.length ? listId : undefined}
        autoComplete="off"
      />
      {domains.length > 0 && (
        <datalist id={listId}>
          {options.map((o) => (
            <option key={o} value={o} />
          ))}
        </datalist>
      )}
    </>
  )
}
