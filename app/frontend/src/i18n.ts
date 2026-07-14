import i18n from 'i18next'
import { initReactI18next } from 'react-i18next'
import en from './locales/en'
import ru from './locales/ru'

const stored = localStorage.getItem('lang')
const initial = stored || (navigator.language.startsWith('ru') ? 'ru' : 'en')

i18n.use(initReactI18next).init({
  resources: {
    en: { translation: en },
    ru: { translation: ru },
  },
  lng: initial,
  fallbackLng: 'en',
  interpolation: { escapeValue: false },
})

export function setLanguage(lng: 'en' | 'ru') {
  localStorage.setItem('lang', lng)
  i18n.changeLanguage(lng)
}

export default i18n
