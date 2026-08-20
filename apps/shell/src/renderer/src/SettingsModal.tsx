import { useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import type { UiTheme } from '../../shared/home-api'
import './settings.css'

// ── Settings modal (opened from the account menu) ─────────
// Two-pane dialog: section nav on the left, fields on the right.

// sorted by ISO 639 language code — native-script labels have no natural
// shared alphabet, so the code is the ordering key
const LANG_OPTIONS = [
  { value: 'ar', label: 'العربية' },
  { value: 'de', label: 'Deutsch' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'fr', label: 'Français' },
  { value: 'he', label: 'עברית' },
  { value: 'hi', label: 'हिन्दी' },
  { value: 'id', label: 'Bahasa Indonesia' },
  { value: 'it', label: 'Italiano' },
  { value: 'ja', label: '日本語' },
  { value: 'ko', label: '한국어' },
  { value: 'ms', label: 'Bahasa Melayu' },
  { value: 'nl', label: 'Nederlands' },
  { value: 'pl', label: 'Polski' },
  { value: 'pt', label: 'Português' },
  { value: 'ru', label: 'Русский' },
  { value: 'th', label: 'ไทย' },
  { value: 'zh', label: '简体中文' },
  { value: 'zh-TW', label: '繁體中文' },
] as const

const THEME_OPTIONS = [
  { value: 'system', labelKey: 'themeSystem' },
  { value: 'light', labelKey: 'themeLight' },
  { value: 'dark', labelKey: 'themeDark' },
] as const satisfies readonly { value: UiTheme; labelKey: StringKey }[]

const CHANNEL_OPTIONS = [
  { value: 'stable', labelKey: 'channelStable' },
  { value: 'beta', labelKey: 'channelBeta' },
] as const satisfies readonly { value: 'stable' | 'beta'; labelKey: StringKey }[]

function formatStars(n: number): string {
  if (n < 1000) return String(n)
  const k = n / 1000
  return `${k >= 100 ? Math.round(k) : (Math.round(k * 10) / 10).toString().replace(/\.0$/, '')}k`
}

type SectionId = 'ai' | 'general' | 'about'

const SECTIONS: readonly { id: SectionId; labelKey: string }[] = [
  { id: 'ai', labelKey: 'AI Provider' },
  { id: 'general', labelKey: '' },
  { id: 'about', labelKey: '' },
]

function SectionIcon({ id }: { id: SectionId }) {
  if (id === 'ai') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <circle cx="8" cy="5.2" r="2.9" stroke="currentColor" strokeWidth="1.3" />
        <path d="M2.7 13.6a5.5 5.5 0 0 1 10.6 0" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      </svg>
    )
  }
  if (id === 'general') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path d="M2 5h8M13 5h1M2 11h1M6 11h8" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
        <circle cx="11.5" cy="5" r="1.7" stroke="currentColor" strokeWidth="1.3" />
        <circle cx="4.5" cy="11" r="1.7" stroke="currentColor" strokeWidth="1.3" />
      </svg>
    )
  }
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6.3" stroke="currentColor" strokeWidth="1.3" />
      <path d="M8 7.4v3.4" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      <circle cx="8" cy="5.1" r="0.8" fill="currentColor" />
    </svg>
  )
}

function Field({ label, value, valueTitle, action }: { label: string; value: string; valueTitle?: string; action?: ReactNode }) {
  return (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-label">{label}</div>
        <div className="set-field-value" data-tip={valueTitle}>{value}</div>
      </div>
      {action}
    </div>
  )
}

interface AiProviderConfig {
  apiKey: string
  model: string
  baseUrl?: string
}

interface AiSettings {
  provider: string
  providers: Record<string, AiProviderConfig>
}

export interface SettingsModalProps {
  onClose: () => void
}

export function SettingsModal({ onClose }: SettingsModalProps) {
  const { lang, setLang, t } = useI18n()
  const [section, setSection] = useState<SectionId>('ai')
  const [theme, setTheme] = useState<UiTheme>('system')
  const [saveDir, setSaveDir] = useState('')
  const [channel, setChannel] = useState<'stable' | 'beta'>('stable')
  const [appVersion, setAppVersion] = useState('')
  const [githubStars, setGithubStars] = useState<number | null>(null)

  // AI settings
  const [aiSettings, setAiSettings] = useState<AiSettings>({ provider: 'prova', providers: {} })
  const [provaApiKey, setProvaApiKey] = useState('')

  useEffect(() => {
    let alive = true
    void window.aiOffice.getTheme?.().then((th) => { if (alive) setTheme(th) })
    void window.aiOffice.getDefaultSaveDir?.().then((dir) => { if (alive && dir) setSaveDir(dir) })
    void window.aiOffice.getUpdateChannel?.().then((ch) => { if (alive) setChannel(ch) })
    void window.aiOffice.getAppVersion?.().then((v) => { if (alive && v) setAppVersion(v) })
    void window.aiOffice.githubStars?.().then((n) => { if (alive && n !== null) setGithubStars(n) })

    // Load AI settings
    void (window as any).aiOfficeAiSettings?.getAiSettings?.().then((settings: AiSettings) => {
      if (!alive || !settings) return
      setAiSettings(settings)
      const prova = settings.providers?.prova
      if (prova) {
        setProvaApiKey(prova.apiKey ?? '')
      }
    })

    return () => { alive = false }
  }, [])

  useEffect(() => {
    const handler = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])

  const applyTheme = (next: UiTheme) => {
    setTheme(next)
    void window.aiOffice.setTheme(next)
    if (next === 'system') document.documentElement.removeAttribute('data-theme')
    else document.documentElement.setAttribute('data-theme', next)
  }

  const changeSaveDir = () => {
    void window.aiOffice.pickDefaultSaveDir?.().then((dir) => { if (dir) setSaveDir(dir) })
  }

  const saveAiSettings = async (provider: string, newSettings: AiSettings) => {
    setAiSettings(newSettings)
    await (window as any).aiOfficeAiSettings?.setAiSettings?.(newSettings)
  }

  const updateProva = async (field: 'baseUrl' | 'apiKey' | 'model', value: string) => {
    if (field === 'apiKey') setProvaApiKey(value)

    const updated = { ...aiSettings, provider: 'prova', providers: { ...aiSettings.providers } }
    updated.providers.prova = {
      baseUrl: 'https://llm.proxsis.com/api/v1',
      apiKey: field === 'apiKey' ? value : provaApiKey,
      model: 'PROVAOffice',
    }
    await saveAiSettings('prova', updated)
  }

  return (
    <div className="set-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose() }}>
      <div className="set-dialog" role="dialog" aria-modal="true" aria-label="Settings">
        <div className="set-header">
          <h2 className="set-title">Settings</h2>
          <button className="set-close" onClick={onClose} aria-label="Close">
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path d="M2 2l10 10M12 2L2 12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
            </svg>
          </button>
        </div>
        <div className="set-body">
          <nav className="set-nav" aria-label="Settings">
            {SECTIONS.map((s) => (
              <button key={s.id} className={`set-nav-item${section === s.id ? ' active' : ''}`} aria-current={section === s.id} onClick={() => setSection(s.id)}>
                <SectionIcon id={s.id} />
                {s.labelKey || t(s.labelKey as StringKey)}
              </button>
            ))}
          </nav>
          <div className="set-pane">
            {section === 'ai' && (
              <>
                <h3 className="set-pane-title">AI Provider</h3>

                <h4 style={{ fontSize: '13px', fontWeight: 600, margin: '12px 0 6px', color: 'var(--text)' }}>ProxsisLLM</h4>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="prova-apikey">API Key (License)</label>
                  </div>
                  <input id="prova-apikey" className="set-input set-input-wide" type="password" placeholder="Enter your license key" value={provaApiKey} onChange={(e) => void setProvaApiKey(e.target.value)} />
                </div>
                <div style={{ marginTop: '8px', display: 'flex', justifyContent: 'flex-end' }}>
                  <button className="set-btn primary" onClick={() => { void updateProva('apiKey', provaApiKey); onClose() }}>
                    Simpan
                  </button>
                </div>
              </>
            )}
            {section === 'general' && (
              <>
                <h3 className="set-pane-title">{t('setSecGeneral')}</h3>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-lang">{t('language')}</label>
                  </div>
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">{LANG_OPTIONS.find((o) => o.value === lang)?.label ?? lang}</span>
                    <select id="set-lang" className="set-select" value={lang} onChange={(e) => setLang(e.target.value as typeof lang)}>
                      {LANG_OPTIONS.map((opt) => (<option key={opt.value} value={opt.value}>{opt.label}</option>))}
                    </select>
                  </span>
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-theme">{t('theme')}</label>
                  </div>
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">{t(THEME_OPTIONS.find((o) => o.value === theme)?.labelKey ?? 'themeSystem')}</span>
                    <select id="set-theme" className="set-select" value={theme} onChange={(e) => applyTheme(e.target.value as UiTheme)}>
                      {THEME_OPTIONS.map((opt) => (<option key={opt.value} value={opt.value}>{t(opt.labelKey)}</option>))}
                    </select>
                  </span>
                </div>
                <Field label={t('saveLocation')} value={saveDir || '—'} valueTitle={saveDir} action={<button className="set-btn" onClick={changeSaveDir}>{t('setChange')}</button>} />
              </>
            )}
            {section === 'about' && (
              <>
                <h3 className="set-pane-title">{t('setSecAbout')}</h3>
                <Field label={t('versionLabel')} value={appVersion || '—'} />
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label" htmlFor="set-channel">{t('updateChannel')}</label>
                  </div>
                  <span className="set-select-wrap">
                    <span className="set-select-text" aria-hidden="true">{t(CHANNEL_OPTIONS.find((o) => o.value === channel)?.labelKey ?? 'channelStable')}</span>
                    <select id="set-channel" className="set-select" value={channel} onChange={(e) => { const next = e.target.value === 'beta' ? 'beta' : 'stable'; setChannel(next); void window.aiOffice.setUpdateChannel(next) }}>
                      {CHANNEL_OPTIONS.map((opt) => (<option key={opt.value} value={opt.value}>{t(opt.labelKey)}</option>))}
                    </select>
                  </span>
                </div>
                <Field label="GitHub" value={githubStars === null ? 'provaoffice' : `provaoffice · ★ ${formatStars(githubStars)}`} action={<button className="set-btn" onClick={() => void window.aiOffice.openGitHubRepo?.()}>{t('starOnGitHub')}</button>} />
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
