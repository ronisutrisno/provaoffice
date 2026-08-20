export interface SlideTheme {
  primary: string
  secondary: string
  accent: string
  background: string
  title: string
  text: string
  footerText?: string
}

export const DEFAULT_THEME: SlideTheme = {
  primary: '#0D2137',
  secondary: '#1A73E8',
  accent: '#00BFA5',
  background: '#FFFFFF',
  title: '#0D2137',
  text: '#333333',
  footerText: '#999999',
}

export interface SlideContent {
  title: string
  subtitle?: string
  eyebrow?: string
  intro?: string
  content: string[] | Array<{ title: string; desc?: string }> | Array<{ big: string; desc: string }>
  layout: 'title' | 'title_content' | 'two_column' | 'cards' | 'rows' | 'stats' | 'section_header' | 'closing' | 'agenda' | 'blank'
  imageUrl?: string
  theme?: Partial<SlideTheme>
}

const W = 1280
const H = 720

function t(theme: SlideTheme, override?: Partial<SlideTheme>): SlideTheme {
  return { ...theme, ...override }
}

function wrapSlide(body: string, theme: SlideTheme): string {
  return `<div style="width:${W}px;height:${H}px;background:${theme.background};font-family:'Segoe UI',system-ui,-apple-system,sans-serif;color:${theme.text};overflow:hidden;position:relative;">${body}</div>`
}

function footerBar(text: string, theme: SlideTheme): string {
  return `<div style="position:absolute;bottom:0;left:0;right:0;height:32px;background:${theme.primary};display:flex;align-items:center;padding:0 40px;font-size:11px;color:${theme.footerText ?? '#aaa'};">${text}</div>`
}

function eyebrowHtml(text: string, theme: SlideTheme): string {
  return `<div style="font-size:12px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:${theme.accent};margin-bottom:8px;">${text}</div>`
}

function titleHtml(title: string, size = 40, color?: string, theme?: SlideTheme): string {
  return `<h1 style="margin:0;font-size:${size}px;font-weight:700;line-height:1.2;color:${color ?? theme?.title ?? '#0D2137'};">${title}</h1>`
}

function bulletsHtml(items: string[], theme: SlideTheme): string {
  return `<ul style="margin:0;padding:0 0 0 20px;list-style:none;">${items.map(i => `<li style="position:relative;padding:6px 0 6px 18px;font-size:16px;line-height:1.5;color:${theme.text};"><span style="position:absolute;left:0;top:10px;width:6px;height:6px;background:${theme.accent};border-radius:50%;"></span>${i}</li>`).join('')}</ul>`
}

function cardsHtml(items: Array<{ title: string; desc?: string }>, theme: SlideTheme): string {
  const w = Math.floor((W - 80 - 20 * (items.length - 1)) / items.length)
  return `<div style="display:flex;gap:20px;padding:0 40px;">${items.map((it, i) => `<div style="flex:1;background:${theme.primary}08;border:1px solid ${theme.primary}15;border-radius:12px;padding:24px 20px;">
    <div style="width:8px;height:8px;background:${theme.accent};border-radius:50%;margin-bottom:12px;"></div>
    <div style="font-size:16px;font-weight:700;color:${theme.primary};margin-bottom:8px;">${it.title}</div>
    ${it.desc ? `<div style="font-size:13px;line-height:1.5;color:${theme.text};">${it.desc}</div>` : ''}
  </div>`).join('')}</div>`
}

function rowsHtml(items: Array<{ title: string; desc?: string }>, theme: SlideTheme): string {
  return `<div style="padding:0 40px;">${items.map(it => `<div style="display:flex;align-items:flex-start;gap:12px;padding:10px 0;border-bottom:1px solid ${theme.primary}10;">
    <div style="width:8px;height:8px;background:${theme.accent};border-radius:50%;margin-top:6px;flex-shrink:0;"></div>
    <div><div style="font-size:15px;font-weight:600;color:${theme.primary};">${it.title}</div>
    ${it.desc ? `<div style="font-size:13px;color:${theme.text};margin-top:2px;">${it.desc}</div>` : ''}
    </div>
  </div>`).join('')}</div>`
}

function statsHtml(items: Array<{ big: string; desc: string }>, theme: SlideTheme): string {
  const w = Math.floor((W - 80 - 20 * (items.length - 1)) / items.length)
  return `<div style="display:flex;gap:20px;padding:0 40px;">${items.map(it => `<div style="flex:1;text-align:center;padding:24px 16px;">
    <div style="font-size:42px;font-weight:800;color:${theme.secondary};line-height:1;">${it.big}</div>
    <div style="font-size:13px;color:${theme.text};margin-top:10px;line-height:1.3;">${it.desc}</div>
  </div>`).join('')}</div>`
}

function twoColumnHtml(left: string, right: string, theme: SlideTheme): string {
  return `<div style="display:flex;gap:40px;padding:0 40px;">
    <div style="flex:1;font-size:15px;line-height:1.6;color:${theme.text};">${left}</div>
    <div style="flex:1;font-size:15px;line-height:1.6;color:${theme.text};">${right}</div>
  </div>`
}

function imageHtml(url: string, w: number, h: number): string {
  return `<img src="${url}" style="width:${w}px;height:${h}px;object-fit:cover;border-radius:8px;" onerror="this.style.display='none'" />`
}

export function generateCoverSlide(title: string, subtitle: string, theme: SlideTheme): string {
  return wrapSlide(`
    <div style="position:absolute;inset:0;background:linear-gradient(135deg,${theme.primary} 0%,${theme.secondary} 100%);display:flex;flex-direction:column;justify-content:center;padding:80px;">
      <div style="font-size:13px;font-weight:600;letter-spacing:3px;text-transform:uppercase;color:${theme.accent};margin-bottom:20px;">PRESENTATION</div>
      <h1 style="margin:0;font-size:48px;font-weight:800;color:#FFFFFF;line-height:1.15;">${title}</h1>
      ${subtitle ? `<div style="font-size:18px;color:rgba(255,255,255,0.7);margin-top:16px;">${subtitle}</div>` : ''}
      <div style="margin-top:auto;width:60px;height:4px;background:${theme.accent};border-radius:2px;"></div>
    </div>`, theme)
}

export function generateContentSlide(slide: SlideContent, slideNum: number, theme: SlideTheme): string {
  const th = t(theme, slide.theme)
  let body = ''

  if (slide.eyebrow) body += eyebrowHtml(slide.eyebrow, th)

  if (slide.layout === 'title_content' || slide.layout === 'title') {
    body += titleHtml(slide.title, 36, undefined, th)
    if (slide.intro) body += `<p style="font-size:14px;color:${th.text};margin:12px 0 20px;line-height:1.5;">${slide.intro}</p>`
    if (Array.isArray(slide.content) && slide.content.length > 0) {
      if (typeof slide.content[0] === 'string') {
        body += `<div style="padding:0 40px;">${bulletsHtml(slide.content as string[], th)}</div>`
      }
    }
  } else if (slide.layout === 'cards') {
    body += titleHtml(slide.title, 32, undefined, th)
    if (slide.intro) body += `<p style="font-size:14px;color:${th.text};margin:12px 0 20px;">${slide.intro}</p>`
    if (Array.isArray(slide.content) && typeof slide.content[0] === 'object' && 'title' in (slide.content[0] as any)) {
      body += cardsHtml(slide.content as Array<{ title: string; desc?: string }>, th)
    }
  } else if (slide.layout === 'rows') {
    body += titleHtml(slide.title, 32, undefined, th)
    if (Array.isArray(slide.content) && typeof slide.content[0] === 'object' && 'title' in (slide.content[0] as any)) {
      body += `<div style="margin-top:16px;">${rowsHtml(slide.content as Array<{ title: string; desc?: string }>, th)}</div>`
    }
  } else if (slide.layout === 'stats') {
    body += titleHtml(slide.title, 32, undefined, th)
    if (Array.isArray(slide.content) && typeof slide.content[0] === 'object' && 'big' in (slide.content[0] as any)) {
      body += `<div style="margin-top:24px;">${statsHtml(slide.content as Array<{ big: string; desc: string }>, th)}</div>`
    }
  } else if (slide.layout === 'two_column') {
    body += titleHtml(slide.title, 32, undefined, th)
    if (Array.isArray(slide.content) && slide.content.length >= 2) {
      const left = typeof slide.content[0] === 'string' ? slide.content[0] : JSON.stringify(slide.content[0])
      const right = typeof slide.content[1] === 'string' ? slide.content[1] : JSON.stringify(slide.content[1])
      body += `<div style="margin-top:16px;">${twoColumnHtml(left, right, th)}</div>`
    }
  } else if (slide.layout === 'section_header') {
    body = `<div style="position:absolute;inset:0;background:${th.primary};display:flex;flex-direction:column;justify-content:center;padding:80px;">
      ${slide.eyebrow ? `<div style="font-size:13px;font-weight:600;letter-spacing:3px;text-transform:uppercase;color:${th.accent};margin-bottom:16px;">${slide.eyebrow}</div>` : ''}
      <h1 style="margin:0;font-size:44px;font-weight:800;color:#FFFFFF;line-height:1.15;">${slide.title}</h1>
      ${slide.intro ? `<div style="font-size:16px;color:rgba(255,255,255,0.7);margin-top:16px;">${slide.intro}</div>` : ''}
      <div style="margin-top:24px;width:60px;height:4px;background:${th.accent};border-radius:2px;"></div>
    </div>`
    return wrapSlide(body, th)
  } else if (slide.layout === 'closing') {
    body = `<div style="position:absolute;inset:0;background:linear-gradient(135deg,${th.primary} 0%,${th.secondary} 100%);display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;padding:80px;">
      <h1 style="margin:0;font-size:48px;font-weight:800;color:#FFFFFF;">${slide.title || 'Terima Kasih'}</h1>
      ${slide.subtitle ? `<div style="font-size:18px;color:rgba(255,255,255,0.7);margin-top:16px;">${slide.subtitle}</div>` : ''}
      <div style="margin-top:24px;width:60px;height:4px;background:${th.accent};border-radius:2px;"></div>
    </div>`
    return wrapSlide(body, th)
  } else if (slide.layout === 'agenda') {
    body += titleHtml(slide.title, 32, undefined, th)
    if (Array.isArray(slide.content) && typeof slide.content[0] === 'string') {
      body += `<div style="margin-top:16px;padding:0 40px;">${(slide.content as string[]).map((item, i) => `
        <div style="display:flex;align-items:center;gap:16px;padding:12px 0;border-bottom:1px solid ${th.primary}10;">
          <div style="width:32px;height:32px;background:${th.secondary};border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff;font-size:14px;font-weight:700;">${i + 1}</div>
          <div style="font-size:16px;color:${th.text};">${item}</div>
        </div>`).join('')}</div>`
    }
  }

  const bodyWithPad = `<div style="padding:40px 0 40px;">${body}</div>`
  const withImage = slide.imageUrl
    ? `<div style="display:flex;gap:24px;padding:0 40px;"><div style="flex:1;">${bodyWithPad}</div><div style="flex:0 0 400px;display:flex;align-items:center;">${imageHtml(slide.imageUrl, 400, 300)}</div></div>`
    : bodyWithPad

  return wrapSlide(`
    <div style="padding:32px 0 0;">
      ${withImage}
    </div>
    ${footerBar(`Slide ${slideNum}`, th)}
  `, th)
}

export function generateAgendaSlide(title: string, items: string[], theme: SlideTheme): string {
  return generateContentSlide({
    title,
    content: items,
    layout: 'agenda',
  }, 0, theme)
}

export function generateClosingSlide(title: string, subtitle: string, theme: SlideTheme): string {
  return generateContentSlide({
    title: title || 'Terima Kasih',
    subtitle,
    content: [],
    layout: 'closing',
  }, 0, theme)
}
