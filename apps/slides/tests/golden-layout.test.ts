import { describe, expect, it } from 'vitest'
import { openPptx } from '@prova/pptx-engine'
import { buildRenderSlide } from '@prova/pptx-render'
import { auditSlideLayout } from '../src/renderer/ai/layout-audit'
import { slideContentToPptxBytes } from '../src/main/html-to-pptx-local'
import type { SlideContent } from '../src/renderer/ai/slide-templates'
import { THEME_PRESETS } from '../src/renderer/ai/slide-templates'

/**
 * Golden layout test: every layout is built with WORST-CASE content (at the
 * density limits the prompt allows) and the deterministic layout audit must
 * find ZERO issues. This is the "guaranteed template" contract — any layout
 * change that would produce an overflowing/colliding slide fails here, before
 * it ever reaches a user.
 */

const FIT_WIDTH_PX = 1280

/** A valid 1×1 PNG data URL — downloadImageToDataUrl fetches data: URLs fine, so the
 *  image-present (narrow content) layout path is actually exercised by the audit. */
const PHOTO = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** Worst-case content per layout, at the density limits the prompt allows. */
const CASES: Array<{ name: string; slide: SlideContent }> = [
  {
    name: 'title_content (5 bullets, max length)',
    slide: {
      title: 'Pertumbuhan Ekonomi Digital Asia Tenggara 2040',
      eyebrow: '01 · PENGANTAR',
      intro: 'Ekonomi digital kawasan diproyeksikan tumbuh pesat dalam dua dekade ke depan.',
      content: [
        'Ekonomi digital ASEAN diproyeksikan mencapai 1 triliun dolar pada 2030',
        'Pengguna internet tumbuh dari 400 juta menjadi 700 juta pengguna aktif',
        'E-commerce mendominasi 60 persen dari total transaksi digital kawasan',
        'Fintech dan pembayaran digital tumbuh lebih dari 25 persen per tahun',
        'Kesenjangan infrastruktur antara kota dan desa masih menjadi tantangan',
      ],
      layout: 'title_content',
      imageUrl: PHOTO,
      imageSide: 'right',
    },
  },
  {
    name: 'two_column (dense columns)',
    slide: {
      title: 'Kenyataan dan Respons Iklim 2040',
      content: [
        '• Suhu global menembus 1,5°C — perkiraan jalur saat ini 1,6–1,8°C di atas pra-industri\n• Gelombang panas, banjir, dan kekeringan ekstrem makin sering\n• Kenaikan muka laut mengancam kota pesisir\n• Gagal panen berulang mendorong 216 juta migran iklim',
        '• Adaptasi jadi industri utama: tangkal perubahan iklim, peringatan dini AI\n• Carbon removal skala gigaton diperlukan untuk mencapai net-zero 2050\n• Kota 15-menit dan infrastruktur hijau jadi standar perencanaan',
      ],
      layout: 'two_column',
    },
  },
  {
    name: 'cards (3 cards, max desc)',
    slide: {
      title: 'Tiga Pilar Transformasi Digital',
      content: [
        { title: 'Kesehatan Presisi', desc: 'Diagnosis berbasis AI dan genomik personal memangkas biaya perawatan hingga 40 persen sekaligus mempercepat penemuan obat baru.' },
        { title: 'Pangan Pintar', desc: 'Pertanian presisi, protein alternatif, dan rantai pasok transparan menjawab ketahanan pangan 9 miliar jiwa global.' },
        { title: 'Kota Otonom', desc: 'Transportasi otonom, jaringan energi terdistribusi, dan tata kelola data warga membentuk kota yang adaptif.' },
      ],
      layout: 'cards',
      imageUrl: PHOTO,
      imageSide: 'left',
    },
  },
  {
    name: 'stats (4 items, max lengths)',
    slide: {
      title: 'Angka-Angka Kunci 2040',
      eyebrow: '02 · RANGKASAN DATA',
      content: [
        { big: '9,2 M', desc: 'Jiwa dunia (UN WPP)' },
        { big: '63%', desc: 'Tinggal di kota' },
        { big: '$15,7 T', desc: 'Ekonomi AI (PwC)' },
        { big: '945 TWh', desc: 'Listrik data center' },
      ],
      layout: 'stats',
      imageUrl: PHOTO,
      imageSide: 'right',
    },
  },
  {
    name: 'timeline (5 steps, max lengths)',
    slide: {
      title: 'Roadmap Transformasi 2025–2040',
      content: [
        { title: '2025–2027', desc: 'Fondasi data & infrastruktur awal' },
        { title: '2028–2030', desc: 'Adopsi AI lintas sektor menanjak' },
        { title: '2031–2034', desc: 'Otomasi proses inti industri' },
        { title: '2035–2037', desc: 'Ekonomi otonom skala regional' },
        { title: '2038–2040', desc: 'Maturitas & regulasi matang' },
      ],
      layout: 'timeline',
    },
  },
  {
    name: 'quote (long quote)',
    slide: {
      title: 'Kutipan',
      content: ['Teknologi tidak pernah menggantikan manusia sepenuhnya — yang terjadi adalah manusia yang menguasai teknologi menggantikan yang tidak. Kuncinya adalah kecepatan belajar, bukan kecepatan mesin.'],
      intro: '— Laporan Prospek Pekerjaan Dunia, 2040',
      layout: 'quote',
    },
  },
  {
    name: 'big_number (long caption)',
    slide: {
      title: 'Sorotan Utama',
      content: ['$15,7 T'],
      intro: 'Kontribusi ekonomi AI terhadap produk domestik bruto global pada 2040 menurut proyeksi konsultan besar',
      layout: 'big_number',
    },
  },
  {
    name: 'comparison (2 dense panels)',
    slide: {
      title: 'PMBOK 7 vs PMBOK 8: Apa yang Berubah?',
      content: [
        { title: 'PMBOK 7 (2021)', desc: '12 prinsip & 8 domain kinerja. Sangat konseptual dan abstrak. Model, metode, artefak dipisah terpisah. Minim panduan praktis langsung.' },
        { title: 'PMBOK 8 (2025–26)', desc: '6 prinsip & 8 praktik yang lebih konkret. Struktur Principles – Performance – Practices. Integrasi agile/hybrid lebih dalam. Topik baru: AI digital, sustainability.' },
      ],
      layout: 'comparison',
      imageUrl: PHOTO,
      imageSide: 'right',
    },
  },
  {
    name: 'rows (5 rows, max lengths)',
    slide: {
      title: 'Lima Langkah Persiapan Organisasi',
      content: [
        { title: 'Asesmen kesiapan', desc: 'Ukur kematangan data, proses, dan kompetensi tim terhadap target transformasi' },
        { title: 'Strategi data', desc: 'Tetapkan tata kelola, kualitas, dan arsitektur data sebagai fondasi semua inisiatif' },
        { title: 'Pilot terukur', desc: 'Mulai dari kasus bernilai tinggi dengan metrik keberhasilan yang jelas' },
        { title: 'Skala bertahap', desc: 'Replikasi ke unit lain dengan pembelajaran dari pilot' },
        { title: 'Kemampuan berkelanjutan', desc: 'Bangun akademi internal dan komunitas praktik' },
      ],
      layout: 'rows',
    },
  },
  {
    name: 'agenda (6 items)',
    slide: {
      title: 'Agenda Pembahasan',
      content: ['Pengantar & Konteks', 'Demografi Dunia 2040', 'Ekonomi Digital & AI', 'Energi & Iklim', 'Geopolitik Baru', 'Implikasi untuk Indonesia'],
      layout: 'agenda',
    },
  },
  {
    name: 'business_case full (banner + 2 cards + table, max)',
    slide: {
      title: 'Rekomendasi: PostgreSQL 18 untuk Kasus Ini',
      layout: 'business_case',
      bc: {
        variant: 'full',
        badge: 'REKOMENDASI AKHIR',
        key_message: 'Untuk profil DB 2,4 GB / 56 tabel: PostgreSQL 18 adalah pilihan paling pas tanpa kompromi apa pun.',
        background: 'Diskusi menelusuri sumber pendapatan Oracle (lisensi per core ± support 22%/tahun ± cloud), klaim gratis total Postgres, dan kelayakan skala PG18.',
        business_problem: 'Keputusan memilih engine sering terdistorsi persepsi Oracle lebih kuat atau lebih cepat, padahal di skala kecil semua engine terasa identik.',
        solution: [
          'Headroom PG18 untuk kasus ini ±1.000x lipat — titik butuh arsitektur khusus baru muncul sekitar 1–2 TB workload OLTP biasa.',
          'Tetap jalankan disiplin dasar: backup terjadwal + uji restore, pantangi tabel dominan via pg_total_relation_size.',
          'Distributed SQL atau Oracle baru masuk akal >10–50 TB — bagi mayoritas aplikasi tidak akan pernah terjadi.',
        ],
        table_title: 'ORACLE VS POSTGRESQL — RINGKASAN',
        table: {
          headers: ['Dimensi', 'Oracle', 'PostgreSQL 18', 'Delta'],
          rows: [
            ['Lisensi', '$37.585/core + opsi', '$0', '$0'],
            ['Support tahunan', '±22% nilai lisensi, naik terus', '$0 (opsional: EDB/Percona)', 'Gratis'],
            ['Audit', 'Risiko temuan jutaan dolar', 'Tidak ada', 'Nihil'],
            ['Skala nyaman Anda', 'Overkill total', 'Sampai ±1–2 TB tanpa arsitektur khusus', '1000x'],
          ],
        },
      },
    },
  },
  {
    name: 'business_case solution (dark card + 3 benefits)',
    slide: {
      title: 'Solusi: Moratorium Bersyarat + Sertifikasi Bertahap',
      layout: 'business_case',
      bc: {
        variant: 'solution',
        badge: 'SOLUSI',
        key_message: 'Moratorium bersyarat di daerah berinsiden, 6 pilar perbaikan dalam 90 hari, lanjutkan hanya untuk SPPG lolos sertifikasi.',
        solution: [
          'Pendek (0–3 bulan): moratorium bersyarat, investigasi KLB tuntas BPOM dan Polri, evaluasi distribusi sesuai mandat Menko Pangan.',
          'Menengah (3–12 bulan): Perpres atau PP standar MBG, sertifikasi higiene 100 persen SPPG, dashboard real-time.',
        ],
        benefits: [
          { title: 'Hemat', desc: 'Penghematan 40% biaya lisensi tahunan dari konsolidasi vendor.' },
          { title: 'Cepat', desc: 'Siklus laporan 12 hari kerja jadi 2–3 hari dengan validasi otomatis.' },
          { title: 'Aman', desc: 'Audit trail lengkap setiap angka tersitasi ke dokumen sumber.' },
        ],
      },
    },
  },
  {
    name: 'business_case metrics (3 numbers + 2 cards)',
    slide: {
      title: 'Dampak: Angka yang Harus Dikejar',
      layout: 'business_case',
      bc: {
        variant: 'metrics',
        badge: 'KPI',
        key_message: 'Tiga target ukur yang dapat langsung diadopsi BGN, BPOM, dan pemerintah daerah dalam 6–12 bulan ke depan.',
        metrics: [
          { big: '4 jam', desc: 'Batas jeda masak-ke-konsumsi sesuai standar BPOM' },
          { big: '100%', desc: 'SPPG laik higiene, halal, dan aman pangan' },
          { big: '21', desc: 'Program pengawasan BPOM 2027 diawasi eksekusinya' },
        ],
        challenge: 'Keracunan berulang di lima daerah dengan standar dapur dan pengawasan yang belum memadai.',
        impact: 'Kepercayaan publik terhadap program pulih dan ekspansi bisa dilanjutkan dengan aman.',
      },
    },
  },
]

describe('golden layout audit (worst-case content, all presets)', () => {
  for (const presetName of Object.keys(THEME_PRESETS)) {
    it(`all layouts pass the audit with theme preset "${presetName}"`, async () => {
      const slides = CASES.map((c) => ({ ...c.slide, theme: THEME_PRESETS[presetName] }))
      const { bytes, imageFailures } = await slideContentToPptxBytes(slides)
      // All photos use a valid data URL — none may fail to embed.
      expect(imageFailures).toEqual([])
      const opened = await openPptx(bytes)
      opened.deck.slides.forEach((s, i) => {
        const rendered = buildRenderSlide(s, opened.deck.size, {
          fitWidthPx: FIT_WIDTH_PX,
          media: () => undefined,
          slideNo: i + 1,
        })
        const issues = auditSlideLayout(rendered)
        expect(issues, `${CASES[i]!.name} [${presetName}]: ${issues.join(' | ')}`).toEqual([])
      })
    })
  }
})

describe('contrast audit', () => {
  it('flags white text on a light card tint (the IFRS regression)', async () => {
    // Simulates the model drawing custom cards via execute_slide_script: a light
    // tinted rect with WHITE text — the exact unreadable pattern from the report.
    const pptx = new (await import('pptxgenjs')).default()
    pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 })
    pptx.layout = 'WIDE'
    const s = pptx.addSlide()
    s.background = { color: 'FFFFFF' }
    s.addShape('rect', { x: 0.7, y: 1.5, w: 3.9, h: 3.6, fill: { color: '1B4332', transparency: 92 } })
    s.addText('IFRS S1', { x: 0.95, y: 1.75, w: 3.4, h: 1.1, fontSize: 21, color: 'FFFFFF', bold: true })
    const bytes = new Uint8Array(await pptx.write({ outputType: 'arraybuffer' }) as ArrayBuffer)
    const opened = await openPptx(bytes)
    const rendered = buildRenderSlide(opened.deck.slides[0]!, opened.deck.size, {
      fitWidthPx: FIT_WIDTH_PX,
      media: () => undefined,
      slideNo: 1,
    })
    const issues = auditSlideLayout(rendered)
    expect(issues.some((i) => i.startsWith('Low contrast'))).toBe(true)
  })

  it('passes white text on a dark panel (legitimate design)', async () => {
    const pptx = new (await import('pptxgenjs')).default()
    pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 })
    pptx.layout = 'WIDE'
    const s = pptx.addSlide()
    s.background = { color: 'FFFFFF' }
    s.addShape('rect', { x: 0.7, y: 1.5, w: 3.9, h: 3.6, fill: { color: '1B4332' } })
    s.addText('IFRS S1', { x: 0.95, y: 1.75, w: 3.4, h: 1.1, fontSize: 21, color: 'FFFFFF', bold: true })
    const bytes = new Uint8Array(await pptx.write({ outputType: 'arraybuffer' }) as ArrayBuffer)
    const opened = await openPptx(bytes)
    const rendered = buildRenderSlide(opened.deck.slides[0]!, opened.deck.size, {
      fitWidthPx: FIT_WIDTH_PX,
      media: () => undefined,
      slideNo: 1,
    })
    expect(auditSlideLayout(rendered)).toEqual([])
  })
})