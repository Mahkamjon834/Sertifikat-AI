import { useState } from 'react'
import './App.css'
import { db } from './firebase/config';
import { addDoc, collection, doc, deleteDoc, updateDoc } from 'firebase/firestore';
import useCollback from './hook/useCollback';

function App() {
  const { data: callbackData } = useCollback('colculation')
  const calculation = callbackData || []

  const [darkMode, setDarkMode] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [fullName, setFullName] = useState('')
  const [surname, setSurname] = useState('')
  const [course, setCourse] = useState('Frontend')
  const [status, setStatus] = useState('Yakunlangan')
  const [mentorName, setMentorName] = useState('')
  const [certificateDate, setCertificateDate] = useState('')
  const [showForm, setShowForm] = useState(false)
  const [certificatePreview, setCertificatePreview] = useState(null)
  const [selectedCourse, setSelectedCourse] = useState('Frontend')

  const [editId, setEditId] = useState(null)
  const [editFullName, setEditFullName] = useState('')
  const [editSurname, setEditSurname] = useState('')
  const [editCourse, setEditCourse] = useState('Frontend')
  const [editStatus, setEditStatus] = useState('Yakunlangan')
  const [editMentorName, setEditMentorName] = useState('')
  const [editCertificateDate, setEditCertificateDate] = useState('')

  // Metrics
  const totalCount = calculation.length
  const frontendCount = calculation.filter(i => (i.course || i.text || 'Frontend') === 'Frontend').length
  const backendCount = calculation.filter(i => (i.course || i.text || 'Frontend') === 'Backend').length
  const englishCount = calculation.filter(i => (i.course || i.text || 'Frontend') === 'English').length

  const deleteDocument = async (id) => {
    deleteDoc(doc(db, 'colculation', id))
      .then(() => console.log('success'))
      .catch((error) => alert(error.message))
  }

  const editDocument = (id) => {
    const item = calculation.find((d) => d.id === id)
    if (!item) return

    setEditId(id)
    setEditFullName(item.fullName || item.title || '')
    setEditSurname(item.surname || '')
    setEditCourse(item.course || item.text || 'Frontend')
    setEditStatus(item.status || 'Yakunlangan')
    setEditMentorName(item.mentorName || '')
    setEditCertificateDate(item.certificateDate || '')
  }

  const handleEditSubmit = async (e) => {
    e.preventDefault()
    try {
      await updateDoc(doc(db, 'colculation', editId), {
        fullName: editFullName.trim(),
        surname: editSurname.trim(),
        course: editCourse,
        status: editStatus,
        mentorName: editMentorName.trim(),
        certificateDate: editCertificateDate,
      })
      setEditId(null)
    } catch (error) {
      alert(error.message)
    }
  }

  const cancelEdit = () => {
    setEditId(null)
  }

  const handlePrintCertificate = (item) => {
    setCertificatePreview(item)
  }

  const handlePrintDownload = () => {
    if (certificatePreview) {
      window.print()
    }
  }

  const handleDownloadPDF = async () => {
    if (!certificatePreview) return
    const element = document.querySelector('.certificate-container')
    if (!element) return

    const html2pdf = window.html2pdf || (await import('html2pdf.js')).default
    const options = {
      margin: 0,
      filename: `${certificatePreview.fullName || 'Sertifikat'}_${certificatePreview.surname || ''}_sertifikat.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 3, useCORS: true, logging: false },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
    }
    html2pdf().set(options).from(element).save()
  }

  const handleDownloadWord = async () => {
    if (!certificatePreview) return
    const element = document.querySelector('.certificate-container')
    if (!element) return

    try {
      const html2canvas = (await import('html2canvas')).default
      const canvas = await html2canvas(element, { scale: 2, useCORS: true })
      
      const imageArrayBuffer = await new Promise((resolve) => {
        canvas.toBlob((blob) => {
          blob.arrayBuffer().then(resolve)
        }, 'image/png')
      })
      
      const { Document, Packer, Paragraph, ImageRun, PageOrientation } = await import('docx')
      const doc = new Document({
        sections: [{
          properties: {
            page: {
              size: {
                orientation: PageOrientation.LANDSCAPE,
                width: 16838,
                height: 11906,
              },
              margin: { top: 0, right: 0, bottom: 0, left: 0 }
            }
          },
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageArrayBuffer,
                  transformation: {
                    width: 794,
                    height: 561,
                  },
                }),
              ],
            }),
          ],
        }],
      })
      
      const blob = await Packer.toBlob(doc)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `${certificatePreview.fullName || 'Sertifikat'}_${certificatePreview.surname || ''}_sertifikat.docx`
      a.click()
      URL.revokeObjectURL(url)
    } catch (err) {
      console.error('Word export error:', err)
      alert('Word eksportida xatolik yuz berdi: ' + err.message)
    }
  }

  const handleDownloadPowerPoint = async () => {
    if (!certificatePreview) return
    const element = document.querySelector('.certificate-container')
    if (!element) return

    try {
      const html2canvas = (await import('html2canvas')).default
      const canvas = await html2canvas(element, { scale: 2, useCORS: true })
      const dataUrl = canvas.toDataURL('image/png')

      const PptxGenJS = (await import('pptxgenjs')).default
      const prs = new PptxGenJS()
      prs.defineLayout({ name: 'A4_LANDSCAPE', width: 10, height: 7.07 })
      let slide = prs.addSlide('A4_LANDSCAPE')
      
      slide.addImage({
        data: dataUrl,
        x: 0,
        y: 0,
        w: 10,
        h: 7.07
      })

      prs.writeFile({ fileName: `${certificatePreview.fullName || 'Sertifikat'}_${certificatePreview.surname || ''}_sertifikat.pptx` })
    } catch (err) {
      console.error('PPT export error:', err)
      alert('PowerPoint eksportida xatolik yuz berdi: ' + err.message)
    }
  }

  const handleSubmit = (e) => {
    e.preventDefault()

    addDoc(collection(db, 'colculation'), {
      fullName: fullName.trim(),
      surname: surname.trim(),
      course,
      status,
      mentorName: mentorName.trim(),
      certificateDate,
      createdAt: new Date(),
    })
      .then(() => console.log('sacces'))
      .catch((error) => alert(error.message))

    e.target.reset()
    setFullName('')
    setSurname('')
    setCourse('Frontend')
    setStatus('Yakunlangan')
    setMentorName('')
    setCertificateDate('')
    setShowForm(false)
  }

  return (
    <div className={`app-shell ${darkMode ? 'dark-mode' : ''}`}>
      <nav className="topbar">
        <div className="brand">
          <div className="brand-badge-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M12 2L3 7V12C3 17.55 6.84 22.74 12 24C17.16 22.74 21 17.55 21 12V7L12 2Z" fill="url(#shieldGrad)" />
              <path d="M12 6L15 12H9L12 6Z" fill="white" opacity="0.9" />
              <path d="M9 14H15V16H9V14Z" fill="white" opacity="0.9" />
              <defs>
                <linearGradient id="shieldGrad" x1="3" y1="2" x2="21" y2="24" gradientUnits="userSpaceOnUse">
                  <stop stopColor="#2563eb" />
                  <stop offset="1" stopColor="#1d4ed8" />
                </linearGradient>
              </defs>
            </svg>
          </div>
          <div>
            <h2>TARGET IT ACADEMY</h2>
            <p>Sertifikatlar Boshqaruv Markazi</p>
          </div>
        </div>

        <div className="nav-buttons">
          <button 
            className={`tab-btn ${selectedCourse === 'Frontend' ? 'active' : ''}`}
            onClick={() => setSelectedCourse('Frontend')}
          >
            Frontend ({frontendCount})
          </button>
          <button 
            className={`tab-btn ${selectedCourse === 'Backend' ? 'active' : ''}`}
            onClick={() => setSelectedCourse('Backend')}
          >
            Backend ({backendCount})
          </button>
          <button 
            className={`tab-btn ${selectedCourse === 'English' ? 'active' : ''}`}
            onClick={() => setSelectedCourse('English')}
          >
            English ({englishCount})
          </button>
        </div>

        <div className="topbar-actions">
          <button 
            className="theme-toggle-btn"
            onClick={() => setDarkMode(!darkMode)}
            title={darkMode ? "Yorug' rejimga o'tish" : "Qorong'u rejimga o'tish"}
          >
            {darkMode ? '☀️' : '🌙'}
          </button>
          <button className="add-btn" onClick={() => setShowForm(true)}>
            + Yangi Sertifikat
          </button>
        </div>
      </nav>

      <header className="hero">
        <div className="hero-content">
          <div className="hero-text-block">
            <span className="eyebrow">Rasmiy Sertifikat Tizimi</span>
            <h1>Talabalar Sertifikatlarini Boshqarish</h1>
            <p className="hero-text">
              Target IT Academy rasmiy sertifikatlarini yaratish, tahrirlash, hamda Print, PDF, Word va PowerPoint formatlarida 1-ga-1 eksport qilish platformasi.
            </p>
          </div>

          <div className="hero-stats-row">
            <div className="stat-card">
              <span className="stat-number">{totalCount}</span>
              <span className="stat-label">Jami Sertifikatlar</span>
            </div>
            <div className="stat-card">
              <span className="stat-number">{frontendCount}</span>
              <span className="stat-label">Frontend</span>
            </div>
            <div className="stat-card">
              <span className="stat-number">{backendCount}</span>
              <span className="stat-label">Backend</span>
            </div>
            <div className="stat-card">
              <span className="stat-number">{englishCount}</span>
              <span className="stat-label">English</span>
            </div>
          </div>
        </div>

        {/* Live Search Input Bar */}
        <div className="search-bar-container">
          <span className="search-icon">🔍</span>
          <input 
            type="text" 
            className="search-input" 
            placeholder="Talaba ismi yoki familiyasi bo'yicha tezkor qidiruv..." 
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
          />
          {searchQuery && (
            <button className="clear-search-btn" onClick={() => setSearchQuery('')}>×</button>
          )}
        </div>
      </header>

      {certificatePreview && (
        <div className="modal-overlay" onClick={() => setCertificatePreview(null)}>
          <div className="certificate-modal" onClick={(e) => e.stopPropagation()}>
            <button className="close-btn" onClick={() => setCertificatePreview(null)}>×</button>
            <div className="action-buttons">
              <button className="download-btn pdf-btn" onClick={handleDownloadPDF} title="PDF sifatida yuklab olish">
                📄 PDF
              </button>
              <button className="download-btn word-btn" onClick={handleDownloadWord} title="Word sifatida yuklab olish">
                📝 Word
              </button>
              <button className="download-btn ppt-btn" onClick={handleDownloadPowerPoint} title="PowerPoint sifatida yuklab olish">
                📊 PowerPoint
              </button>
              <button className="print-btn" onClick={handlePrintDownload}>🖨️ Print</button>
            </div>
            
            <div className="certificate-container" id="certificate-node">
              {/* Background Geometry Layer */}
              <svg className="certificate-bg-svg" viewBox="0 0 1000 707" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">
                {/* Canvas Outer Frame & Borders */}
                <rect width="1000" height="707" fill="#fbfbfa" />
                <rect x="14" y="14" width="972" height="679" fill="none" stroke="#0e0f12" strokeWidth="3" />
                <rect x="18" y="18" width="964" height="671" fill="none" stroke="#0e0f12" strokeWidth="1" />

                {/* Bottom Light Tan/Gold Strip */}
                <rect x="14" y="662" width="972" height="31" fill="#eadeca" />

                {/* Right Geometric Shapes */}
                <g className="cert-right-art">
                  {/* Top-Right Black Layer */}
                  <polygon points="530,662 1000,0 1000,662" fill="#0c0d10" />
                  
                  {/* Inset Gold Line in Black Region */}
                  <polyline points="750,22 978,22 978,480" fill="none" stroke="#d4af37" strokeWidth="2.5" />

                  {/* Golden Ribbon Accent behind Red */}
                  <polygon points="515,662 715,0 745,0 1000,400 1000,440 515,662" fill="#e5b839" />

                  {/* Main Red Triangle / Diagonal Fold */}
                  <polygon points="530,662 730,0 1000,0 1000,270 530,662" fill="#d31a1a" />
                  
                  {/* Dark Red Shadow / Fold Layer */}
                  <polygon points="730,0 1000,0 1000,270 830,175" fill="#b01313" />
                  
                  {/* Gold Highlight Line across Red Fold */}
                  <polyline points="730,0 1000,270" fill="none" stroke="#fce28b" strokeWidth="3" />
                </g>
              </svg>

              {/* Main Content Area */}
              <div className="cert-content">
                {/* Header Logos */}
                <div className="cert-header-row">
                  <div className="logo-box target-logo">
                    <svg width="42" height="42" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <circle cx="20" cy="20" r="17" stroke="#dc2626" strokeWidth="3.5" fill="none" />
                      <circle cx="20" cy="20" r="10" stroke="#000" strokeWidth="1.5" fill="none" />
                      <circle cx="20" cy="20" r="5" fill="#dc2626" />
                      <line x1="20" y1="0" x2="20" y2="40" stroke="#dc2626" strokeWidth="2" />
                      <line x1="0" y1="20" x2="40" y2="20" stroke="#dc2626" strokeWidth="2" />
                    </svg>
                    <div className="logo-text">
                      <span className="target-title">TARGET</span>
                      <span className="target-sub">IT SCHOOL</span>
                    </div>
                  </div>

                  <div className="logo-box itpark-logo">
                    <svg width="42" height="42" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <circle cx="20" cy="20" r="19" fill="#16a34a" />
                      <circle cx="20" cy="14" r="3.5" fill="white" />
                      <circle cx="13" cy="25" r="3.5" fill="white" />
                      <circle cx="27" cy="25" r="3.5" fill="white" />
                      <line x1="20" y1="14" x2="13" y2="25" stroke="white" strokeWidth="2" />
                      <line x1="20" y1="14" x2="27" y2="25" stroke="white" strokeWidth="2" />
                      <line x1="13" y1="25" x2="27" y2="25" stroke="white" strokeWidth="2" />
                    </svg>
                    <div className="logo-text">
                      <span className="park-title">IT PARK</span>
                      <span className="park-sub">START local & GO global</span>
                    </div>
                  </div>
                </div>

                {/* Organization Header Text */}
                <div className="cert-org-header">
                  <p>O'ZBEKISTON RESPUBLIKASI TOSHKENT SHAHRI</p>
                  <p className="org-bold">"TARGET IT ACADEMY" MA'SULIYATI CHEKLANGAN JAMIYAT TOMONIDAN</p>
                </div>

                {/* Student Name Section */}
                <div className="cert-name-section">
                  <div className="name-line-container">
                    <span className="name-value">
                      {certificatePreview.fullName || 'Ali'} {certificatePreview.surname || 'Karimov'}
                    </span>
                    <span className="ga-label">ga</span>
                  </div>
                </div>

                {/* Course Details */}
                <div className="cert-course-details">
                  <h3 className="course-title-quotes">
                    "WEB DASTURLASH - {certificatePreview.course ? certificatePreview.course.toUpperCase() : 'FRONTEND VA BACKEND'}"
                  </h3>
                  <p className="course-subtitle">
                    YO'NALISHINI MUVAFFAQIYATLI TAMOMLAGANLIGI UCHUN
                  </p>
                </div>

                {/* Sertifikat Title */}
                <div className="cert-main-title-box">
                  <h1 className="cert-big-title">SERTIFIKAT</h1>
                  <p className="cert-taqdim-text">TAQDIM ETILDI</p>
                </div>

                {/* Footer Metadata */}
                <div className="cert-footer-row">
                  <div className="footer-item date-item">
                    <div className="sig-line">
                      <span className="item-value">
                        {certificatePreview.certificateDate 
                          ? new Date(certificatePreview.certificateDate).toLocaleDateString('uz-UZ') 
                          : '12.05.2024'}
                      </span>
                    </div>
                    <span className="item-label">Sana</span>
                  </div>

                  <div className="footer-item mentor-item">
                    <span className="item-label-top">Mentor:</span>
                    <div className="sig-line">
                      <span className="item-value">
                        {certificatePreview.mentorName || "No'monjonov Shoxruhmirzo"}
                      </span>
                    </div>
                  </div>

                  <div className="footer-item director-item">
                    <span className="item-label-top">Direktor:</span>
                    <div className="sig-line signature-wrapper">
                      <span className="item-value">Sayfuddinov Abdulloh</span>
                      {/* Authentic Blue Ink Signature Overlay */}
                      <svg className="director-signature-svg" viewBox="0 0 120 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M5 25 C 20 5, 35 35, 45 15 C 55 -5, 60 30, 75 10 C 85 2, 90 28, 105 18 C 112 12, 115 22, 118 20" stroke="#1d4ed8" strokeWidth="2.5" strokeLinecap="round" fill="none" />
                        <path d="M25 28 C 45 12, 70 32, 95 16" stroke="#1d4ed8" strokeWidth="1.8" strokeLinecap="round" fill="none" />
                      </svg>
                    </div>
                  </div>
                </div>

                {/* Bottom Left Codes */}
                <div className="cert-codes-row">
                  <span>NAS.UZ001</span>
                  <span>MT.0324-05</span>
                </div>
              </div>

              {/* Gold Official Seal / Stamp */}
              <div className="cert-official-seal">
                <svg className="seal-svg" viewBox="0 0 160 160" xmlns="http://www.w3.org/2000/svg">
                  <defs>
                    <radialGradient id="goldGrad" cx="50%" cy="50%" r="50%">
                      <stop offset="0%" stopColor="#fef08a" />
                      <stop offset="50%" stopColor="#eab308" />
                      <stop offset="100%" stopColor="#ca8a04" />
                    </radialGradient>
                    <linearGradient id="blueGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                      <stop offset="0%" stopColor="#1e3a8a" />
                      <stop offset="100%" stopColor="#0f172a" />
                    </linearGradient>
                    <path id="textCirclePath" d="M 80, 80 m -62, 0 a 62,62 0 1,1 124,0 a 62,62 0 1,1 -124,0" />
                  </defs>

                  {/* Scalloped / Starburst Outer Ring */}
                  <g fill="url(#goldGrad)">
                    {Array.from({ length: 36 }).map((_, i) => (
                      <circle
                        key={i}
                        cx={80 + 72 * Math.cos((i * 10 * Math.PI) / 180)}
                        cy={80 + 72 * Math.sin((i * 10 * Math.PI) / 180)}
                        r="7"
                      />
                    ))}
                    <circle cx="80" cy="80" r="70" />
                  </g>

                  {/* Blue Ring */}
                  <circle cx="80" cy="80" r="62" fill="url(#blueGrad)" stroke="#fef08a" strokeWidth="2" />

                  {/* Circular Text */}
                  <text fill="#ffffff" fontSize="7.2" fontWeight="bold" letterSpacing="0.8">
                    <textPath href="#textCirclePath" startOffset="0%">
                      O'ZBEKISTON RESPUBLIKASI TOSHKENT SHAHARI MAS'ULIYATI CHEKLANGAN JAMIYATI *
                    </textPath>
                  </text>

                  {/* Inner Box & Content */}
                  <circle cx="80" cy="80" r="44" fill="url(#blueGrad)" stroke="#eab308" strokeWidth="1.5" />
                  <rect x="46" y="52" width="68" height="56" fill="none" stroke="#fef08a" strokeWidth="1.5" rx="4" />
                  
                  <text x="80" y="68" textAnchor="middle" fill="#ffffff" fontSize="8" fontWeight="900" letterSpacing="0.5">
                    TARGET IT
                  </text>
                  <text x="80" y="78" textAnchor="middle" fill="#ffffff" fontSize="7.5" fontWeight="900" letterSpacing="0.5">
                    ACADEMY
                  </text>
                  
                  <line x1="52" y1="84" x2="108" y2="84" stroke="#fef08a" strokeWidth="1" />
                  
                  <text x="80" y="93" textAnchor="middle" fill="#fef08a" fontSize="6.5" fontWeight="bold">
                    XUJJATLAR
                  </text>
                  <text x="80" y="101" textAnchor="middle" fill="#fef08a" fontSize="6.5" fontWeight="bold">
                    UCHUN
                  </text>
                </svg>
              </div>
            </div>
          </div>
        </div>
      )}

      {editId && (
        <div className="modal-overlay" onClick={cancelEdit}>
          <form className="modal-card" onSubmit={handleEditSubmit} onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <div>
                <p className="eyebrow">Tahrirlash</p>
                <h2>Sertifikat ma'lumotlari</h2>
              </div>
              <button type="button" className="close-btn" onClick={cancelEdit}>
                ×
              </button>
            </div>

            <label className="field">
              <span>Ism</span>
              <input
                type="text"
                value={editFullName}
                onChange={(e) => setEditFullName(e.target.value)}
              />
            </label>

            <label className="field">
              <span>Familiya</span>
              <input
                type="text"
                value={editSurname}
                onChange={(e) => setEditSurname(e.target.value)}
              />
            </label>

            <label className="field">
              <span>Kurs</span>
              <select value={editCourse} onChange={(e) => setEditCourse(e.target.value)}>
                <option value="Frontend">Frontend</option>
                <option value="Backend">Backend</option>
                <option value="English">English</option>
              </select>
            </label>

            <label className="field">
              <span>Holat</span>
              <select value={editStatus} onChange={(e) => setEditStatus(e.target.value)}>
                <option value="Yakunlangan">Yakunlangan</option>
                <option value="Jarayonda">Jarayonda</option>
                <option value="Sertifikat olgan">Sertifikat olgan</option>
              </select>
            </label>

            <label className="field">
              <span>Usozning ismi</span>
              <input
                type="text"
                value={editMentorName}
                onChange={(e) => setEditMentorName(e.target.value)}
                placeholder="Masalan: No'monjonov Shoxruh mirzo"
              />
            </label>

            <label className="field">
              <span>Sertifikat sanasi</span>
              <input
                type="date"
                value={editCertificateDate}
                onChange={(e) => setEditCertificateDate(e.target.value)}
              />
            </label>

            <div className="modal-actions">
              <button type="submit" className="submit-btn">Saqlash</button>
              <button type="button" className="cancel-btn" onClick={cancelEdit}>
                Bekor qilish
              </button>
            </div>
          </form>
        </div>
      )}

      {showForm && (
        <div className="modal-overlay" onClick={() => setShowForm(false)}>
          <form className="modal-card" onSubmit={handleSubmit} onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <div>
                <p className="eyebrow">Yangi sertifikat</p>
                <h2>Add sertifikat</h2>
              </div>
              <button type="button" className="close-btn" onClick={() => setShowForm(false)}>
                ×
              </button>
            </div>

            <label className="field">
              <span>Ism</span>
              <input
                type="text"
                placeholder="Masalan: Ali"
                value={fullName}
                onChange={(e) => setFullName(e.target.value)}
              />
            </label>

            <label className="field">
              <span>Familiya</span>
              <input
                type="text"
                placeholder="Masalan: Karimov"
                value={surname}
                onChange={(e) => setSurname(e.target.value)}
              />
            </label>

            <label className="field">
              <span>Kurs</span>
              <select value={course} onChange={(e) => setCourse(e.target.value)}>
                <option value="Frontend">Frontend</option>
                <option value="Backend">Backend</option>
                <option value="English">English</option>
              </select>
            </label>

            <label className="field">
              <span>Holat</span>
              <select value={status} onChange={(e) => setStatus(e.target.value)}>
                <option value="Yakunlangan">Yakunlangan</option>
                <option value="Jarayonda">Jarayonda</option>
                <option value="Sertifikat olgan">Sertifikat olgan</option>
              </select>
            </label>

            <label className="field">
              <span>Usozning ismi</span>
              <input
                type="text"
                placeholder="Masalan: No'monjonov Shoxruh mirzo"
                value={mentorName}
                onChange={(e) => setMentorName(e.target.value)}
              />
            </label>

            <label className="field">
              <span>Sertifikat sanasi</span>
              <input
                type="date"
                value={certificateDate}
                onChange={(e) => setCertificateDate(e.target.value)}
              />
            </label>

            <button className="submit-btn">Saqlash</button>
          </form>
        </div>
      )}

      <section className="certificate-list">
        {(() => {
          const filtered = calculation.filter(item => {
            const itemCourse = item.course || item.text || 'Frontend'
            const isCourseMatch = itemCourse === selectedCourse
            
            if (!searchQuery.trim()) return isCourseMatch
            
            const fullNameStr = `${item.fullName || item.title || ''} ${item.surname || ''}`.toLowerCase()
            return isCourseMatch && fullNameStr.includes(searchQuery.toLowerCase().trim())
          })

          if (filtered.length === 0) {
            return (
              <div className="empty-state">
                <div className="empty-icon">📜</div>
                <h3>{searchQuery ? 'Qidiruv bo‘yicha sertifikat topilmadi' : 'Hozircha ushbu bo‘limda sertifikat yo‘q'}</h3>
                <p>{searchQuery ? 'Boshqa ism yoki familiya kiritib ko‘ring.' : 'Yuqoridagi "+ Yangi Sertifikat" tugmasini bosib birinchi sertifikatni saqlang.'}</p>
              </div>
            )
          }

          return filtered.map((item) => {
            const displayName = `${item.fullName || item.title || 'Sertifikat'} ${item.surname || ''}`.trim()
            const courseName = item.course || item.text || 'Frontend'
            const statusName = item.status || 'Yakunlangan'
            const initials = (item.fullName || item.title || 'S').charAt(0).toUpperCase()
            const formattedDate = item.certificateDate 
              ? new Date(item.certificateDate).toLocaleDateString('uz-UZ') 
              : '12.05.2024'

            return (
              <article className="certificate-card" key={item.id}>
                <div className="card-top-bar">
                  <span className="serial-badge">NAS.UZ001</span>
                  <div className="certificate-badge">
                    <span className="status-dot"></span>
                    {statusName}
                  </div>
                </div>

                <div className="card-main-info">
                  <div className="avatar-circle">{initials}</div>
                  <div className="user-details">
                    <h3>{displayName}</h3>
                    <span className="course-pill">{courseName} Yo'nalishi</span>
                  </div>
                </div>

                <div className="card-meta-info">
                  <div className="meta-item">
                    <span className="meta-label">Mentor:</span>
                    <span className="meta-val">{item.mentorName || "No'monjonov Shoxruhmirzo"}</span>
                  </div>
                  <div className="meta-item">
                    <span className="meta-label">Sana:</span>
                    <span className="meta-val">{formattedDate}</span>
                  </div>
                </div>

                <div className="card-actions-grid">
                  <button className="card-action-btn primary-action" onClick={() => handlePrintCertificate(item)} title="Sertifikatni Ko'rish va Chop etish">
                    👁️ Ko'rish & Export
                  </button>
                  <button className="card-action-btn edit-action" onClick={() => editDocument(item.id)} title="Tahrirlash">
                    ✏️ Edit
                  </button>
                  <button className="card-action-btn danger-action" onClick={() => deleteDocument(item.id)} title="O'chirish">
                    🗑️
                  </button>
                </div>
              </article>
            )
          })
        })()}
      </section>
    </div>
  )
}

export default App
