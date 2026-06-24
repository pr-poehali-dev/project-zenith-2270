import { useState } from "react"
import { motion } from "framer-motion"
import { GlassBackground } from "@/components/GlassBackground"
import { GlassField } from "@/components/GlassField"
import Icon from "@/components/ui/icon"
import {
  CaseData, defaultData,
  buildOpis, buildPostanovlenieProdlenie, buildSpravka,
  buildPostanovlenieOtkaz, buildUvedomlenieZayavitel,
  buildUvedomlenieProcuror, buildZaprosMincifra,
  downloadDoc,
} from "@/lib/docTemplates"

const SECTIONS = [
  { id: "case", label: "Дело / КУСП" },
  { id: "applicant", label: "Заявитель" },
  { id: "executor", label: "Исполнитель" },
  { id: "chiefs", label: "Руководство" },
  { id: "crime", label: "Факты дела" },
  { id: "mincifra", label: "МинЦифр" },
]

const DOCS = [
  { id: "opis", label: "Опись", icon: "List", build: buildOpis },
  { id: "prodlenie", label: "Постановление о продлении", icon: "Clock", build: buildPostanovlenieProdlenie },
  { id: "spravka", label: "Справка", icon: "FileCheck", build: buildSpravka },
  { id: "otkaz", label: "Постановление об отказе", icon: "FileX", build: buildPostanovlenieOtkaz },
  { id: "uved_zayavitel", label: "Уведомление заявителю", icon: "Mail", build: buildUvedomlenieZayavitel },
  { id: "uved_procuror", label: "Уведомление прокурору", icon: "Briefcase", build: buildUvedomlenieProcuror },
  { id: "mincifra", label: "Запрос в МинЦифр", icon: "Send", build: buildZaprosMincifra },
]

const glassStyle = {
  background: "rgba(255,255,255,0.45)",
  backdropFilter: "blur(40px) saturate(180%)",
  WebkitBackdropFilter: "blur(40px) saturate(180%)",
  border: "1px solid rgba(255,255,255,0.5)",
  boxShadow: "inset 0 1px 1px rgba(255,255,255,0.9), 0 16px 40px rgba(0,0,0,0.08)",
}

export default function DocGeneratorPage() {
  const [data, setData] = useState<CaseData>(defaultData)
  const [activeSection, setActiveSection] = useState("case")
  const [loading, setLoading] = useState<string | null>(null)

  const upd = (field: keyof CaseData) => (value: string) =>
    setData(prev => ({ ...prev, [field]: value }))

  const handleDownload = async (doc: typeof DOCS[0]) => {
    setLoading(doc.id)
    await downloadDoc(doc.build(data), `${doc.label}_КУСП${data.kuspNumber}.docx`)
    setLoading(null)
  }

  const handleDownloadAll = async () => {
    setLoading("all")
    for (const doc of DOCS) {
      await downloadDoc(doc.build(data), `${doc.label}_КУСП${data.kuspNumber}.docx`)
    }
    setLoading(null)
  }

  return (
    <main className="relative min-h-screen overflow-auto px-4 py-6 md:px-8 md:py-8">
      <GlassBackground />

      <div className="relative z-10 mx-auto max-w-7xl">
        <motion.div
          initial={{ opacity: 0, y: -12 }}
          animate={{ opacity: 1, y: 0 }}
          className="mb-6 flex items-center gap-3"
        >
          <div
            className="flex h-12 w-12 shrink-0 items-center justify-center rounded-2xl"
            style={{ background: "rgba(255,255,255,0.7)", border: "1px solid rgba(255,255,255,0.6)", boxShadow: "inset 0 1px 2px rgba(255,255,255,1), 0 4px 12px rgba(0,0,0,0.06)" }}
          >
            <Icon name="Shield" size={22} className="text-blue-700" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-gray-800">МВД — Генератор документов</h1>
            <p className="text-[13px] text-gray-500">Заполните поля один раз — все документы подставят данные автоматически</p>
          </div>
        </motion.div>

        <div className="grid gap-5 lg:grid-cols-[420px_1fr]">
          {/* ── ЛЕВАЯ ПАНЕЛЬ: ФОРМА ── */}
          <motion.div
            initial={{ opacity: 0, x: -16 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ delay: 0.1 }}
            className="flex flex-col gap-4"
          >
            <div className="flex flex-wrap gap-1.5 rounded-2xl p-2" style={glassStyle}>
              {SECTIONS.map(s => (
                <button
                  key={s.id}
                  onClick={() => setActiveSection(s.id)}
                  className="rounded-xl px-3 py-1.5 text-[13px] font-medium transition-all"
                  style={activeSection === s.id
                    ? { background: "linear-gradient(135deg, #3b82f6, #6366f1)", color: "#fff", boxShadow: "0 4px 12px rgba(99,102,241,0.35)" }
                    : { background: "rgba(255,255,255,0.5)", color: "#4b5563" }
                  }
                >
                  {s.label}
                </button>
              ))}
            </div>

            <div className="space-y-3.5 rounded-3xl p-5" style={glassStyle}>
              {activeSection === "case" && <>
                <GlassField label="Номер КУСП" icon="Hash" value={data.kuspNumber} onChange={upd("kuspNumber")} placeholder="9152" />
                <GlassField label="Дата регистрации КУСП" icon="Calendar" value={data.kuspDate} onChange={upd("kuspDate")} placeholder="14.04.2026" />
                <GlassField label="Дата начала" icon="CalendarClock" value={data.dateStart} onChange={upd("dateStart")} placeholder="14.04.2026" />
                <GlassField label="Дата окончания" icon="CalendarCheck" value={data.dateEnd} onChange={upd("dateEnd")} placeholder="23.04.2026" />
                <GlassField label="Год" icon="BookOpen" value={data.year} onChange={upd("year")} placeholder="2026" />
                <GlassField label="Дата постановления о продлении" icon="Calendar" value={data.datePostanovlenie} onChange={upd("datePostanovlenie")} placeholder="16.04.2026" />
                <GlassField label="Дата справки" icon="Calendar" value={data.dateSpravka} onChange={upd("dateSpravka")} placeholder="20.04.2026" />
                <GlassField label="Дата постановления об отказе" icon="Calendar" value={data.dateOtkaz} onChange={upd("dateOtkaz")} placeholder="23.04.2026" />
              </>}

              {activeSection === "applicant" && <>
                <GlassField label="ФИО заявителя (кратко, род. пад.)" icon="User" value={data.applicantFio} onChange={upd("applicantFio")} placeholder="Максаковой Д.В." />
                <GlassField label="ФИО заявителя (полностью)" icon="User" value={data.applicantFioFull} onChange={upd("applicantFioFull")} placeholder="Максакова Дарья Владимировна" />
                <GlassField label="Дата рождения заявителя" icon="Cake" value={data.applicantBirthDate} onChange={upd("applicantBirthDate")} placeholder="19.06.2002" />
                <GlassField label="Адрес заявителя" icon="MapPin" value={data.applicantAddress} onChange={upd("applicantAddress")} placeholder="г. Воронеж, ул. ..." />
                <GlassField label="Обращение в письме" icon="MessageSquare" value={data.applicantGreeting} onChange={upd("applicantGreeting")} placeholder="Уважаемая Дарья Владимировна!" />
              </>}

              {activeSection === "executor" && <>
                <GlassField label="ФИО исполнителя" icon="UserCheck" value={data.executorFio} onChange={upd("executorFio")} placeholder="А.А. Горшков" />
                <GlassField label="Звание исполнителя" icon="Award" value={data.executorRank} onChange={upd("executorRank")} placeholder="лейтенант полиции" />
                <GlassField label="Телефон исполнителя" icon="Phone" value={data.executorPhone} onChange={upd("executorPhone")} placeholder="(960) 108-90-69" />
                <GlassField label="Email исполнителя" icon="AtSign" value={data.executorEmail} onChange={upd("executorEmail")} placeholder="agorshkov109@mvd.ru" />
              </>}

              {activeSection === "chiefs" && <>
                <GlassField label="ФИО начальника" icon="ShieldCheck" value={data.chiefFio} onChange={upd("chiefFio")} placeholder="В.А. Бортников" />
                <GlassField label="Звание начальника" icon="Award" value={data.chiefRank} onChange={upd("chiefRank")} placeholder="подполковник полиции" />
                <GlassField label="ФИО зам. начальника ОУР" icon="Users" value={data.deputyFio} onChange={upd("deputyFio")} placeholder="Н.Е. Родин" />
                <GlassField label="ФИО прокурора (дат. пад.)" icon="Scale" value={data.procFioShort} onChange={upd("procFioShort")} placeholder="А.А. Дмитриеву" />
                <GlassField label="Имя-отчество прокурора" icon="Scale" value={data.procFioFull} onChange={upd("procFioFull")} placeholder="Алексей Александрович" />
                <GlassField label="Звание прокурора" icon="Award" value={data.procRank} onChange={upd("procRank")} placeholder="старшему советнику юстиции" />
              </>}

              {activeSection === "crime" && <>
                <GlassField label="Дата события" icon="AlertCircle" value={data.crimeDate} onChange={upd("crimeDate")} placeholder="06.04.2026" />
                <GlassField label="Время события" icon="Clock" value={data.crimeTime} onChange={upd("crimeTime")} placeholder="14 часов 00 минут" />
                <GlassField label="Описание события (для постановления об отказе)" icon="AlignLeft" value={data.crimeDescription} onChange={upd("crimeDescription")} multiline placeholder="Позвонил незнакомый мужчина..." />
              </>}

              {activeSection === "mincifra" && <>
                <GlassField label="ФИО фигуранта запроса" icon="User" value={data.mincifraPersonFio} onChange={upd("mincifraPersonFio")} placeholder="Карпунина Елена Михайловна" />
                <GlassField label="Дата рождения фигуранта" icon="Cake" value={data.mincifraPersonBirth} onChange={upd("mincifraPersonBirth")} placeholder="21.08.1976" />
                <GlassField label="Период с" icon="Calendar" value={data.mincifraDateFrom} onChange={upd("mincifraDateFrom")} placeholder="01.11.2025" />
                <GlassField label="Период по" icon="Calendar" value={data.mincifraDateTo} onChange={upd("mincifraDateTo")} placeholder="30.11.2025" />
                <GlassField label="ФИО директора МинЦифр (дат. пад.)" icon="User" value={data.mincifraDirectorFio} onChange={upd("mincifraDirectorFio")} placeholder="Цветкову С.В." />
                <GlassField label="Имя-отчество директора МинЦифр" icon="User" value={data.mincifraDirectorFullName} onChange={upd("mincifraDirectorFullName")} placeholder="Сергей Валерьевич" />
              </>}
            </div>
          </motion.div>

          {/* ── ПРАВАЯ ПАНЕЛЬ: ДОКУМЕНТЫ ── */}
          <motion.div
            initial={{ opacity: 0, x: 16 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ delay: 0.15 }}
            className="flex flex-col gap-4"
          >
            <motion.button
              onClick={handleDownloadAll}
              disabled={loading === "all"}
              whileHover={{ scale: 1.01, y: -2 }}
              whileTap={{ scale: 0.99 }}
              className="flex items-center justify-center gap-2.5 rounded-2xl py-4 text-[15px] font-semibold text-white disabled:opacity-70"
              style={{ background: "linear-gradient(135deg, #2563eb, #7c3aed)", boxShadow: "0 8px 24px rgba(37,99,235,0.35)" }}
            >
              <Icon name={loading === "all" ? "Loader" : "PackageOpen"} size={20} />
              {loading === "all" ? "Формирую все документы..." : "Скачать весь пакет (7 документов)"}
            </motion.button>

            <div className="rounded-3xl p-4 space-y-2.5" style={glassStyle}>
              <p className="text-[13px] font-medium text-gray-500 px-1 pb-1">Или скачайте по одному:</p>
              {DOCS.map((doc, i) => (
                <motion.div
                  key={doc.id}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  transition={{ delay: 0.1 + i * 0.05 }}
                  className="flex items-center gap-3 rounded-2xl px-4 py-3.5"
                  style={{
                    background: "rgba(255,255,255,0.6)",
                    border: "1px solid rgba(255,255,255,0.6)",
                    boxShadow: "inset 0 1px 1px rgba(255,255,255,0.9), 0 2px 8px rgba(0,0,0,0.04)",
                  }}
                >
                  <div
                    className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl"
                    style={{ background: "rgba(255,255,255,0.8)", border: "1px solid rgba(255,255,255,0.6)" }}
                  >
                    <Icon name={doc.icon} size={18} className="text-blue-600" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-[14px] font-semibold text-gray-800">{doc.label}</p>
                    <p className="text-[12px] text-gray-400">КУСП № {data.kuspNumber} от {data.kuspDate}</p>
                  </div>
                  <motion.button
                    onClick={() => handleDownload(doc)}
                    disabled={loading === doc.id}
                    whileHover={{ scale: 1.05 }}
                    whileTap={{ scale: 0.95 }}
                    className="flex items-center gap-1.5 rounded-xl px-3 py-2 text-[13px] font-medium text-white disabled:opacity-60"
                    style={{ background: "linear-gradient(135deg, #3b82f6, #6366f1)", boxShadow: "0 4px 12px rgba(99,102,241,0.3)" }}
                  >
                    <Icon name={loading === doc.id ? "Loader" : "Download"} size={14} />
                    {loading === doc.id ? "..." : ".docx"}
                  </motion.button>
                </motion.div>
              ))}
            </div>

            <div
              className="rounded-2xl px-4 py-3 flex items-start gap-3"
              style={{ background: "rgba(239,246,255,0.7)", border: "1px solid rgba(147,197,253,0.5)" }}
            >
              <Icon name="Info" size={16} className="text-blue-500 mt-0.5 shrink-0" />
              <p className="text-[13px] text-blue-700 leading-relaxed">
                Данные подставляются во все документы автоматически. Измените любое поле — и при скачивании уже будет обновлённый текст.
              </p>
            </div>
          </motion.div>
        </div>
      </div>
    </main>
  )
}
