import { useState } from "react"
import { motion } from "framer-motion"
import { Document, Packer, Paragraph, TextRun, AlignmentType } from "docx"
import { saveAs } from "file-saver"
import { GlassBackground } from "@/components/GlassBackground"
import { GlassField } from "@/components/GlassField"
import Icon from "@/components/ui/icon"

interface DocData {
  fullName: string
  birthDate: string
  registration: string
  passport: string
  docDate: string
  body: string
}

const initialData: DocData = {
  fullName: "Иванов Иван Иванович",
  birthDate: "01.01.1990",
  registration: "г. Москва, ул. Примерная, д. 1, кв. 1",
  passport: "1234 567890",
  docDate: "24.06.2026",
  body:
    "Настоящим подтверждаю достоверность указанных выше сведений. Прошу рассмотреть данное заявление и принять соответствующее решение в установленном порядке.",
}

function DocGeneratorPage() {
  const [data, setData] = useState<DocData>(initialData)

  const update = (field: keyof DocData) => (value: string) =>
    setData((prev) => ({ ...prev, [field]: value }))

  const handleDownload = async () => {
    const line = (label: string, value: string) =>
      new Paragraph({
        spacing: { after: 160 },
        children: [
          new TextRun({ text: `${label}: `, bold: true, size: 24 }),
          new TextRun({ text: value, size: 24 }),
        ],
      })

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 320 },
              children: [new TextRun({ text: "ЗАЯВЛЕНИЕ", bold: true, size: 32 })],
            }),
            line("ФИО", data.fullName),
            line("Дата рождения", data.birthDate),
            line("Адрес прописки", data.registration),
            line("Паспорт", data.passport),
            new Paragraph({
              spacing: { before: 240, after: 240 },
              children: [new TextRun({ text: data.body, size: 24 })],
            }),
            new Paragraph({
              spacing: { before: 480 },
              children: [
                new TextRun({ text: `Дата: ${data.docDate}`, size: 24 }),
                new TextRun({ text: "          Подпись: ______________", size: 24 }),
              ],
            }),
          ],
        },
      ],
    })

    const blob = await Packer.toBlob(doc)
    saveAs(blob, `Документ_${data.fullName.replace(/\s+/g, "_")}.docx`)
  }

  return (
    <main className="relative min-h-screen overflow-hidden px-4 py-8 md:px-8 md:py-10">
      <GlassBackground />

      <motion.div
        initial={{ opacity: 0, y: 16 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
        className="relative z-10 mx-auto max-w-6xl"
      >
        <div className="mb-8 flex items-center gap-3">
          <div
            className="flex h-12 w-12 items-center justify-center rounded-2xl text-purple-700"
            style={{
              background: "rgba(255,255,255,0.7)",
              border: "1px solid rgba(255,255,255,0.6)",
              boxShadow: "inset 0 1px 2px rgba(255,255,255,1), 0 4px 12px rgba(0,0,0,0.06)",
            }}
          >
            <Icon name="FileText" size={24} />
          </div>
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-gray-800">Генератор документов</h1>
            <p className="text-[13px] text-gray-500">Заполните поля слева — документ соберётся автоматически</p>
          </div>
        </div>

        <div className="grid gap-6 lg:grid-cols-[400px_1fr]">
          {/* Форма */}
          <div
            className="space-y-4 rounded-3xl p-5"
            style={{
              background: "rgba(255,255,255,0.45)",
              backdropFilter: "blur(40px) saturate(180%)",
              WebkitBackdropFilter: "blur(40px) saturate(180%)",
              border: "1px solid rgba(255,255,255,0.5)",
              boxShadow: "inset 0 1px 1px rgba(255,255,255,0.9), 0 16px 40px rgba(0,0,0,0.08)",
            }}
          >
            <GlassField label="ФИО" icon="User" value={data.fullName} onChange={update("fullName")} placeholder="Фамилия Имя Отчество" />
            <GlassField label="Дата рождения" icon="Calendar" value={data.birthDate} onChange={update("birthDate")} placeholder="ДД.ММ.ГГГГ" />
            <GlassField label="Адрес прописки" icon="MapPin" value={data.registration} onChange={update("registration")} placeholder="Город, улица, дом, квартира" />
            <GlassField label="Паспорт" icon="CreditCard" value={data.passport} onChange={update("passport")} placeholder="Серия и номер" />
            <GlassField label="Дата документа" icon="CalendarClock" value={data.docDate} onChange={update("docDate")} placeholder="ДД.ММ.ГГГГ" />
            <GlassField label="Текст документа" icon="AlignLeft" value={data.body} onChange={update("body")} multiline placeholder="Основной текст..." />

            <motion.button
              onClick={handleDownload}
              whileHover={{ scale: 1.02, y: -2 }}
              whileTap={{ scale: 0.98 }}
              className="flex w-full items-center justify-center gap-2 rounded-2xl py-3.5 text-[15px] font-semibold text-white"
              style={{
                background: "linear-gradient(135deg, #9333ea, #ec4899)",
                boxShadow: "0 8px 24px rgba(147,51,234,0.35)",
              }}
            >
              <Icon name="Download" size={18} />
              Скачать Word-документ
            </motion.button>
          </div>

          {/* Предпросмотр */}
          <div
            className="rounded-3xl p-2"
            style={{
              background: "rgba(255,255,255,0.45)",
              backdropFilter: "blur(40px) saturate(180%)",
              WebkitBackdropFilter: "blur(40px) saturate(180%)",
              border: "1px solid rgba(255,255,255,0.5)",
              boxShadow: "inset 0 1px 1px rgba(255,255,255,0.9), 0 16px 40px rgba(0,0,0,0.08)",
            }}
          >
            <div className="rounded-[20px] bg-white px-8 py-12 shadow-sm md:px-14 md:py-16 min-h-[600px]">
              <h2 className="mb-10 text-center text-xl font-bold uppercase tracking-wide text-gray-900">Заявление</h2>

              <div className="space-y-4 text-[15px] leading-relaxed text-gray-800">
                <p><span className="font-semibold">ФИО:</span> {data.fullName || "—"}</p>
                <p><span className="font-semibold">Дата рождения:</span> {data.birthDate || "—"}</p>
                <p><span className="font-semibold">Адрес прописки:</span> {data.registration || "—"}</p>
                <p><span className="font-semibold">Паспорт:</span> {data.passport || "—"}</p>

                <p className="pt-4 whitespace-pre-wrap">{data.body || "—"}</p>

                <div className="flex items-center justify-between pt-12">
                  <span>Дата: {data.docDate || "—"}</span>
                  <span>Подпись: ______________</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </motion.div>
    </main>
  )
}

export default DocGeneratorPage
