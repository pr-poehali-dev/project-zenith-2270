import {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType,
  convertInchesToTwip, UnderlineType,
} from "docx"
import { saveAs } from "file-saver"

export interface CaseData {
  kuspNumber: string       // 9152
  kuspDate: string         // 14.04.2026
  dateStart: string        // 14.04.2026
  dateEnd: string          // 23.04.2026
  year: string             // 2026

  // Заявитель
  applicantFio: string            // Максакова Д.В.
  applicantFioFull: string        // Максакова Дарья Владимировна
  applicantBirthDate: string      // 19.06.2002
  applicantAddress: string        // Ул. Генерала Лизюкова, д. 28, кв. 103, г. Воронеж
  applicantGreeting: string       // Уважаемая Елена Михайловна!

  // Исполнитель
  executorFio: string             // А.А. Горшков
  executorRank: string            // лейтенант полиции
  executorPhone: string           // (960) 108-90-69
  executorEmail: string           // agorshkov109@mvd.ru

  // Начальник
  chiefFio: string                // В.А. Бортников
  chiefRank: string               // подполковник полиции
  deputyFio: string               // Н.Е. Родин

  // Прокурор
  procFioShort: string            // А.А. Дмитриеву
  procFioFull: string             // Алексей Александрович
  procRank: string                // старшему советнику юстиции

  // Даты документов
  datePostanovlenie: string       // 16.04.2026
  dateSpravka: string             // 20.04.2026
  dateOtkaz: string               // 23.04.2026

  // Факты дела
  crimeDate: string               // 06.04.2026
  crimeTime: string               // 14 часов 00 минут
  crimeDescription: string        // описание события

  // Запрос МинЦифр
  mincifraPersonFio: string       // Карпунина Елена Михайловна
  mincifraPersonBirth: string     // 21.08.1976
  mincifraDateFrom: string        // 01.11.2025
  mincifraDateTo: string          // 30.11.2025
  mincifraDirectorFio: string     // Цветкову С.В.
  mincifraDirectorFullName: string // Сергей Валерьевич
}

export const defaultData: CaseData = {
  kuspNumber: "9152",
  kuspDate: "14.04.2026",
  dateStart: "14.04.2026",
  dateEnd: "23.04.2026",
  year: "2026",

  applicantFio: "Максаковой Д.В.",
  applicantFioFull: "Максакова Дарья Владимировна",
  applicantBirthDate: "19.06.2002",
  applicantAddress: "Ул. Генерала Лизюкова, д. 28, кв. 103, г. Воронеж",
  applicantGreeting: "Уважаемая Дарья Владимировна!",

  executorFio: "А.А. Горшков",
  executorRank: "лейтенант полиции",
  executorPhone: "(960) 108-90-69",
  executorEmail: "agorshkov109@mvd.ru",

  chiefFio: "В.А. Бортников",
  chiefRank: "подполковник полиции",
  deputyFio: "Н.Е. Родин",

  procFioShort: "А.А. Дмитриеву",
  procFioFull: "Алексей Александрович",
  procRank: "старшему советнику юстиции",

  datePostanovlenie: "16.04.2026",
  dateSpravka: "20.04.2026",
  dateOtkaz: "23.04.2026",

  crimeDate: "06.04.2026",
  crimeTime: "14 часов 00 минут",
  crimeDescription: "на телефон позвонил ранее незнакомый мужчина, который представился сотрудником службы доставки СДЭК",

  mincifraPersonFio: "Карпунина Елена Михайловна",
  mincifraPersonBirth: "21.08.1976",
  mincifraDateFrom: "01.11.2025",
  mincifraDateTo: "30.11.2025",
  mincifraDirectorFio: "Цветкову С.В.",
  mincifraDirectorFullName: "Сергей Валерьевич",
}

// ─── helpers ────────────────────────────────────────────────────────────────

const FONT = "Times New Roman"
const SIZE = 24 // 12pt in half-points

function p(text: string, opts: { bold?: boolean; align?: (typeof AlignmentType)[keyof typeof AlignmentType]; indent?: number; before?: number; after?: number } = {}) {
  return new Paragraph({
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing: { before: opts.before ?? 0, after: opts.after ?? 120 },
    indent: opts.indent ? { left: convertInchesToTwip(opts.indent) } : undefined,
    children: [new TextRun({ text, bold: opts.bold, font: FONT, size: SIZE })],
  })
}

function pRuns(runs: { text: string; bold?: boolean; underline?: boolean }[], opts: { align?: (typeof AlignmentType)[keyof typeof AlignmentType]; before?: number; after?: number } = {}) {
  return new Paragraph({
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing: { before: opts.before ?? 0, after: opts.after ?? 120 },
    children: runs.map(r => new TextRun({
      text: r.text,
      bold: r.bold,
      underline: r.underline ? { type: UnderlineType.SINGLE } : undefined,
      font: FONT,
      size: SIZE,
    })),
  })
}

function empty(n = 1) {
  return Array.from({ length: n }, () => p(""))
}

function sigLine(rank: string, name: string) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 120 },
    children: [
      new TextRun({ text: rank, font: FONT, size: SIZE }),
      new TextRun({ text: "\t\t\t\t\t" + name, font: FONT, size: SIZE }),
    ],
  })
}

function headerBlock(d: CaseData) {
  return [
    p("МВД РОСССИ", { bold: true, align: AlignmentType.CENTER }),
    p("ОТДЕЛ ПОЛИЦИИ № 4 УМВД РОССИИ ПО Г. ВОРОНЕЖУ", { bold: true, align: AlignmentType.CENTER }),
    p("ОУР ОП № 4 УМВД РОСССИ ПО Г. ВОРОНЕЖУ", { bold: true, align: AlignmentType.CENTER }),
    ...empty(2),
    p("Материал об отказе в возбуждении уголовного дела", { align: AlignmentType.CENTER }),
    ...empty(1),
    p(`КУСП № ${d.kuspNumber} от ${d.kuspDate}`, { align: AlignmentType.CENTER }),
    ...empty(3),
  ]
}

function executorFooter(d: CaseData) {
  return [
    p("Оперуполномоченный ОРП ИТТ ОУР"),
    p(`ОП № 4 УМВД России по г. Воронежу`),
    sigLine(d.executorRank, d.executorFio),
  ]
}

// ─── ОПИСЬ ──────────────────────────────────────────────────────────────────

export function buildOpis(d: CaseData): Document {
  const rows: TableRow[] = []

  const headerRow = new TableRow({
    children: ["№ п/п", "№ документа", "Краткое содержание документа", "№ листов дела", "Примечание"].map(
      (text, i) => new TableCell({
        width: { size: [8, 8, 50, 17, 17][i], type: WidthType.PERCENTAGE },
        children: [p(text, { bold: true, align: AlignmentType.CENTER })],
      }),
    ),
  })
  rows.push(headerRow)

  const items = [
    ["1", "", "Заявление", "1", ""],
    ["2", "", "Объяснение", "2-3", ""],
    ["3", "", "Постановление", "4", ""],
    ["4", "", "Справка", "5", ""],
    ["5", "", "Запрос МинЦифр", "6", ""],
    ["6", "", "Постановление", "7", ""],
    ["7", "", "Уведомление", "8", ""],
    ["8", "", "Уведомление", "9", ""],
  ]

  for (let i = 0; i < 35; i++) {
    const item = items[i] ?? [`${i + 1}`, "", "", "", ""]
    rows.push(new TableRow({
      children: item.map((text, ci) => new TableCell({
        width: { size: [8, 8, 50, 17, 17][ci], type: WidthType.PERCENTAGE },
        children: [p(text, { align: AlignmentType.CENTER })],
      })),
    }))
  }

  return new Document({
    sections: [{
      children: [
        ...headerBlock(d),
        p("Начат: " + d.dateStart, { align: AlignmentType.RIGHT }),
        p("Окончен: " + d.dateEnd, { align: AlignmentType.RIGHT }),
        ...empty(1),
        p("г. Воронеж"),
        p(d.year + " год"),
        ...empty(2),
        p("ОПИСЬ", { bold: true, align: AlignmentType.CENTER }),
        ...empty(1),
        new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }),
        ...empty(2),
        p("Опись составил(а):"),
        ...executorFooter(d),
      ],
    }],
  })
}

// ─── ПОСТАНОВЛЕНИЕ О ПРОДЛЕНИИ ───────────────────────────────────────────────

export function buildPostanovlenieProdlenie(d: CaseData): Document {
  return new Document({
    sections: [{
      children: [
        p("ПРОДЛИТЬ", { bold: true }),
        p("Срок проверки до 10 суток"),
        p("Начальник отдела полиции № 4"),
        p("УМВД России по г. Воронежу"),
        sigLine(d.chiefRank, d.chiefFio),
        p(`«16» апреля ${d.year} г.`),
        ...empty(2),
        p("ПОСТАНОВЛЕНИЕ", { bold: true, align: AlignmentType.CENTER }),
        p("о возбуждении перед начальником органа дознания ходатайства о продлении срока проверки сообщения о преступлении.", { align: AlignmentType.CENTER }),
        ...empty(1),
        pRuns([
          { text: "г. Воронеж" },
          { text: `\t\t\t\t\t\t\t\t«16» апреля ${d.year} г.` },
        ]),
        ...empty(1),
        p(`Оперуполномоченный ОРП ИТТ ОУР ОП № 4 УМВД России по г. Воронежу младший лейтенант полиции ${d.executorFio}, рассмотрев материал проверки КУСП № ${d.kuspNumber} от ${d.kuspDate} зарегистрированный в отделе полиции № 4 УМВД России по г. Воронежу,`),
        ...empty(1),
        p("УСТАНОВИЛ:", { bold: true }),
        ...empty(1),
        p("Необходимо провести ОРМ, с целью установления принадлежности абонентского номера, а так же провести иные проверочные мероприятия, в которых возникает необходимость, принять решение в порядке ст.ст. 144-145 УПК РФ."),
        p("На основании изложенного и руководствуясь ч. 3 ст. 144 УПК РФ"),
        ...empty(1),
        p("ПОСТАНОВИЛ:", { bold: true }),
        ...empty(1),
        p("Ходатайствовать перед начальником органа дознания о продлении срока проверки сообщения о преступлении до 10 суток."),
        ...empty(3),
        ...executorFooter(d),
      ],
    }],
  })
}

// ─── СПРАВКА ─────────────────────────────────────────────────────────────────

export function buildSpravka(d: CaseData): Document {
  return new Document({
    sections: [{
      children: [
        pRuns([
          { text: "г. Воронеж" },
          { text: `\t\t\t\t\t\t\t\t${d.dateSpravka} г.` },
        ]),
        ...empty(1),
        p("СПРАВКА", { bold: true, align: AlignmentType.CENTER }),
        ...empty(1),
        p(`Сообщаю Вам, что мною в ходе проведения проверки по материалу проверки КУСП ${d.kuspNumber} от ${d.kuspDate} установлено, что ${d.applicantFio} звонили с ранее неизвестных абонентских номеров, которые у неё так же не сохранились.`),
        p(`${d.applicantFio} материальный ущерб не причинён, денежные средства она не переводила.`),
        ...empty(4),
        ...executorFooter(d),
      ],
    }],
  })
}

// ─── ПОСТАНОВЛЕНИЕ ОБ ОТКАЗЕ ─────────────────────────────────────────────────

export function buildPostanovlenieOtkaz(d: CaseData): Document {
  return new Document({
    sections: [{
      children: [
        pRuns([
          { text: "                                                                                          УТВЕРЖДАЮ" },
        ], { align: AlignmentType.RIGHT }),
        p("Начальник отдела полиции № 4", { align: AlignmentType.RIGHT }),
        p("УМВД России по г. Воронеж", { align: AlignmentType.RIGHT }),
        sigLine(d.chiefRank, d.chiefFio),
        p(`«23» апреля ${d.year} г.`, { align: AlignmentType.RIGHT }),
        ...empty(2),
        p("П О С Т А Н О В Л Е Н И Е", { bold: true, align: AlignmentType.CENTER }),
        p("об отказе в возбуждении уголовного дела", { align: AlignmentType.CENTER }),
        ...empty(1),
        pRuns([{ text: `г. Воронеж` }, { text: `\t\t\t\t\t\t\t\t${d.dateOtkaz} г.` }]),
        ...empty(1),
        p(`Оперуполномоченный ОРП ИТТ ОУР ОП № 4 УМВД России по г. Воронежу ${d.executorRank} Горшков А.А., рассмотрев материал проверки, зарегистрированный в отделе полиции № 4 УМВД России по г. Воронежу КУСП № ${d.kuspNumber} от ${d.kuspDate}.`),
        ...empty(1),
        p("У С Т А Н О В И Л:", { bold: true }),
        ...empty(1),
        p(`         ${d.dateStart} в отделе полиции № 4 УМВД России по г. Воронежу зарегистрировано заявление гражданки ${d.applicantFio} по факту передачи своих данных неизвестным лицам.`),
        p(`         В ходе проведения проверки была опрошена ${d.applicantFioFull}, ${d.applicantBirthDate} г.р., которая пояснила, что ${d.crimeDate} примерно в ${d.crimeTime} ${d.crimeDescription}. Денежные средства с банковских карт не списаны, материальный ущерб не причинен.`),
        ...empty(1),
        p("         Принимая во внимание, что имеются достаточные данные, указывающие на отсутствие признаков преступления, предусмотренного ст.159 УК РФ, ст. 272 УК РФ, на основании п. 1 ч. 1 ст. 24 УПК РФ."),
        ...empty(1),
        p("П О С Т А Н О В И Л:", { bold: true }),
        ...empty(1),
        p(`1. Отказать в возбуждении уголовного дела по сообщению о совершении преступления предусмотренного ст.159 УК РФ, ст. 272 УК РФ, на основании п. 1 ч.1 ст. 24 УПК РФ, за отсутствием события преступления.`),
        p(`2. Копию постановления направить заявителю гр. ${d.applicantFio}, прокурору Коминтерновского района г. Воронежа ${d.procRank} ${d.procFioShort}, а также другим заинтересованным лицам.`),
        ...empty(1),
        p("Настоящее постановление может быть обжаловано прокурору Коминтерновского района либо в суд в порядке, установленном главой 16 УПК РФ."),
        ...empty(2),
        ...executorFooter(d),
        ...empty(1),
        p(`Копия настоящего постановления направлена прокурору Коминтерновского района г. Воронежа, заявителю гр. ${d.applicantFio}.`),
        ...empty(2),
        ...executorFooter(d),
        ...empty(2),
        p("СОГЛАСЕН"),
        p("Заместитель начальника ОУР – начальник отделения ИТТ"),
        p("ОП № 4 УМВД России по г. Воронежу"),
        sigLine("капитан полиции", d.deputyFio),
      ],
    }],
  })
}

// ─── УВЕДОМЛЕНИЕ ЗАЯВИТЕЛЮ ───────────────────────────────────────────────────

export function buildUvedomlenieZayavitel(d: CaseData): Document {
  return new Document({
    sections: [{
      children: [
        p("ГУ МВД России по Воронежской области", { bold: true }),
        ...empty(1),
        p("Управление Министерства внутренних дел"),
        p("Российской Федерации"),
        p("(УМВД России по г. Воронежу)"),
        ...empty(1),
        p("ул. Хользунова, 123а, Воронеж, 394088"),
        p("тел. (473) 269-65-24"),
        ...empty(2),
        p(d.applicantFio),
        p(d.applicantAddress),
        ...empty(2),
        p("О направлении ответа", { bold: true }),
        ...empty(2),
        p(d.applicantGreeting),
        ...empty(1),
        p(`Сообщаем Вам, что в отделе полиции № 4 Управления министерства внутренних дел Российской Федерации по г. Воронежу зарегистрирован материал проверки КУСП № ${d.kuspNumber} от ${d.kuspDate}. По данному материалу вынесено постановление об отказе в возбуждении уголовного дела.`),
        p("Вы имеете право обжаловать данное постановление в порядке ст. 124 уголовно – процессуального кодекса Российской Федерации «Порядок рассмотрения жалобы прокурором, руководителем следственного органа), ст. 125 УПК РФ «Судебный порядок рассмотрения жалоб»."),
        ...empty(1),
        p("Приложение: на 1 листе."),
        ...empty(3),
        sigLine("Начальник отдела полиции № 4", d.chiefFio),
      ],
    }],
  })
}

// ─── УВЕДОМЛЕНИЕ ПРОКУРОРУ ───────────────────────────────────────────────────

export function buildUvedomlenieProcuror(d: CaseData): Document {
  return new Document({
    sections: [{
      children: [
        p("ГУ МВД России по Воронежской области", { bold: true }),
        ...empty(1),
        p("Управление Министерства внутренних дел"),
        p("Российской Федерации"),
        p("(ОП №4 УМВД России по г. Воронежу)"),
        ...empty(1),
        p("ул. Хользунова, 123а, Воронеж, 394088"),
        p("тел. (473) 269-65-24"),
        ...empty(2),
        p("Прокурору Коминтерновского района г. Воронежа"),
        p(d.procRank),
        p(d.procFioShort),
        ...empty(2),
        p(`О направлении копии постановления по КУСП № ${d.kuspNumber} от ${d.kuspDate}`, { bold: true }),
        ...empty(2),
        p(`Уважаемый ${d.procFioFull}!`),
        ...empty(1),
        p(`Направляем Вам копию постановления об отказе в возбуждении уголовного дела по материалу проверки КУСП № ${d.kuspNumber} от ${d.kuspDate}.`),
        ...empty(3),
        sigLine("Начальник отдела полиции № 4", d.chiefFio),
        sigLine(d.chiefRank, ""),
        ...empty(3),
        p(`Исп. ${d.executorFio}`),
        p(`тел. ${d.executorPhone}`),
        ...empty(1),
        p(`отп. 2 экз.`),
        p(`1 экз. – в Прокуратуру Коминтерновского района г. Воронежа`),
        p(`2 экз. -  в дело`),
        p(`Исп. ${d.executorFio}`),
        p(`тел. ${d.executorPhone}`),
      ],
    }],
  })
}

// ─── ЗАПРОС МИНЦИФР ──────────────────────────────────────────────────────────

export function buildZaprosMincifra(d: CaseData): Document {
  return new Document({
    sections: [{
      children: [
        p("ГУ МВД России по Воронежской области", { bold: true }),
        ...empty(1),
        p("Управление Министерства внутренних дел"),
        p("Российской Федерации"),
        p("(УМВД России по г. Воронежу)"),
        ...empty(1),
        p("ул. Хользунова, 123а, Воронеж, 394088"),
        p("тел. (473) 269-65-24"),
        ...empty(2),
        p("Директору департамента развития инфраструктуры электронного правительства министерства цифрового развития, связи и массовых коммуникаций Российской Федерации"),
        p(d.mincifraDirectorFio),
        p("Пресненская набережная, д. 10, стр. 2, IQ – квартал, г. Москва, 125039"),
        ...empty(2),
        p("О направлении запроса", { bold: true }),
        ...empty(2),
        p(`Уважаемый ${d.mincifraDirectorFullName}!`),
        ...empty(1),
        p(`В отдел полиции № 4 УМВД России по г. Воронежу поступил материал проверки КУСП № ${d.kuspNumber} от ${d.kuspDate} по факту заявления ${d.applicantFio}, по факту неправомерного доступа к личному кабинету ${d.applicantFio} на Едином портале «Государственные услуги».`),
        p(`На основании ст. 6 Федерального закона «Об оперативно – розыскной деятельности» от 12.08.1995 № 144-ФЗ, прошу сообщить в наш адрес информацию об аккаунте ${d.mincifraPersonFio}, ${d.mincifraPersonBirth} г.р.:`),
        ...empty(1),
        p(`1. Абонентские номера, используемые для регистрации, перерегистрации, авторизации в сервисе за период с 00 часов 00 минут ${d.mincifraDateFrom} по 23 часа 59 минут ${d.mincifraDateTo}.`),
        p(`2. Электронные почтовые ящики, используемые для регистрации, перерегистрации личного кабинета в сервисе за период с 00 часов 00 минут ${d.mincifraDateFrom} по 23 часа 59 минут ${d.mincifraDateTo}.`),
        p(`3. IP-адреса, MAC-адрес, IMEI, cookie-файлы, использовавшиеся при регистрации за период с 00 часов 00 минут ${d.mincifraDateFrom} по 23 часа 59 минут ${d.mincifraDateTo}.`),
        p(`4. Действия, производимые в личном кабинете за период с 00 часов 00 минут ${d.mincifraDateFrom} по 23 часа 59 минут ${d.mincifraDateTo}.`),
        ...empty(1),
        p(`Информацию прошу предоставить в наш адрес, а также продублировать на электронный почтовый ящик ${d.executorEmail}`),
        ...empty(3),
        sigLine("Начальник отдела полиции № 4", d.chiefFio),
        sigLine(d.chiefRank, ""),
        ...empty(3),
        p(`Исп. ${d.executorFio}`),
        p(`тел. ${d.executorPhone}`),
      ],
    }],
  })
}

// ─── Скачивание ──────────────────────────────────────────────────────────────

export async function downloadDoc(doc: Document, filename: string) {
  const blob = await Packer.toBlob(doc)
  saveAs(blob, filename)
}