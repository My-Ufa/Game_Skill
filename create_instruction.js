const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, ExternalHyperlink
} = require("docx");
const fs = require("fs");

const BLUE = "1F4E79";
const LIGHT_BLUE = "D6E4F0";
const MID_BLUE = "2E75B6";
const GRAY = "595959";
const WHITE = "FFFFFF";
const BORDER_COLOR = "BBCFE0";

const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
const borders = { top: border, bottom: border, left: border, right: border };

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, color: BLUE, font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, color: MID_BLUE, font: "Arial" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 180, after: 60 },
    children: [new TextRun({ text, bold: true, size: 24, color: GRAY, font: "Arial" })]
  });
}

function p(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: "000000", ...options })]
  });
}

function code(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 360 },
    children: [new TextRun({ text, font: "Courier New", size: 20, color: "C7254E" })]
  });
}

function bullet(text, bold_prefix = null) {
  const children = [];
  if (bold_prefix) {
    children.push(new TextRun({ text: bold_prefix + " ", bold: true, size: 22, font: "Arial" }));
  }
  children.push(new TextRun({ text, size: 22, font: "Arial" }));
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun("")] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 1 } },
    children: [new TextRun("")]
  });
}

function codeBlock(lines) {
  const rows = lines.map(line =>
    new TableRow({
      children: [new TableCell({
        borders: noBorders,
        width: { size: 8400, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 80, right: 80 },
        children: [new Paragraph({
          spacing: { before: 20, after: 20 },
          children: [new TextRun({ text: line, font: "Courier New", size: 19, color: "C7254E" })]
        })]
      })]
    })
  );
  return new Table({
    width: { size: 8400, type: WidthType.DXA },
    columnWidths: [8400],
    shading: { fill: "F4F4F4", type: ShadingType.CLEAR },
    rows
  });
}

function infoBox(lines) {
  const children = lines.map(line =>
    new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text: line, size: 22, font: "Arial", color: "1F4E79" })]
    })
  );
  return new Table({
    width: { size: 8960, type: WidthType.DXA },
    columnWidths: [8960],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 8960, type: WidthType.DXA },
        shading: { fill: "E8F4FD", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children
      })]
    })]
  });
}

function stepsTable() {
  const steps = [
    ["1", "Транскрибация", "Структурированный конспект встречи"],
    ["2", "Анализ рынка", "Цифры: объём рынка, игроки, динамика"],
    ["3", "Описание игры", "Контекст / конфликт / механики / гипотезы"],
    ["4", "Состав ролей", "10–15 ролей = реальные рос. компании"],
    ["5", "Экономика", "Excel: сценарий MIN/MAX по 3 тактам"],
    ["6", "Плацдарм", "PPTX-схема мира игры"],
    ["7", "Ролевые листы", "PPTX: 1 лист = 1 роль"],
    ["8", "Установка на игру", "PPTX-презентация для старта (14 слайдов)"],
  ];

  const headerRow = new TableRow({
    children: [
      ["№", 700],
      ["Шаг", 2200],
      ["Что получаешь", 5500],
    ].map(([text, size]) => new TableCell({
      borders,
      width: { size, type: WidthType.DXA },
      shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        children: [new TextRun({ text, bold: true, size: 21, font: "Arial", color: WHITE })]
      })]
    }))
  });

  const dataRows = steps.map(([num, name, result], i) =>
    new TableRow({
      children: [
        [num, 700],
        [name, 2200],
        [result, 5500],
      ].map(([text, size]) => new TableCell({
        borders,
        width: { size, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? "FFFFFF" : "EFF6FB", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text, size: 21, font: "Arial" })]
        })]
      }))
    })
  );

  return new Table({
    width: { size: 8400, type: WidthType.DXA },
    columnWidths: [700, 2200, 5500],
    rows: [headerRow, ...dataRows]
  });
}

function scopeTable() {
  const rows_data = [
    ["user scope", "Плагин доступен тебе во всех проектах на этой машине. Хранится в ~/.claude/", "Если работаешь одна"],
    ["project scope", "Плагин записывается в .claude/settings.json репозитория — виден всей команде", "Если команда работает вместе"],
    ["local scope", "Только тебе и только в этом репо. Хранится в settings.local.json (обычно в .gitignore)", "Личная установка в одном репо"],
  ];

  const headerRow = new TableRow({
    children: [
      ["Вариант", 2000],
      ["Что значит", 4200],
      ["Когда выбирать", 2600],
    ].map(([text, size]) => new TableCell({
      borders,
      width: { size, type: WidthType.DXA },
      shading: { fill: MID_BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        children: [new TextRun({ text, bold: true, size: 21, font: "Arial", color: WHITE })]
      })]
    }))
  });

  const dataRows = rows_data.map(([opt, desc, when], i) =>
    new TableRow({
      children: [
        [opt, 2000, true],
        [desc, 4200, false],
        [when, 2600, false],
      ].map(([text, size, mono]) => new TableCell({
        borders,
        width: { size, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? "FFFFFF" : "EFF6FB", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text, size: 21, font: mono ? "Courier New" : "Arial", color: mono ? "C7254E" : "000000" })]
        })]
      }))
    })
  );

  return new Table({
    width: { size: 8800, type: WidthType.DXA },
    columnWidths: [2000, 4200, 2600],
    rows: [headerRow, ...dataRows]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 260 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 260 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: MID_BLUE },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1300, bottom: 1440, left: 1300 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 1 } },
          children: [
            new TextRun({ text: "sd-game-builder", bold: true, size: 20, font: "Arial", color: MID_BLUE }),
            new TextRun({ text: "   —   Инструкция по установке и использованию", size: 20, font: "Arial", color: GRAY }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 1 } },
          children: [
            new TextRun({ text: "Стр. ", size: 18, font: "Arial", color: GRAY }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: GRAY }),
          ]
        })]
      })
    },
    children: [
      // Title block
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 80 },
        children: [new TextRun({ text: "Конструктор СД-игры", bold: true, size: 52, font: "Arial", color: BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "Плагин sd-game-builder для Claude Code", size: 26, font: "Arial", color: GRAY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "Инструкция по установке и использованию", size: 22, font: "Arial", color: GRAY })]
      }),

      divider(),

      // Section 1
      h1("1. Установка плагина"),

      p("Установка занимает 3 команды и не требует перезапуска Claude Code."),
      spacer(),

      h3("Шаг 1. Добавить маркетплейс"),
      codeBlock(["/plugin marketplace add My-Ufa/Game_Skill"]),
      spacer(),

      h3("Шаг 2. Установить плагин"),
      codeBlock(["/plugin install sd-game-builder@Game_Skill"]),
      spacer(),

      p("Claude Code спросит, куда установить. Выбери нужный scope:"),
      spacer(),
      scopeTable(),
      spacer(),
      p("Рекомендация: если работаешь одна — выбирай user scope."),
      spacer(),

      h3("Шаг 3. Активировать без перезапуска"),
      codeBlock(["/reload-plugins"]),

      divider(),

      // Section 2
      h1("2. Как вызвать скилл"),

      h2("Автоматически — просто напиши фразу"),
      p("Скилл срабатывает сам, если в сообщении есть любое из слов:"),
      spacer(),
      infoBox([
        "деловая игра · СД-игра · ситуационно-деятельностная игра · игротехник",
        "плацдарм · ролевые листы · установка на игру · экономика игры",
        "состав ролей · собрать игру · разработать игру · анализ рынка для игры",
      ]),
      spacer(),

      h3("Примеры фраз:"),
      bullet("Хочу собрать СД-игру под ритейл"),
      bullet("Помоги разработать игру для банковского сектора"),
      bullet("Нужны роли для деловой игры про ИИ"),
      bullet("Рассчитай экономику игры"),
      bullet("Создай установку на игру"),
      spacer(),

      h2("Явно — через slash-команду"),
      codeBlock(["/sd-game-builder:sd-game-skill"]),

      divider(),

      // Section 3
      h1("3. Что происходит после запуска"),

      p("Claude предложит выбрать точку входа:"),
      spacer(),
      infoBox([
        "С чего начнём?",
        "1. Есть запись встречи → Шаг 1 (Транскрибация)",
        "2. Отрасль известна, нужны цифры → Шаг 2 (Анализ рынка)",
        "3. Рынок понятен, строим описание → Шаг 3 (Описание игры)",
        "4. Описание есть, подбираем роли → Шаг 4 (Состав ролей)",
        "Или просто скажи, что нужно.",
      ]),
      spacer(),
      p("Шаги не обязаны идти по порядку — можно начать с любого."),
      spacer(),

      p("Как ведёт диалог:", { bold: true }),
      bullet("Задаёт 1–2 вопроса за раз, не вываливает анкету целиком"),
      bullet("Сам считает цифры и предлагает их на проверку"),
      bullet("После каждого шага обновляет накопленное состояние игры (JSON-блок)"),
      bullet("Явно называет пробелы и противоречия"),
      bullet("Ссылается на данные из предыдущих шагов"),

      divider(),

      // Section 4
      h1("4. Все 8 шагов и артефакты"),
      spacer(),
      stepsTable(),

      divider(),

      // Section 5
      h1("5. Полезные команды по ходу работы"),

      h3("Начать с конкретного шага"),
      codeBlock(["Перейди сразу к шагу 4, отрасль — нефтепереработка"]),
      spacer(),

      h3("Скорректировать баланс ролей"),
      codeBlock(["Роль Газпром нефти слишком сильная — сбалансируй"]),
      spacer(),

      h3("Проверить экономику"),
      codeBlock(["Никто не должен уйти в минус на первом такте — проверь"]),
      spacer(),

      h3("Зафиксировать данные для следующих шагов"),
      codeBlock(["Запомни эти цифры для следующих шагов"]),

      divider(),

      // Footer note
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 0 },
        children: [new TextRun({ text: "github.com/My-Ufa/Game_Skill", size: 19, font: "Courier New", color: MID_BLUE })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/Users/dobro/Documents/Game_Skill/sd-game-builder-инструкция.docx", buffer);
  console.log("Done: sd-game-builder-инструкция.docx");
}).catch(err => {
  console.error(err);
  process.exit(1);
});
