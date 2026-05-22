const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, LevelFormat, PageBreak
} = require("docx");
const fs = require("fs");

// Colors
const BLUE    = "1F4E79";
const MBLU    = "2E75B6";
const LGRAY   = "595959";
const WHITE   = "FFFFFF";
const BC      = "C8D8E8";

// ── primitive helpers ────────────────────────────────────────────────────────

function run(text, opts) {
  return new TextRun(Object.assign({ text, font: "Arial", size: 22, color: "1A1A1A" }, opts || {}));
}

function monoRun(text, color) {
  return new TextRun({ text, font: "Courier New", size: 20, color: color || "A8FF60" });
}

function para(children, opts) {
  return new Paragraph(Object.assign({ spacing: { before: 60, after: 60 }, children }, opts || {}));
}

function gap() {
  return para([run(" ")], { spacing: { before: 40, after: 40 } });
}

function divider() {
  return para([run(" ")], {
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BC, space: 1 } }
  });
}

function h1(text) {
  return para([run(text, { bold: true, size: 34, color: BLUE })], {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 }
  });
}

function h2(text) {
  return para([run(text, { bold: true, size: 26, color: MBLU })], {
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 80 }
  });
}

function h3(text) {
  return para([run(text, { bold: true, size: 23 })], {
    spacing: { before: 200, after: 60 }
  });
}

function txt(text) {
  return para([run(text)]);
}

function bul(text) {
  return para([run(text)], {
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 }
  });
}

function num(children) {
  if (typeof children === "string") children = [run(children)];
  return para(children, {
    numbering: { reference: "steps", level: 0 },
    spacing: { before: 80, after: 80 }
  });
}

// ── single-row, single-cell box ──────────────────────────────────────────────

function box(paragraphs, bg, borderColor) {
  var bc2 = borderColor || BC;
  var b = { style: BorderStyle.SINGLE, size: 1, color: bc2 };
  var borders = { top: b, bottom: b, left: b, right: b };
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: borders,
            width: { size: 9026, type: WidthType.DXA },
            shading: { fill: bg, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            children: paragraphs
          })
        ]
      })
    ]
  });
}

function note(text, type) {
  var configs = {
    info:  { bg: "EFF6FB", fg: MBLU,    icon: "ℹ  " },
    ok:    { bg: "EBF5EB", fg: "375623", icon: "✓  " },
    warn:  { bg: "FFF4E5", fg: "7F3F00", icon: "⚠  " },
    error: { bg: "FFF0F0", fg: "C00000", icon: "✗  " }
  };
  var c = configs[type || "info"];
  return box(
    [para([new TextRun({ text: c.icon + text, font: "Arial", size: 21, color: c.fg })],
          { spacing: { before: 0, after: 0 } })],
    c.bg
  );
}

function codeBox(lines) {
  var nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  var noBorders = { top: nb, bottom: nb, left: nb, right: nb };
  var rows = lines.map(function(line) {
    return new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: 9026, type: WidthType.DXA },
          shading: { fill: "1E1E1E", type: ShadingType.CLEAR },
          margins: { top: 28, bottom: 28, left: 200, right: 200 },
          children: [
            para([monoRun(line || " ")], { spacing: { before: 8, after: 8 } })
          ]
        })
      ]
    });
  });
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: rows
  });
}

function errorBox(lines) {
  var nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  var noBorders = { top: nb, bottom: nb, left: nb, right: nb };
  var rows = lines.map(function(line) {
    return new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: 9026, type: WidthType.DXA },
          shading: { fill: "2D0000", type: ShadingType.CLEAR },
          margins: { top: 28, bottom: 28, left: 200, right: 200 },
          children: [
            para([monoRun(line || " ", "FF8080")], { spacing: { before: 8, after: 8 } })
          ]
        })
      ]
    });
  });
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: rows
  });
}

function stepBadge(num2, title, desc) {
  var b = { style: BorderStyle.SINGLE, size: 1, color: BC };
  var borders = { top: b, bottom: b, left: b, right: b };
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [800, 8226],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: borders,
            width: { size: 800, type: WidthType.DXA },
            shading: { fill: MBLU, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 80, right: 80 },
            verticalAlign: "center",
            children: [
              para([run(num2, { bold: true, size: 32, color: WHITE })],
                   { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 } })
            ]
          }),
          new TableCell({
            borders: borders,
            width: { size: 8226, type: WidthType.DXA },
            shading: { fill: "EFF6FB", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 180, right: 120 },
            children: [
              para([run(title, { bold: true, size: 24, color: BLUE })],
                   { spacing: { before: 0, after: 20 } }),
              para([run(desc, { size: 21, color: LGRAY })],
                   { spacing: { before: 0, after: 0 } })
            ]
          })
        ]
      })
    ]
  });
}

function tableRow(cells, widths, headerBg) {
  var b = { style: BorderStyle.SINGLE, size: 1, color: BC };
  var borders = { top: b, bottom: b, left: b, right: b };
  return new TableRow({
    children: cells.map(function(cellData, i) {
      return new TableCell({
        borders: borders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: headerBg || "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 130, right: 130 },
        children: [
          para(
            Array.isArray(cellData.runs) ? cellData.runs : [run(cellData.text || "", cellData.opts)],
            { spacing: { before: 0, after: 0 } }
          )
        ]
      });
    })
  });
}

// ── page break ───────────────────────────────────────────────────────────────
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ── DOCUMENT ─────────────────────────────────────────────────────────────────

var headerPar = new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BC, space: 1 } },
  children: [
    run("Конструктор СД-игры", { bold: true, size: 19, color: MBLU }),
    run("   ·   Полная инструкция установки для Mac", { size: 19, color: LGRAY })
  ]
});

var footerPar = new Paragraph({
  alignment: AlignmentType.CENTER,
  border: { top: { style: BorderStyle.SINGLE, size: 4, color: BC, space: 1 } },
  children: [
    run("Стр. ", { size: 18, color: LGRAY }),
    new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: LGRAY }),
    run(" / ", { size: 18, color: LGRAY }),
    new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: LGRAY })
  ]
});

// ── cheatsheet table ─────────────────────────────────────────────────────────
var cheatRows = [
  ["node --version",                                   "Проверить есть ли Node.js"],
  ["npm install -g @anthropic-ai/claude-code",         "Установить Claude Code"],
  ["claude --version",                                 "Проверить установку Claude Code"],
  ["claude",                                           "Запустить Claude Code"],
  ["/plugin marketplace add My-Ufa/Game_Skill",        "Добавить маркетплейс плагина"],
  ["/plugin install sd-game-builder@Game_Skill",       "Установить плагин"],
  ["/reload-plugins",                                  "Активировать плагины без перезапуска"],
  ["/sd-game-builder:sd-game-skill",                   "Явно вызвать скилл"],
  ["sudo chown -R $(whoami) /usr/local/lib/node_modules", "Исправить ошибку прав доступа"],
];

var b = { style: BorderStyle.SINGLE, size: 1, color: BC };
var brd = { top: b, bottom: b, left: b, right: b };

var cheatTableRows = [
  new TableRow({
    children: [
      new TableCell({
        borders: brd, width: { size: 3800, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 130, right: 130 },
        children: [para([run("Команда", { bold: true, size: 21, color: WHITE })], { spacing: { before: 0, after: 0 } })]
      }),
      new TableCell({
        borders: brd, width: { size: 5226, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 130, right: 130 },
        children: [para([run("Что делает", { bold: true, size: 21, color: WHITE })], { spacing: { before: 0, after: 0 } })]
      })
    ]
  })
].concat(cheatRows.map(function(row, i) {
  return new TableRow({
    children: [
      new TableCell({
        borders: brd, width: { size: 3800, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? "F4F4F4" : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 130, right: 130 },
        children: [para([new TextRun({ text: row[0], font: "Courier New", size: 18, color: "C7254E" })], { spacing: { before: 0, after: 0 } })]
      }),
      new TableCell({
        borders: brd, width: { size: 5226, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? "F4F4F4" : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 130, right: 130 },
        children: [para([run(row[1], { size: 21 })], { spacing: { before: 0, after: 0 } })]
      })
    ]
  });
}));

var cheatTable = new Table({
  width: { size: 9026, type: WidthType.DXA },
  columnWidths: [3800, 5226],
  rows: cheatTableRows
});

// ── scope table ───────────────────────────────────────────────────────────────
var scopeRows = [
  ["user scope", "Плагин доступен тебе во всех проектах на этой машине", "Если работаешь одна — выбирай этот"],
  ["project scope", "Плагин для всей команды в этом репозитории (.claude/settings.json)", "Командная работа"],
  ["local scope", "Только тебе и только в текущей папке (settings.local.json)", "Личная установка в одном проекте"],
];

var scopeTableRows = [
  new TableRow({
    children: [
      new TableCell({ borders: brd, width: { size: 2600, type: WidthType.DXA }, shading: { fill: MBLU, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 130, right: 130 }, children: [para([run("Вариант", { bold: true, size: 21, color: WHITE })], { spacing: { before: 0, after: 0 } })] }),
      new TableCell({ borders: brd, width: { size: 4026, type: WidthType.DXA }, shading: { fill: MBLU, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 130, right: 130 }, children: [para([run("Что значит", { bold: true, size: 21, color: WHITE })], { spacing: { before: 0, after: 0 } })] }),
      new TableCell({ borders: brd, width: { size: 2400, type: WidthType.DXA }, shading: { fill: MBLU, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 130, right: 130 }, children: [para([run("Когда выбирать", { bold: true, size: 21, color: WHITE })], { spacing: { before: 0, after: 0 } })] }),
    ]
  })
].concat(scopeRows.map(function(row, i) {
  var bg = i % 2 === 0 ? WHITE : "EFF6FB";
  return new TableRow({
    children: [
      new TableCell({ borders: brd, width: { size: 2600, type: WidthType.DXA }, shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 130, right: 130 }, children: [para([new TextRun({ text: row[0], font: "Courier New", size: 19, color: "C7254E" })], { spacing: { before: 0, after: 0 } })] }),
      new TableCell({ borders: brd, width: { size: 4026, type: WidthType.DXA }, shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 130, right: 130 }, children: [para([run(row[1], { size: 21 })], { spacing: { before: 0, after: 0 } })] }),
      new TableCell({ borders: brd, width: { size: 2400, type: WidthType.DXA }, shading: { fill: bg, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 130, right: 130 }, children: [para([run(row[2], { size: 21 })], { spacing: { before: 0, after: 0 } })] }),
    ]
  });
}));

var scopeTable = new Table({
  width: { size: 9026, type: WidthType.DXA },
  columnWidths: [2600, 4026, 2400],
  rows: scopeTableRows
});

// ── terminal output box ──────────────────────────────────────────────────────
var terminalBox = (function() {
  var nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  var noBorders = { top: nb, bottom: nb, left: nb, right: nb };
  var lines2 = [
    "С чего начнём?",
    "1. Есть запись встречи → Шаг 1 (Транскрибация)",
    "2. Отрасль известна, нужны цифры → Шаг 2 (Анализ рынка)",
    "3. Рынок понятен, строим описание → Шаг 3 (Описание игры)",
    "4. Описание есть, подбираем роли → Шаг 4 (Состав ролей)",
    "Или просто скажи, что нужно."
  ];
  var rows2 = lines2.map(function(line, i) {
    return new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: 9026, type: WidthType.DXA },
          shading: { fill: "1E1E1E", type: ShadingType.CLEAR },
          margins: { top: i === 0 ? 120 : 10, bottom: i === lines2.length-1 ? 120 : 10, left: 200, right: 200 },
          children: [
            para(
              [new TextRun({ text: line, font: "Courier New", size: 20, color: i === 0 ? "FFFFFF" : "A8FF60", bold: i === 0 })],
              { spacing: { before: 8, after: 8 } }
            )
          ]
        })
      ]
    });
  });
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: rows2 });
})();

// ── assemble document ────────────────────────────────────────────────────────

var children = [

  // TITLE
  para([run("Конструктор СД-игры", { bold: true, size: 56, color: BLUE })],
    { alignment: AlignmentType.CENTER, spacing: { before: 400, after: 100 } }),
  para([run("Плагин sd-game-builder для Claude Code", { size: 28, color: LGRAY })],
    { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 } }),
  para([run("Полная инструкция установки с нуля  ·  Mac", { size: 22, color: LGRAY })],
    { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 } }),
  gap(),
  note("Инструкция рассчитана на тех, кто никогда не пользовался терминалом и не устанавливал разработческие инструменты. Следуй шагам по порядку.", "info"),
  divider(),

  // OVERVIEW
  h1("Что нам понадобится"),
  txt("Для работы плагина нужно последовательно установить три вещи:"),
  gap(),
  stepBadge("1", "Node.js", "Среда выполнения JavaScript — нужна для работы Claude Code"),
  gap(),
  stepBadge("2", "Claude Code", "Приложение от Anthropic, в котором работает плагин"),
  gap(),
  stepBadge("3", "Плагин sd-game-builder", "Сам скилл — три команды внутри Claude Code"),
  divider(),

  // STEP 1 — Terminal
  h1("Шаг 1. Открываем Терминал"),
  txt("Терминал — программа, уже установленная на каждом Mac. Искать её не надо."),
  gap(),
  h3("Как открыть:"),
  num([run("Нажми "), run("Command (⌘) + Пробел", { bold: true }), run(" — откроется Spotlight")]),
  num([run("Напечатай "), run("Терминал", { bold: true }), run(" или "), run("Terminal", { bold: true })]),
  num([run("Нажми Enter")]),
  gap(),
  txt("Откроется окно с мигающим курсором. Это и есть Терминал."),
  gap(),
  note("В терминале ты вводишь команды и нажимаешь Enter. Команды выполняются на твоём компьютере.", "info"),
  divider(),

  // STEP 2 — Node.js
  h1("Шаг 2. Устанавливаем Node.js"),
  txt("Сначала проверь — возможно, Node.js уже есть. Введи в терминале:"),
  gap(),
  codeBox(["node --version"]),
  gap(),
  note("Если ответ выглядит как  v20.x.x  — Node.js уже установлен. Переходи к Шагу 3.", "ok"),
  note("Если ответ  command not found  — нужно установить. Читай дальше.", "error"),
  gap(),
  h2("Установка через официальный сайт (проще всего)"),
  num([run("Перейди на сайт "), run("nodejs.org", { bold: true, color: MBLU })]),
  num([run("Нажми большую зелёную кнопку "), run("LTS", { bold: true }), run(" (не Current)")]),
  num([run("Скачается файл .pkg — открой его и нажимай «Продолжить»")]),
  num([run("После установки закрой терминал, открой снова и проверь:")]),
  gap(),
  codeBox(["node --version"]),
  gap(),
  note("Должно появиться что-то вроде  v20.18.0  — значит всё OK.", "ok"),
  divider(),

  // STEP 3 — Claude Code
  h1("Шаг 3. Устанавливаем Claude Code"),
  txt("Введи в терминале:"),
  gap(),
  codeBox(["npm install -g @anthropic-ai/claude-code"]),
  gap(),
  txt("Подождёт 1–2 минуты, появится много текста — это нормально."),
  gap(),
  h2("Если появилась ошибка EACCES (нет прав)"),
  txt("Ты можешь увидеть такую ошибку:"),
  gap(),
  errorBox([
    "npm error code EACCES",
    "npm error syscall mkdir",
    "npm error Error: EACCES: permission denied, mkdir '/usr/local/lib/node_modules'",
  ]),
  gap(),
  txt("Это означает: нет прав на запись в системную папку. Исправляется одной командой:"),
  gap(),
  codeBox(["sudo chown -R $(whoami) /usr/local/lib/node_modules"]),
  gap(),
  note("Терминал попросит пароль от твоего Mac. Вводи его — символы не будут видны, это нормально. Нажми Enter.", "warn"),
  gap(),
  txt("После этого снова запусти установку:"),
  gap(),
  codeBox(["npm install -g @anthropic-ai/claude-code"]),
  gap(),
  h2("Проверить что Claude Code установился"),
  codeBox(["claude --version"]),
  gap(),
  note("Если видишь номер версии — установка прошла успешно.", "ok"),
  divider(),

  // STEP 4 — Login
  pageBreak(),
  h1("Шаг 4. Входим в аккаунт"),
  txt("Запускаем Claude Code:"),
  gap(),
  codeBox(["claude"]),
  gap(),
  txt("При первом запуске откроется браузер для входа в аккаунт Anthropic."),
  gap(),
  note("Входи в тот же аккаунт, что используешь на claude.ai. Нужен платный план Pro или выше.", "warn"),
  gap(),
  txt("После входа вернись в терминал — там появится приглашение Claude Code. Теперь можно общаться с Claude прямо из терминала."),
  divider(),

  // STEP 5 — Plugin
  h1("Шаг 5. Устанавливаем плагин"),
  txt("Теперь ты внутри Claude Code. Вводи команды в строку ввода — там же, где пишешь вопросы Claude."),
  gap(),
  h2("Команда 1 — добавить маркетплейс"),
  codeBox(["/plugin marketplace add My-Ufa/Game_Skill"]),
  gap(),
  txt("Claude Code скачает список плагинов из репозитория."),
  gap(),
  h2("Команда 2 — установить плагин"),
  codeBox(["/plugin install sd-game-builder@Game_Skill"]),
  gap(),
  txt("Claude Code спросит, куда установить. Выбери стрелками и нажми Enter:"),
  gap(),
  scopeTable,
  gap(),
  note("Рекомендую выбрать  user scope  — плагин будет доступен тебе во всех проектах.", "info"),
  gap(),
  h2("Команда 3 — активировать без перезапуска"),
  codeBox(["/reload-plugins"]),
  gap(),
  note("Плагин установлен. Он будет загружаться автоматически при каждом запуске Claude Code.", "ok"),
  divider(),

  // STEP 6 — Use
  h1("Шаг 6. Проверяем и используем"),
  txt("Напиши прямо в Claude Code любую из этих фраз:"),
  gap(),
  codeBox([
    "Хочу собрать СД-игру",
    "Помоги разработать деловую игру",
    "Нужны роли для игры про ритейл",
  ]),
  gap(),
  txt("Claude ответит как игротехник и предложит выбрать шаг:"),
  gap(),
  terminalBox,
  gap(),
  note("Шаги не обязаны идти по порядку — можно начать с любого.", "info"),
  gap(),
  h2("Как запускать в следующий раз"),
  txt("Каждый раз просто открываешь Терминал и пишешь:"),
  gap(),
  codeBox(["claude"]),
  gap(),
  note("Устанавливать плагин повторно не нужно. Он уже установлен и загружается автоматически.", "ok"),
  divider(),

  // TROUBLESHOOTING
  h1("Решение типичных проблем"),

  h2("«command not found: claude» после установки"),
  txt("npm установил Claude Code, но терминал его не видит. Введи:"),
  gap(),
  codeBox(["export PATH=\"$PATH:$(npm bin -g)\""]),
  gap(),
  txt("Затем снова:"),
  gap(),
  codeBox(["claude"]),
  gap(),

  h2("«npm error EACCES: permission denied»"),
  txt("Нет прав на системную папку. Введи:"),
  gap(),
  codeBox(["sudo chown -R $(whoami) /usr/local/lib/node_modules"]),
  gap(),
  txt("Введи пароль Mac (символы не видны), нажми Enter. Затем:"),
  gap(),
  codeBox(["npm install -g @anthropic-ai/claude-code"]),
  gap(),

  h2("«/plugin: command not found»"),
  note("Команды /plugin работают ТОЛЬКО внутри Claude Code. Убедись, что ты запустил команду claude и находишься в чате с Claude, а не просто в обычном терминале.", "error"),
  gap(),

  h2("Claude Code не открывает браузер для входа"),
  codeBox(["claude auth login"]),
  gap(),

  h2("Скилл не срабатывает на фразу — явный вызов"),
  codeBox(["/sd-game-builder:sd-game-skill"]),
  divider(),

  // CHEATSHEET
  pageBreak(),
  h1("Шпаргалка — все команды"),
  gap(),
  cheatTable,
  divider(),

  para([run("github.com/My-Ufa/Game_Skill", { size: 19, font: "Courier New", color: MBLU })],
    { alignment: AlignmentType.CENTER, spacing: { before: 200, after: 0 } }),
];

var doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 260 } } } }] },
      { reference: "steps",   levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 260 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: MBLU },
        paragraph: { spacing: { before: 280, after: 80 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1300, right: 1200, bottom: 1300, left: 1200 }
      }
    },
    headers: { default: new Header({ children: [headerPar] }) },
    footers: { default: new Footer({ children: [footerPar] }) },
    children: children
  }]
});

Packer.toBuffer(doc).then(function(buf) {
  fs.writeFileSync("/Users/dobro/Documents/Game_Skill/sd-game-builder-установка-с-нуля.docx", buf);
  console.log("Done.");
}).catch(function(e) {
  console.error(e);
  process.exit(1);
});
