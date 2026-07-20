const fs = require("fs");
const path = require("path");
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 20 * 1024 * 1024,
  },
});

const PORT = Number(process.env.PORT || 3000);
const API_TOKEN = String(process.env.API_TOKEN || "").trim();
const APPLICATION_TEMPLATE_PATH = path.join(
  __dirname,
  "templates",
  "zayavka_perevezennia_template.xlsx"
);
const APPLICATION_DOCX_TEMPLATE_PATH = path.join(
  __dirname,
  "templates",
  "zayavka_perevezennia_docx_template.docx"
);

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

function clean(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function stripDimensionUnit(value) {
  return clean(value)
    .replace(
      /\s*(?:м\.?|метр|метра|метрів)\s*$/iu,
      ""
    )
    .trim();
}

function stripWordsMarker(value) {
  return clean(value)
    .replace(
      /\(\s*словами(?:\s*,\s*з\s+урахуванням\s+ПДВ)?\s*\)/giu,
      ""
    )
    .replace(/\s+/g, " ")
    .trim();
}

function buildReceivedDriverText(data) {
  /*
    Насамперед використовуємо вже готове поле з n8n.
    Воно має формат: ПІБ, ЄДДР.
  */
  const readyText = clean(
    firstValue(data, [
      "driver.received_driver_field",
      "received_driver_field",
    ])
  );

  if (readyText) {
    return readyText;
  }

  const fullName = clean(
    firstValue(data, [
      "driver.full_name",
      "received_driver_full_name",
      "driver_full_name",
      "received_driver",
    ])
  );

  const eddr = clean(
    firstValue(data, [
      "driver.eddr",
      "received_driver_eddr",
      "driver_eddr",
      "eddr",
    ])
  );

  /*
    Резервне формування:
    - ПІБ;
    - ЄДДР, якщо він є;
    - без номера посвідчення водія;
    - частини розділяються комою і пробілом.
  */
  return [fullName, eddr]
    .filter(Boolean)
    .join(", ");
}

function getByPath(object, path) {
  return String(path)
    .split(".")
    .reduce((current, key) => {
      if (
        current === null ||
        current === undefined ||
        typeof current !== "object"
      ) {
        return undefined;
      }

      return current[key];
    }, object);
}

function firstValue(object, paths, fallback = "") {
  for (const path of paths) {
    const value = getByPath(object, path);

    if (
      value !== undefined &&
      value !== null &&
      clean(value) !== ""
    ) {
      return value;
    }
  }

  return fallback;
}

function normalizePayload(payload) {
  const candidate =
    payload.correct_ttn_data ??
    payload.data ??
    payload.result ??
    payload;

  if (
    candidate &&
    typeof candidate === "object" &&
    (
      candidate.generation_ready !== undefined ||
      candidate.vehicle ||
      candidate.carrier ||
      candidate.driver ||
      candidate.supplier ||
      candidate.client ||
      candidate.route ||
      candidate.cargo
    )
  ) {
    return candidate;
  }

  return (
    candidate.correct_document ??
    candidate.output ??
    candidate.document ??
    candidate
  );
}

function parseDate(value) {
  const text = clean(value);

  let match = text.match(
    /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/
  );

  if (match) {
    return {
      day: match[1].padStart(2, "0"),
      month: match[2].padStart(2, "0"),
      year: match[3],
    };
  }

  match = text.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/
  );

  if (match) {
    return {
      day: match[3].padStart(2, "0"),
      month: match[2].padStart(2, "0"),
      year: match[1],
    };
  }

  return {
    day: "",
    month: "",
    year: "",
  };
}

function monthNameUkrainian(month) {
  const months = {
    "01": "січня",
    "02": "лютого",
    "03": "березня",
    "04": "квітня",
    "05": "травня",
    "06": "червня",
    "07": "липня",
    "08": "серпня",
    "09": "вересня",
    "10": "жовтня",
    "11": "листопада",
    "12": "грудня",
  };

  return months[String(month).padStart(2, "0")] || "";
}

function safeFileName(value) {
  const cleaned = clean(value)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\.+$/g, "")
    .slice(0, 180);

  return cleaned || "TTN_generated.xlsx";
}

function withUnit(value, unit) {
  const text = clean(value);

  if (!text) return "";
  if (new RegExp(`(?:^|\\s)${unit}$`, "i").test(text)) {
    return text;
  }

  return `${text} ${unit}`;
}

function setCell(sheet, address, value, options = {}) {
  const cell = sheet.getCell(address);
  cell.value = value === undefined || value === null
    ? ""
    : value;

  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: options.wrapText ?? true,
    shrinkToFit: options.shrinkToFit ?? false,
    vertical: options.vertical ?? "center",
    horizontal:
      options.horizontal ??
      cell.alignment?.horizontal ??
      "left",
  };
}

function cloneStyle(value) {
  if (!value || typeof value !== "object") {
    return value;
  }

  return JSON.parse(JSON.stringify(value));
}

function applyCellAlignment(cell, options = {}) {
  cell.alignment = {
    ...(cell.alignment || {}),
    wrapText: options.wrapText ?? true,
    shrinkToFit: options.shrinkToFit ?? false,
    vertical: options.vertical ?? "center",
    horizontal:
      options.horizontal ??
      cell.alignment?.horizontal ??
      "left",
  };
}

/*
  У merged-полях рядка 7 назву поля залишаємо без підкреслення,
  а вставлене значення підкреслюємо.

  Приклад:
  Автомобіль MAN TGX...
  де "Автомобіль" — без підкреслення,
  "MAN TGX..." — з підкресленням.
*/
function setLabelWithUnderlinedValue(
  sheet,
  address,
  label,
  value,
  options = {}
) {
  const cell = sheet.getCell(address);
  const cleanedLabel = clean(label);
  const cleanedValue = clean(value);
  const baseFont = cloneStyle(cell.font || {});

  const labelFont = {
    ...baseFont,
    underline: false,
  };

  const valueFont = {
    ...baseFont,
    underline: true,
  };

  cell.value = {
    richText: cleanedValue
      ? [
          {
            text: `${cleanedLabel} `,
            font: labelFont,
          },
          {
            text: cleanedValue,
            font: valueFont,
          },
        ]
      : [
          {
            text: cleanedLabel,
            font: labelFont,
          },
        ],
  };

  applyCellAlignment(cell, options);
}

/*
  Для окремих комірок габаритів у клітинці міститься тільки
  вставлене значення, тому підкреслюємо весь текст/число.
  Підписи "(довжина, м)", "(ширина, м)", "(висота, м)"
  розташовані в інших комірках і не змінюються.
*/
function setUnderlinedCellValue(
  sheet,
  address,
  value,
  options = {}
) {
  const cell = sheet.getCell(address);

  cell.value = value === undefined || value === null
    ? ""
    : value;

  cell.font = {
    ...cloneStyle(cell.font || {}),
    underline: true,
  };

  applyCellAlignment(cell, options);
}

/*
  Білий фон для товарних рядків.

  У деяких XLSX-шаблонах зелена заливка зберігається не тільки
  безпосередньо у cell.fill, а й у спільному стилі або в умовному
  форматуванні. Тому фінально:
  1) відокремлюємо стиль кожної комірки;
  2) примусово задаємо білий фон;
  3) видаляємо умовне форматування, яке перетинає товарні рядки.
*/
const CARGO_WHITE_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: {
    argb: "FFFFFFFF",
  },
  bgColor: {
    argb: "FFFFFFFF",
  },
};

function forceWhiteCellBackground(cell) {
  cell.style = cloneStyle(cell.style || {});
  cell.fill = cloneStyle(CARGO_WHITE_FILL);
}

function columnLettersToNumber(value) {
  const letters = clean(value).toUpperCase();

  if (!/^[A-Z]+$/.test(letters)) {
    return null;
  }

  let result = 0;

  for (const letter of letters) {
    result = result * 26 +
      (letter.charCodeAt(0) - 64);
  }

  return result;
}

function parseSheetReferencePart(value) {
  const part = clean(value)
    .replace(/\$/g, "")
    .split("!")
    .pop();

  if (!part) {
    return null;
  }

  // Звичайний діапазон клітинок, наприклад A29:L33.
  let match = part.match(
    /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i
  );

  if (match) {
    const left = columnLettersToNumber(match[1]);
    const top = Number(match[2]);
    const right = columnLettersToNumber(
      match[3] || match[1]
    );
    const bottom = Number(match[4] || match[2]);

    if (
      Number.isFinite(left) &&
      Number.isFinite(top) &&
      Number.isFinite(right) &&
      Number.isFinite(bottom)
    ) {
      return {
        left: Math.min(left, right),
        right: Math.max(left, right),
        top: Math.min(top, bottom),
        bottom: Math.max(top, bottom),
      };
    }
  }

  // Цілі колонки, наприклад A:A або A:L.
  match = part.match(/^([A-Z]+)(?::([A-Z]+))?$/i);

  if (match) {
    const left = columnLettersToNumber(match[1]);
    const right = columnLettersToNumber(
      match[2] || match[1]
    );

    if (
      Number.isFinite(left) &&
      Number.isFinite(right)
    ) {
      return {
        left: Math.min(left, right),
        right: Math.max(left, right),
        top: 1,
        bottom: 1048576,
      };
    }
  }

  // Цілі рядки, наприклад 29:33.
  match = part.match(/^(\d+)(?::(\d+))?$/);

  if (match) {
    const top = Number(match[1]);
    const bottom = Number(match[2] || match[1]);

    if (
      Number.isFinite(top) &&
      Number.isFinite(bottom)
    ) {
      return {
        left: 1,
        right: 16384,
        top: Math.min(top, bottom),
        bottom: Math.max(top, bottom),
      };
    }
  }

  return null;
}

function rangesIntersect(first, second) {
  return !(
    first.right < second.left ||
    first.left > second.right ||
    first.bottom < second.top ||
    first.top > second.bottom
  );
}

function referenceIntersectsCargoRows(reference) {
  const cargoRange = {
    left: 1,
    right: 12,
    top: 29,
    bottom: 33,
  };

  return String(reference ?? "")
    .split(/[\s,]+/)
    .map(parseSheetReferencePart)
    .filter(Boolean)
    .some(range => {
      return rangesIntersect(range, cargoRange);
    });
}

function removeCargoConditionalFormatting(sheet) {
  if (
    typeof sheet.removeConditionalFormatting ===
    "function"
  ) {
    sheet.removeConditionalFormatting(cf => {
      return !referenceIntersectsCargoRows(cf?.ref);
    });
    return;
  }

  if (Array.isArray(sheet.conditionalFormattings)) {
    sheet.conditionalFormattings =
      sheet.conditionalFormattings.filter(cf => {
        return !referenceIntersectsCargoRows(cf?.ref);
      });
  }
}

function removeCargoRowBackgrounds(sheet) {
  removeCargoConditionalFormatting(sheet);

  for (let row = 29; row <= 33; row += 1) {
    for (let column = 1; column <= 12; column += 1) {
      forceWhiteCellBackground(
        sheet.getCell(row, column)
      );
    }
  }
}

function toNumericCellValue(value) {
  if (
    typeof value === "number" &&
    Number.isFinite(value)
  ) {
    return value;
  }

  let text = clean(value)
    .replace(/[₴грнuah]/giu, "")
    .replace(/[\s  ]/g, "")
    .trim();

  if (!text) {
    return "";
  }

  const commaIndex = text.lastIndexOf(",");
  const dotIndex = text.lastIndexOf(".");

  if (commaIndex >= 0 && dotIndex >= 0) {
    if (commaIndex > dotIndex) {
      text = text
        .replace(/\./g, "")
        .replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (commaIndex >= 0) {
    text = text.replace(",", ".");
  }

  text = text.replace(/[^0-9.+-]/g, "");

  const number = Number(text);

  return Number.isFinite(number)
    ? number
    : value;
}

/*
  Усі значення маси брутто у payload зберігаються в тоннах:
  "18.900" означає 18.900 т.

  У комірки XLSX також передаємо числове значення у тоннах:
  18.900 т -> 18.9
  7.500 т  -> 7.5

  Відображення трьох знаків після коми забезпечує формат комірки
  у шаблоні, наприклад: 18,900.
*/
function grossWeightTonnesToCellValue(value) {
  const tonnes = toNumericCellValue(value);

  if (
    typeof tonnes !== "number" ||
    !Number.isFinite(tonnes)
  ) {
    return "";
  }

  return Number(tonnes.toFixed(3));
}

function applyTransportationWarning(
  sheet,
  transportationType,
  isValid
) {
  if (isValid) {
    return;
  }

  const cell = sheet.getCell("H7");

  /*
    У шаблоні кілька підписів використовують спільний об'єкт стилю.
    Якщо змінити fill/font напряму, ExcelJS може застосувати зміну
    до інших клітинок із тим самим стилем.

    Спочатку повністю відокремлюємо стиль H7 від спільного стилю,
    а потім змінюємо тільки фон.
  */
  cell.style = cloneStyle(cell.style || {});

  /*
    Не змінюємо:
    - текст;
    - шрифт;
    - колір шрифту;
    - жирність;
    - вирівнювання;
    - інші поля шаблону.

    Підсвічується тільки merged-поле H7:J7
    світло-червоним фоном.
  */
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: "FFFFC7CE",
    },
  };
}

function clearCargoRows(sheet) {
  for (let row = 29; row <= 33; row += 1) {
    for (let column = 1; column <= 12; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.value = "";
      forceWhiteCellBackground(cell);
    }
  }

  for (const address of ["G34", "I34", "L34"]) {
    sheet.getCell(address).value = "";
  }
}

function cargoItemsFrom(data) {
  const items =
    firstValue(data, [
      "cargo.items",
      "cargo_items",
      "items",
    ], []);

  return Array.isArray(items) ? items : [];
}

function valueFromItem(item, paths, fallback = "") {
  return firstValue(item || {}, paths, fallback);
}

function fillCargoTable(sheet, data) {
  clearCargoRows(sheet);

  const items = cargoItemsFrom(data).slice(0, 5);

  items.forEach((item, index) => {
    const row = 29 + index;

    setCell(sheet, `A${row}`, index + 1, {
      horizontal: "center",
      wrapText: false,
    });

    setCell(
      sheet,
      `B${row}`,
      valueFromItem(item, [
        "name",
        "cargo_name",
        "title",
      ])
    );

    setCell(
      sheet,
      `C${row}`,
      valueFromItem(item, [
        "animal_id",
        "animal_identification_number",
      ])
    );

    setCell(
      sheet,
      `D${row}`,
      valueFromItem(item, [
        "animal_type",
        "type_of_animal",
      ])
    );

    setCell(
      sheet,
      `E${row}`,
      valueFromItem(item, [
        "temperature",
        "temperature_mode",
        "temperature_regime",
      ])
    );

    setCell(
      sheet,
      `F${row}`,
      valueFromItem(item, [
        "unit",
        "unit_of_measurement",
      ]),
      {
        horizontal: "center",
        wrapText: false,
      }
    );

    setCell(
      sheet,
      `G${row}`,
      valueFromItem(item, [
        "quantity_places",
        "places_count",
        "quantity",
        "places",
      ]),
      {
        horizontal: "center",
        wrapText: false,
      }
    );

    setCell(
      sheet,
      `H${row}`,
      valueFromItem(item, [
        "price_without_vat",
        "price",
        "unit_price",
      ]),
      {
        horizontal: "right",
        wrapText: false,
      }
    );

    setCell(
      sheet,
      `I${row}`,
      valueFromItem(item, [
        "total_with_vat",
        "total_sum_with_vat",
        "sum",
      ]),
      {
        horizontal: "right",
        wrapText: false,
      }
    );

    setCell(
      sheet,
      `J${row}`,
      valueFromItem(item, [
        "package_type",
        "packaging",
      ])
    );

    setCell(
      sheet,
      `K${row}`,
      valueFromItem(item, [
        "cargo_document",
        "document",
      ])
    );

    /*
      Маса брутто товарного рядка у payload зберігається в тоннах.
      У XLSX також записуємо числове значення у тоннах:
      "18.900" -> 18.9.
    */
    setCell(
      sheet,
      `L${row}`,
      grossWeightTonnesToCellValue(
        valueFromItem(item, [
          "gross_weight",
          "weight",
        ])
      ),
      {
        horizontal: "right",
        wrapText: false,
      }
    );
  });

  setCell(
    sheet,
    "G34",
    firstValue(data, [
      "cargo.total_places",
      "places_count",
      "cargo_total_places",
      "totals.places_count",
    ]),
    {
      horizontal: "center",
      wrapText: false,
    }
  );

  setCell(
    sheet,
    "I34",
    firstValue(data, [
      "cargo.total_sum_with_vat",
      "cargo.vat_amount",
      "total_sum_with_vat",
      "totals.total_sum_with_vat",
      "vat_amount",
    ]),
    {
      horizontal: "right",
      wrapText: false,
    }
  );

  /*
    L34 — загальна маса брутто.
    У payload значення зберігається у тоннах, наприклад "18.900".
    У XLSX також записуємо число у тоннах: 18.9.
    Форматування комірки виконує шаблон.
  */
  setCell(
    sheet,
    "L34",
    grossWeightTonnesToCellValue(
      firstValue(data, [
        "cargo.gross_weight",
        "gross_weight",
        "totals.gross_weight",
      ])
    ),
    {
      horizontal: "right",
      wrapText: false,
    }
  );

  /*
    Фінальний прохід після запису даних.
    Гарантовано прибирає зелене оформлення з A29:L33,
    включно зі спільними стилями та умовним форматуванням.
  */
  removeCargoRowBackgrounds(sheet);
}

function buildTransportWeightsText(data) {
  const raw = firstValue(data, [
    "vehicle.raw_transport_weights_text",
    "raw_transport_weights_text",
  ]);

  if (clean(raw)) return clean(raw);

  const values = [
    firstValue(data, [
      "vehicle.transport_empty_weight_total",
      "transport_empty_weight_total",
    ]),
    firstValue(data, [
      "vehicle.transport_max_loaded_weight",
      "transport_max_loaded_weight",
    ]),
    firstValue(data, [
      "vehicle.transport_actual_gross_weight",
      "transport_actual_gross_weight",
    ]),
  ].map(clean);

  if (values.every(Boolean)) {
    return values.join(", ");
  }

  return values.filter(Boolean).join(", ");
}

function fillWorkbook(sheet, data) {
  const documentNumber = clean(
    firstValue(data, [
      "document.number",
      "document.number",
      "document_number",
      "ttn_number",
      "number",
    ])
  );

  const documentDate = clean(
    firstValue(data, [
      "document.date",
      "document.date",
      "document_date",
      "ttn_date",
      "date",
    ])
  );

  const date = parseDate(documentDate);
  const title = [
    documentNumber
      ? `№ ${documentNumber}`
      : "",
    date.day && date.month && date.year
      ? `"${date.day}" ${monthNameUkrainian(date.month)} ${date.year} року`
      : documentDate,
  ]
    .filter(Boolean)
    .join(" ");

  setCell(sheet, "A3", title, {
    horizontal: "center",
    shrinkToFit: true,
  });

  setCell(
    sheet,
    "C5",
    firstValue(data, [
      "document.place",
      "document_place",
      "place",
    ])
  );

  const automobile = clean(
    firstValue(data, [
      "vehicle.automobile",
      "automobile",
      "truck",
      "expected.automobile",
    ])
  );

  const trailer = clean(
    firstValue(data, [
      "vehicle.trailer",
      "trailer",
      "expected.trailer",
    ])
  );

  const transportationType = clean(
    firstValue(data, [
      "transportation.type",
      "transportation_type",
      "transportation_kind",
    ])
  );

  const transportationIsValid =
    getByPath(
      data,
      "transportation.is_valid"
    ) === true;

  setLabelWithUnderlinedValue(
    sheet,
    "A7",
    "Автомобіль",
    automobile
  );

  setLabelWithUnderlinedValue(
    sheet,
    "F7",
    "Причіп/напівпричіп",
    trailer
  );

  setLabelWithUnderlinedValue(
    sheet,
    "H7",
    "Вид перевезень",
    transportationType
  );

  applyTransportationWarning(
    sheet,
    transportationType,
    transportationIsValid
  );

  setCell(
    sheet,
    "C9",
    firstValue(data, [
      "carrier.text",
      "carrier",
      "expected.carrier",
    ])
  );

  setCell(
    sheet,
    "H9",
    firstValue(data, [
      "driver.driver_field",
      "driver",
      "expected.driver",
    ])
  );

  setCell(
    sheet,
    "C11",
    firstValue(data, [
      "supplier.text",
      "supplier",
      "sender",
    ])
  );

  setCell(
    sheet,
    "C13",
    firstValue(data, [
      "client.text",
      "client",
      "customer",
      "receiver",
    ])
  );

  setCell(
    sheet,
    "C15",
    firstValue(data, [
      "route.loading_point",
      "loading_point",
      "route_from",
    ])
  );

  setCell(
    sheet,
    "J15",
    firstValue(data, [
      "route.unloading_point",
      "unloading_point",
      "route_to",
    ])
  );

  setCell(
    sheet,
    "C17",
    stripWordsMarker(
      firstValue(data, [
        "cargo.places_count_words",
        "places_count_words",
      ])
    ),
    {
      horizontal: "center",
    }
  );

  setCell(
    sheet,
    "F17",
    stripWordsMarker(
      firstValue(data, [
        "cargo.gross_weight_words",
        "gross_weight_words",
      ])
    ),
    {
      horizontal: "center",
    }
  );

  setCell(
    sheet,
    "J17",
    buildReceivedDriverText(data)
  );

  setUnderlinedCellValue(
    sheet,
    "F19",
    stripDimensionUnit(
      firstValue(
        data,
        ["dimensions.length", "length"]
      )
    ),
    {
      horizontal: "center",
      wrapText: false,
    }
  );

  setUnderlinedCellValue(
    sheet,
    "G19",
    stripDimensionUnit(
      firstValue(
        data,
        ["dimensions.width", "width"]
      )
    ),
    {
      horizontal: "center",
      wrapText: false,
    }
  );

  setUnderlinedCellValue(
    sheet,
    "I19",
    stripDimensionUnit(
      firstValue(
        data,
        ["dimensions.height", "height"]
      )
    ),
    {
      horizontal: "center",
      wrapText: false,
    }
  );

  setCell(
    sheet,
    "J19",
    buildTransportWeightsText(data),
    {
      horizontal: "center",
      shrinkToFit: true,
    }
  );

  setCell(
    sheet,
    "D21",
    stripWordsMarker(
      firstValue(data, [
        "cargo.total_sum_words",
        "total_sum_words",
      ])
    )
  );

  /*
    K21 — окрема сума ПДВ.
    Використовуємо тільки розраховане перевірене поле з
    Build Correct TTN Data. OCR-текст і сума словами тут
    більше не використовуються.
  */
  setCell(
    sheet,
    "K21",
    toNumericCellValue(
      firstValue(data, [
        "cargo.vat_amount_calculated",
        "vat_amount_calculated",
        "cargo.vat_amount_text",
      ])
    ),
    {
      horizontal: "right",
      shrinkToFit: true,
    }
  );

  setCell(
    sheet,
    "D23",
    firstValue(data, [
      "document.cargo_document_text",
      "cargo_document_text",
    ])
  );

  fillCargoTable(sheet, data);
}


function truthy(value) {
  const text = clean(value).toLowerCase();

  return [
    "1",
    "true",
    "yes",
    "y",
    "так",
    "json"
  ].includes(text);
}

function parseJsonField(value, fieldName) {
  if (value === undefined || value === null || value === "") {
    return {};
  }

  if (typeof value === "object") {
    return value;
  }

  try {
    return JSON.parse(String(value));
  } catch (error) {
    const message = fieldName
      ? `Field '${fieldName}' is not valid JSON.`
      : "Request body is not valid JSON.";

    const customError = new Error(message);
    customError.statusCode = 400;
    throw customError;
  }
}

function normalizeApplicationPayload(payload) {
  const candidate =
    payload.application_data ??
    payload.applicationData ??
    payload.data ??
    payload.result ??
    payload;

  if (
    candidate &&
    typeof candidate === "object" &&
    candidate.application_data &&
    typeof candidate.application_data === "object"
  ) {
    return candidate.application_data;
  }

  return candidate && typeof candidate === "object"
    ? candidate
    : {};
}

function parseRequestPayload(req) {
  if (req.body && req.body.payload !== undefined) {
    return parseJsonField(req.body.payload, "payload");
  }

  return req.body && typeof req.body === "object"
    ? req.body
    : {};
}

function flattenObject(value, prefix = "", output = {}) {
  if (
    value === null ||
    value === undefined ||
    typeof value !== "object" ||
    value instanceof Date ||
    Buffer.isBuffer(value)
  ) {
    if (prefix) output[prefix] = value;
    return output;
  }

  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      flattenObject(item, prefix ? `${prefix}.${index}` : String(index), output);
    });
    return output;
  }

  for (const [key, item] of Object.entries(value)) {
    flattenObject(item, prefix ? `${prefix}.${key}` : key, output);
  }

  return output;
}

function datePartsForApplication(value) {
  const parsed = parseDate(value);

  return {
    day: parsed.day,
    dayNoZero: parsed.day ? String(Number(parsed.day)) : "",
    month: parsed.month,
    monthName: monthNameUkrainian(parsed.month),
    year: parsed.year,
  };
}

function extractApplicationDateParts(value) {
  const text = clean(value);

  let match = text.match(
    /(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/
  );

  if (match) {
    return {
      day: match[1].padStart(2, "0"),
      month: match[2].padStart(2, "0"),
      year: match[3],
    };
  }

  match = text.match(
    /(\d{4})-(\d{1,2})-(\d{1,2})/
  );

  if (match) {
    return {
      day: match[3].padStart(2, "0"),
      month: match[2].padStart(2, "0"),
      year: match[1],
    };
  }

  return {
    day: "",
    month: "",
    year: "",
  };
}

function formatApplicationDate(parts) {
  if (!parts.day || !parts.month || !parts.year) {
    return "";
  }

  return `${parts.day}.${parts.month}.${parts.year}`;
}

function addApplicationDays(value, days) {
  const parts = extractApplicationDateParts(value);

  if (!parts.day || !parts.month || !parts.year) {
    return "";
  }

  const date = new Date(Date.UTC(
    Number(parts.year),
    Number(parts.month) - 1,
    Number(parts.day)
  ));

  if (Number.isNaN(date.getTime())) {
    return "";
  }

  date.setUTCDate(date.getUTCDate() + Number(days || 0));

  return [
    String(date.getUTCDate()).padStart(2, "0"),
    String(date.getUTCMonth() + 1).padStart(2, "0"),
    String(date.getUTCFullYear()),
  ].join(".");
}

function applicationTimeSuffix(value) {
  const text = clean(value);
  const match = text.match(/(?:^|\s)(?:з\s*)?(\d{1,2}:\d{2})(?:\s.*)?$/iu);

  return match && match[1]
    ? match[1]
    : "";
}

function escapeRegExp(value) {
  return String(value ?? "")
    .replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeComparableText(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[’'`]/g, "")
    .replace(/[–—−]/g, "-")
    .replace(/[^a-zа-яіїєґ0-9]+/giu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function removeExactText(value, textToRemove) {
  const source = clean(value);
  const target = clean(textToRemove);

  if (!source || !target) {
    return source;
  }

  return source
    .replace(new RegExp(escapeRegExp(target), "giu"), " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildApplicationTruckType(data) {
  /*
    Поле договору-заявки «ТИП АВТОМОБІЛЯ» повинно містити
    тільки значення поля «Тип авто» з адмінки.

    Не додаємо:
    - тип кузова авто;
    - тип причепа;
    - кузов причепа;
    - модель або номер причепа.

    Плоске поле truck_type має найвищий пріоритет, тому що саме
    його редагує користувач в адмінці.
  */
  let result = clean(
    firstValue(data, [
      "truck_type",
      "truck.vehicle_type_text",
      "truck.vehicle_type",
      "truck.type",
      "vehicle.truck.vehicle_type_text",
      "vehicle.truck.vehicle_type",
      "vehicle.truck.type",
      "truck_vehicle_type",
      "automobile_type",
    ])
  );

  if (!result) {
    return "";
  }

  /*
    Захист для старих клієнтів, які могли передати в truck_type
    комбінований рядок «тип авто / тип причепа».
  */
  const trailerValues = [
    firstValue(data, [
      "trailer_type",
      "trailer.vehicle_type_text",
      "trailer.vehicle_type",
      "trailer.type",
      "vehicle.trailer.vehicle_type_text",
      "vehicle.trailer.vehicle_type",
      "vehicle.trailer.type",
      "trailer_vehicle_type",
    ]),
    firstValue(data, [
      "trailer_body_type",
      "trailer.body_type",
      "vehicle.trailer.body_type",
      "trailer_body",
    ]),
    firstValue(data, [
      "trailer_model",
      "trailer.brand_model",
      "trailer.model",
      "vehicle.trailer.brand_model",
      "vehicle.trailer.model",
    ]),
    firstValue(data, [
      "trailer_plate",
      "trailer.plate",
      "vehicle.trailer.plate",
    ]),
  ].filter(value => clean(value));

  for (const trailerValue of trailerValues) {
    result = removeExactText(result, trailerValue);
  }

  const segments = result
    .split(/\s*(?:\/|\||\+|;)\s*|[\r\n]+/u)
    .map(item => clean(item))
    .filter(Boolean)
    .filter(item =>
      !/(?:^|\s)(?:причіп|напівпричіп|прицеп|полуприцеп|semi-?trailer|trailer)(?:\s|$)/iu.test(item)
    );

  if (segments.length > 0) {
    result = segments[0];
  }

  result = result
    .replace(
      /\s+(?:причіп|напівпричіп|прицеп|полуприцеп|semi-?trailer|trailer)\s*[:\-–—]?\s*.*$/iu,
      ""
    )
    .replace(/[\s,;:.\-–—/|+]+$/u, "")
    .replace(/\s+/g, " ")
    .trim();

  return result;
}

function buildApplicationAliases(data) {
  const aliases = {};
  const appDate = firstValue(data, [
    "application_date",
    "date",
    "loading_date",
    "loading_datetime",
  ]);
  const appDateParts = datePartsForApplication(appDate);

  aliases.application_day = appDateParts.dayNoZero || appDateParts.day;
  aliases.application_month = appDateParts.month;
  aliases.application_month_name = appDateParts.monthName;
  aliases.application_year = appDateParts.year;

  aliases.loading_date = firstValue(data, [
    "loading_date",
    "loading_datetime",
    "application_date",
  ]);

  const servicePeriodValue = firstValue(data, [
    "service_period",
    "transport_period",
    "loading_date",
    "loading_datetime",
    "application_date",
  ]);

  const serviceStartDate = formatApplicationDate(
    extractApplicationDateParts(servicePeriodValue)
  );
  const serviceEndDate = addApplicationDays(
    serviceStartDate,
    2
  );

  const existingUnloadingDateTime = firstValue(data, [
    "unloading_datetime",
    "unloading_date",
  ]);
  const unloadingTime = applicationTimeSuffix(
    existingUnloadingDateTime
  );

  aliases.service_period =
    serviceStartDate && serviceEndDate
      ? `${serviceStartDate} – ${serviceEndDate}`
      : firstValue(data, ["service_period", "transport_period"]);

  // У DOCX-шаблоні поле «СТРОК ПЕРЕВЕЗЕННЯ» використовує
  // placeholder {{transport_period}}. Воно повинно формуватися
  // за тією самою схемою: перша дата – перша дата + 2 дні.
  aliases.transport_period = aliases.service_period;

  aliases.unloading_date =
    serviceEndDate ||
    firstValue(data, [
      "unloading_date",
      "unloading_datetime",
    ]);

  aliases.unloading_datetime = serviceEndDate
    ? [serviceEndDate, unloadingTime]
        .filter(Boolean)
        .join(" ")
    : existingUnloadingDateTime;

  aliases.truck_type =
    buildApplicationTruckType(data) ||
    firstValue(data, ["truck_type"]);

  aliases.route_from = firstValue(data, [
    "route_from",
    "loading_address",
  ]);
  aliases.route_to = firstValue(data, [
    "route_to",
    "unloading_address",
  ]);

  aliases.price = firstValue(data, [
    "price",
    "price_text",
  ]);

  const priceTextValue = clean(aliases.price);
  if (/\bбез\s*пдв\b/iu.test(priceTextValue)) {
    aliases.vat_mode = "без ПДВ";
    aliases.price = priceTextValue.replace(/\s*грн\s*без\s*пдв\s*$/iu, "").trim();
  } else if (/\bз\s*пдв\b/iu.test(priceTextValue)) {
    aliases.vat_mode = "з ПДВ";
    aliases.price = priceTextValue.replace(/\s*грн\s*з\s*пдв\s*$/iu, "").trim();
  } else {
    aliases.vat_mode = firstValue(data, ["vat_mode", "price_vat_mode"]);
    aliases.price = priceTextValue.replace(/\s*грн\s*$/iu, "").trim();
  }

  aliases.carrier_full_name = firstValue(data, [
    "carrier_full_name",
    "carrier_name",
  ]);
  aliases.carrier_tax_id = firstValue(data, [
    "carrier_tax_id",
    "carrier_edrpou",
    "carrier_ipn",
  ]);
  aliases.carrier_address = firstValue(data, ["carrier_address"]);
  aliases.carrier_iban = firstValue(data, ["carrier_iban", "iban"]);
  aliases.carrier_bank = firstValue(data, ["carrier_bank", "bank"]);
  aliases.carrier_mfo = firstValue(data, ["carrier_mfo", "mfo"]);
  aliases.carrier_signer = firstValue(data, [
    "carrier_signer",
    "carrier_director",
  ]);

  aliases.manager_phone = firstValue(data, ["manager_phone"]);
  aliases.manager_name = firstValue(data, ["manager_name"]);
  aliases.loading_time = firstValue(data, ["loading_time"]);

  return aliases;
}

function placeholderValue(data, key, flattened, aliases) {
  const forcedApplicationAliasKeys = new Set([
    "truck_type",
    "service_period",
    "transport_period",
    "unloading_date",
    "unloading_datetime",
  ]);

  if (
    forcedApplicationAliasKeys.has(key) &&
    Object.prototype.hasOwnProperty.call(aliases, key) &&
    clean(aliases[key]) !== ""
  ) {
    return aliases[key];
  }

  const direct = getByPath(data, key);

  if (
    direct !== undefined &&
    direct !== null &&
    clean(direct) !== ""
  ) {
    return direct;
  }

  if (
    Object.prototype.hasOwnProperty.call(aliases, key) &&
    clean(aliases[key]) !== ""
  ) {
    return aliases[key];
  }

  if (
    Object.prototype.hasOwnProperty.call(flattened, key) &&
    flattened[key] !== undefined &&
    flattened[key] !== null
  ) {
    return flattened[key];
  }

  return "";
}

function replacePlaceholdersInString(text, data, flattened, aliases) {
  return String(text).replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (match, rawKey) => {
    const key = clean(rawKey);
    const value = placeholderValue(data, key, flattened, aliases);
    return value === undefined || value === null
      ? ""
      : String(value);
  });
}

function fillApplicationWorkbook(workbook, data) {
  const flattened = flattenObject(data);
  const aliases = buildApplicationAliases(data);

  for (const sheet of workbook.worksheets) {
    const lowerName = clean(sheet.name).toLowerCase();

    if (["мапінг", "mapping", "map", "технічний"].includes(lowerName)) {
      sheet.state = "hidden";
    }

    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (typeof cell.value === "string" && cell.value.includes("{{")) {
          cell.value = replacePlaceholdersInString(
            cell.value,
            data,
            flattened,
            aliases
          );
          cell.alignment = {
            ...(cell.alignment || {}),
            wrapText: true,
            vertical: cell.alignment?.vertical || "middle",
          };
          return;
        }

        if (
          cell.value &&
          typeof cell.value === "object" &&
          Array.isArray(cell.value.richText)
        ) {
          const text = cell.value.richText.map((part) => part.text || "").join("");
          if (text.includes("{{")) {
            cell.value = replacePlaceholdersInString(
              text,
              data,
              flattened,
              aliases
            );
            cell.alignment = {
              ...(cell.alignment || {}),
              wrapText: true,
              vertical: cell.alignment?.vertical || "middle",
            };
          }
        }
      });
    });
  }

  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;
}


function applicationDocxFileName(data, requestedFileName) {
  const baseName = applicationFileName(data, requestedFileName)
    .replace(/\.xlsx$/iu, ".docx")
    .replace(/\.xls$/iu, ".docx")
    .replace(/\.doc$/iu, ".docx");

  if (/\.docx$/iu.test(baseName)) {
    return safeFileName(baseName);
  }

  return safeFileName(`${baseName}.docx`);
}

function applicationTemplateData(data) {
  const flattened = flattenObject(data);
  const aliases = buildApplicationAliases(data);

  const output = {
    ...flattened,
    ...aliases,
    ...data,
  };

  /*
    Ці поля формуються сервером за правилами заявки
    й повинні мати пріоритет над старими значеннями,
    які могла надіслати закешована версія адмінки.
  */
  for (const key of [
    "truck_type",
    "service_period",
    "transport_period",
    "unloading_date",
    "unloading_datetime",
  ]) {
    if (
      Object.prototype.hasOwnProperty.call(aliases, key) &&
      clean(aliases[key]) !== ""
    ) {
      output[key] = aliases[key];
    }
  }

  return output;
}

function docxErrorText(error) {
  if (error && error.properties && Array.isArray(error.properties.errors)) {
    return error.properties.errors
      .map((item) => item && item.message ? item.message : String(item))
      .filter(Boolean)
      .join("; ");
  }

  return error && error.message
    ? error.message
    : "Failed to fill application DOCX template.";
}

function fillApplicationDocx(templateBuffer, data) {
  const zip = new PizZip(templateBuffer);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: {
      start: "{{",
      end: "}}",
    },
    nullGetter() {
      return "";
    },
  });

  doc.render(applicationTemplateData(data));

  return doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

function applicationFileName(data, requestedFileName) {
  if (clean(requestedFileName)) {
    return safeFileName(requestedFileName);
  }

  const filename = firstValue(data, ["filename", "file_name"]);
  if (clean(filename)) {
    return safeFileName(filename);
  }

  const date = firstValue(data, [
    "application_date",
    "loading_date",
    "loading_datetime",
  ], "без_дати");
  const route = firstValue(data, ["route_text"], "заявка");

  return safeFileName(`Заявка_${date}_${route}.xlsx`);
}

function requireApiToken(req, res, next) {
  if (!API_TOKEN) {
    next();
    return;
  }

  const received = clean(
    req.get("x-api-key") ||
    req.get("authorization")
  ).replace(/^Bearer\s+/i, "");

  if (received !== API_TOKEN) {
    res.status(401).json({
      ok: false,
      error: "Unauthorized",
    });
    return;
  }

  next();
}

app.get("/health", (req, res) => {
  res.json({
    ok: true,
    service: "ttn-xlsx-service",
    version: "5.9.2",
  });
});

app.post(
  "/fill-ttn",
  requireApiToken,
  upload.single("template"),
  async (req, res) => {
    try {
      if (!req.file || !req.file.buffer) {
        res.status(400).json({
          ok: false,
          error:
            "Missing XLSX template in multipart field 'template'.",
        });
        return;
      }

      if (!req.body.payload) {
        res.status(400).json({
          ok: false,
          error:
            "Missing JSON payload in multipart field 'payload'.",
        });
        return;
      }

      let payload;

      try {
        payload = JSON.parse(req.body.payload);
      } catch (error) {
        res.status(400).json({
          ok: false,
          error: "Field 'payload' is not valid JSON.",
        });
        return;
      }

      const data = normalizePayload(payload);

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);

      const sheet =
        workbook.getWorksheet("TTN") ||
        workbook.worksheets[0];

      if (!sheet) {
        res.status(400).json({
          ok: false,
          error: "The workbook has no worksheets.",
        });
        return;
      }

      fillWorkbook(sheet, data);

      workbook.calcProperties.fullCalcOnLoad = true;
      workbook.calcProperties.forceFullCalc = true;

      const documentNumber = clean(
        firstValue(data, [
          "document.number",
          "document_number",
          "ttn_number",
          "number",
        ])
      );

      const documentDate = clean(
        firstValue(data, [
          "document.date",
          "document_date",
          "ttn_date",
          "date",
        ])
      );

      const requestedFileName = clean(
        req.body.outputFileName
      );

      const generatedFileName = safeFileName(
        requestedFileName ||
        `ТТН_${documentNumber || "без_номера"}_${documentDate || "без_дати"}.xlsx`
      );

      const outputBuffer =
        await workbook.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename*=UTF-8''${encodeURIComponent(generatedFileName)}`
      );

      res.setHeader(
        "X-Generated-File-Name",
        encodeURIComponent(generatedFileName)
      );

      res.status(200).send(Buffer.from(outputBuffer));
    } catch (error) {
      console.error(error);

      res.status(500).json({
        ok: false,
        error:
          error && error.message
            ? error.message
            : "Failed to fill TTN template.",
      });
    }
  }
);



async function fillApplicationXlsxHandler(req, res) {
  try {
    let payload;

    try {
      payload = parseRequestPayload(req);
    } catch (error) {
      res.status(error.statusCode || 400).json({
        ok: false,
        error: error.message,
      });
      return;
    }

    const data = normalizeApplicationPayload(payload);

    if (!data || typeof data !== "object") {
      res.status(400).json({
        ok: false,
        error: "Missing application_data object.",
      });
      return;
    }

    const templateBuffer = req.file && req.file.buffer
      ? req.file.buffer
      : fs.readFileSync(APPLICATION_TEMPLATE_PATH);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBuffer);

    fillApplicationWorkbook(workbook, data);

    const generatedFileName = applicationFileName(
      data,
      clean(req.body?.outputFileName || payload.outputFileName)
    );

    const outputBuffer = await workbook.xlsx.writeBuffer();
    const mimeType =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    const shouldReturnJson =
      truthy(req.body?.returnJson) ||
      truthy(req.body?.return_json) ||
      truthy(payload.returnJson) ||
      truthy(payload.return_json) ||
      truthy(req.query.returnJson) ||
      truthy(req.query.return_json) ||
      clean(req.query.format).toLowerCase() === "json";

    res.setHeader("Cache-Control", "no-store");

    if (shouldReturnJson) {
      res.status(200).json({
        ok: true,
        filename: generatedFileName,
        mime_type: mimeType,
        file_base64: Buffer.from(outputBuffer).toString("base64"),
      });
      return;
    }

    res.setHeader("Content-Type", mimeType);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeURIComponent(generatedFileName)}`
    );
    res.setHeader(
      "X-Generated-File-Name",
      encodeURIComponent(generatedFileName)
    );

    res.status(200).send(Buffer.from(outputBuffer));
  } catch (error) {
    console.error(error);

    res.status(500).json({
      ok: false,
      error:
        error && error.message
          ? error.message
          : "Failed to fill application XLSX template.",
    });
  }
}

async function fillApplicationDocxHandler(req, res) {
  try {
    let payload;

    try {
      payload = parseRequestPayload(req);
    } catch (error) {
      res.status(error.statusCode || 400).json({
        ok: false,
        error: error.message,
      });
      return;
    }

    const data = normalizeApplicationPayload(payload);

    if (!data || typeof data !== "object") {
      res.status(400).json({
        ok: false,
        error: "Missing application_data object.",
      });
      return;
    }

    const templateBuffer = req.file && req.file.buffer
      ? req.file.buffer
      : fs.readFileSync(APPLICATION_DOCX_TEMPLATE_PATH);

    let outputBuffer;

    try {
      outputBuffer = fillApplicationDocx(templateBuffer, data);
    } catch (error) {
      res.status(400).json({
        ok: false,
        error: docxErrorText(error),
      });
      return;
    }

    const generatedFileName = applicationDocxFileName(
      data,
      clean(
        req.body?.outputFileName ||
        payload.outputFileName ||
        payload.filename
      )
    );

    const mimeType =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    const shouldReturnJson =
      truthy(req.body?.returnJson) ||
      truthy(req.body?.return_json) ||
      truthy(payload.returnJson) ||
      truthy(payload.return_json) ||
      truthy(req.query.returnJson) ||
      truthy(req.query.return_json) ||
      clean(req.query.format).toLowerCase() === "json";

    res.setHeader("Cache-Control", "no-store");

    if (shouldReturnJson) {
      res.status(200).json({
        ok: true,
        filename: generatedFileName,
        mime_type: mimeType,
        file_base64: Buffer.from(outputBuffer).toString("base64"),
      });
      return;
    }

    res.setHeader("Content-Type", mimeType);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeURIComponent(generatedFileName)}`
    );
    res.setHeader(
      "X-Generated-File-Name",
      encodeURIComponent(generatedFileName)
    );

    res.status(200).send(Buffer.from(outputBuffer));
  } catch (error) {
    console.error(error);

    res.status(500).json({
      ok: false,
      error:
        error && error.message
          ? error.message
          : "Failed to fill application DOCX template.",
    });
  }
}

/*
  Канонічний endpoint заявки тепер завжди формує DOCX.

  Раніше:
  - /fill-application формував XLSX;
  - /fill-application-docx формував DOCX.

  Через це різні або закешовані версії адмінки могли отримувати різний формат.

  Тепер:
  - /fill-application       -> DOCX;
  - /fill-application-docx  -> DOCX;
  - /fill-application-xlsx  -> XLSX, лише для явного запиту Excel.
*/
app.post(
  "/fill-application-xlsx",
  requireApiToken,
  upload.single("template"),
  fillApplicationXlsxHandler
);

app.post(
  ["/fill-application", "/fill-application-docx"],
  requireApiToken,
  upload.single("template"),
  fillApplicationDocxHandler
);

app.listen(PORT, "0.0.0.0", () => {
  console.log(
    `TTN XLSX/DOCX service v5.9.2 is running on port ${PORT}`
  );
});
