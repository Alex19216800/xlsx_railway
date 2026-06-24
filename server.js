const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 20 * 1024 * 1024,
  },
});

const PORT = Number(process.env.PORT || 3000);
const API_TOKEN = String(process.env.API_TOKEN || "").trim();

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
    У полі «отримав водій/експедитор»:
    - ПІБ;
    - ЄДДР, якщо він є;
    - без номера посвідчення водія.
  */
  return [
    fullName,
    eddr
  ]
    .filter(Boolean)
    .join(" ");
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
  const columns = [
    "A", "B", "C", "D", "E", "F",
    "G", "H", "I", "J", "K", "L"
  ];

  for (let row = 29; row <= 33; row += 1) {
    for (const column of columns) {
      sheet.getCell(`${column}${row}`).value = "";
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

    setCell(
      sheet,
      `L${row}`,
      valueFromItem(item, [
        "gross_weight",
        "weight",
      ]),
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

  setCell(
    sheet,
    "L34",
    firstValue(data, [
      "cargo.gross_weight",
      "gross_weight",
      "totals.gross_weight",
    ]),
    {
      horizontal: "right",
      wrapText: false,
    }
  );
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

  setCell(
    sheet,
    "A7",
    automobile
      ? `Автомобіль ${automobile}`
      : "Автомобіль"
  );

  setCell(
    sheet,
    "F7",
    trailer
      ? `Причіп/напівпричіп ${trailer}`
      : "Причіп/напівпричіп"
  );

  setCell(
    sheet,
    "H7",
    transportationType
      ? `Вид перевезень ${transportationType}`
      : "Вид перевезень"
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

  setCell(
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

  setCell(
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

  setCell(
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

  setCell(
    sheet,
    "K21",
    firstValue(data, [
      "cargo.vat_amount_text",
      "vat_amount_text",
      "vat_amount_words",
    ]),
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
    version: "5.0.0",
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

app.listen(PORT, "0.0.0.0", () => {
  console.log(
    `TTN XLSX service v5.0.0 is running on port ${PORT}`
  );
});
