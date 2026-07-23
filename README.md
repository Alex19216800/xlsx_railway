# TTN XLSX Service v5

Railway-сервіс для заповнення XLSX-шаблону товарно-транспортної накладної.

## Endpoint

### GET /health

Перевірка, що сервіс працює.

### POST /fill-ttn

`multipart/form-data`:

- `template` — XLSX-файл шаблону;
- `payload` — JSON-рядок з `correct_ttn_data` або повним output ноди `Build Correct TTN Data`;
- `outputFileName` — необов’язкова назва готового файла.

Відповідь: готовий XLSX-файл.

## Railway

1. Створи новий GitHub-репозиторій.
2. Завантаж у нього файли цього архіву.
3. У Railway натисни `New Project → Deploy from GitHub Repo`.
4. Додай змінну:
   - `API_TOKEN` — довільний складний секретний рядок.
5. Railway автоматично виконає `npm install` і `npm start`.
6. Перевір:
   - `https://YOUR-SERVICE.up.railway.app/health`

## n8n: HTTP Request

- Method: `POST`
- URL: `https://YOUR-SERVICE.up.railway.app/fill-ttn`
- Authentication: None
- Header:
  - `x-api-key: <API_TOKEN>`
- Send Body: ON
- Body Content Type: `Multipart Form-Data`

Параметри:

1. Parameter Type: `n8n Binary File`
   - Name: `template`
   - Input Data Field Name: `data`

2. Parameter Type: `Form Data`
   - Name: `payload`
   - Value:
     `{{ JSON.stringify($('Build Correct TTN Data').first().json.correct_ttn_data || $('Build Correct TTN Data').first().json) }}`

3. Parameter Type: `Form Data`
   - Name: `outputFileName`
   - Value:
     `{{ 'ТТН_' + ($('Build Correct TTN Data').first().json.correct_ttn_data?.document_number || $('Build Correct TTN Data').first().json.document?.document_number || 'без_номера') + '.xlsx' }}`

Response:
- Response Format: `File`
- Put Output in Field: `data`

## Google Drive

Перед HTTP Request:

- Google Drive node
- Resource: `File`
- Operation: `Download`
- File: вибрати XLSX-шаблон
- Put Output File in Field: `data`

Після HTTP Request можна зробити дві гілки:

1. Google Drive — Upload готового XLSX.
2. Telegram — Send Document з Binary Property `data`.

## Обмеження першої версії

- Максимум 5 позицій вантажу — рядки 29–33.
- Рядок 34 — підсумки.
- Шаблон повинен містити аркуш `TTN`.
- Якщо аркуша `TTN` немає, використовується перший аркуш.


## Fix in v2

The service now reads the full nested `correct_ttn_data` object instead of
discarding everything except `correct_ttn_data.document`.

Supported nested sections:
`document`, `vehicle`, `transportation`, `carrier`, `driver`,
`supplier`, `client`, `route`, `dimensions`, `cargo`.


## Зміни v3

1. Довжина, ширина та висота записуються без `м`.
2. У полі «отримав водій/експедитор» записуються ПІБ і ЄДДР без посвідчення.
3. Із текстових полів видаляється службовий напис `(словами)`.
4. Неправильний або відсутній «Вид перевезень» не блокує створення XLSX.
5. `/health` повертає `"version": "3.0.0"`.


## Зміни v4

Якщо поле «Вид перевезень» неправильне або порожнє:

- XLSX однаково формується;
- фактичне значення залишається в комірці;
- перед значенням додається символ `⚠`;
- комірка виділяється світло-червоним;
- текст стає темно-червоним і жирним;
- до комірки додається примітка з допустимими значеннями.

Якщо значення правильне, оформлення шаблону не змінюється.

`/health` повертає `"version": "4.0.0"`.


## Виправлення v5

- Підсвічується лише поле `H7:J7` з неправильним видом перевезень.
- Підсвічування — тільки світло-червоний фон.
- Текст, шрифт, жирність і вирівнювання не змінюються.
- Символ `⚠` у значення більше не додається.
- Інші поля шаблону не отримують червоне оформлення.
- Перед зміною фону стиль `H7` клонуються, щоб не змінювати інші
  клітинки, які в шаблоні використовують спільний стиль.
- `/health` повертає `"version": "5.0.0"`.


---

## Endpoint заявки

### POST /fill-application

Заповнює XLSX-шаблон договору-заявки на перевезення.

Може працювати у двох режимах.

### Варіант 1: JSON без передачі шаблону

Сервіс використовує вбудований шаблон:

- `templates/zayavka_perevezennia_template.xlsx`

Request:

```json
{
  "filename": "Заявка_09.07.2026_Дніпро_Дніпро.xlsx",
  "returnJson": true,
  "application_data": {
    "application_number": "09/07",
    "application_date": "09.07.2026",
    "customer_name": "ТОВ ОВОПРАЙМ",
    "carrier_name": "ФОП ...",
    "route_from": "Дніпро",
    "route_to": "Дніпро",
    "price_text": "2500,00 грн",
    "vehicle_driver_text": "DAF ..."
  }
}
```

Якщо `returnJson: true`, відповідь:

```json
{
  "ok": true,
  "filename": "Заявка_09.07.2026_Дніпро_Дніпро.xlsx",
  "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "file_base64": "..."
}
```

Цей режим зручний для адмінки: n8n просто повертає JSON в `Respond to Webhook`, а фронт завантажує файл із `file_base64`.

### Варіант 2: multipart/form-data з власним шаблоном

Параметри:

- `template` — XLSX-файл шаблону;
- `payload` — JSON-рядок з `application_data`;
- `outputFileName` — необов’язкова назва файла;
- `returnJson` — `true`, якщо треба повернути JSON з `file_base64`.

### n8n: HTTP Request для заявки

- Method: `POST`
- URL: `https://YOUR-SERVICE.up.railway.app/fill-application`
- Header:
  - `x-api-key: <API_TOKEN>`
- Send Body: ON
- Body Content Type: JSON

Body:

```json
{
  "filename": "{{ $json.filename }}",
  "returnJson": true,
  "application_data": {{ $json.application_data }}
}
```

Після HTTP Request:

- `Respond to Webhook`
- `Respond With: JSON`
- `Response Body: {{ $json }}`

Адмінка вже підтримує `file_base64`.


## Виправлення v5.3

- У товарних рядках `A29:L33` примусово прибирається зелений фон.
- Стиль кожної комірки попередньо клонується, тому зміна не зачіпає межі, шрифт і вирівнювання інших клітинок.
- Із товарного діапазону також видаляється умовне форматування, яке могло повторно вмикати зелений фон після запису значень.
- Очищення фону виконується повторно після повного заповнення таблиці.
- `K21` записується як звичайне числове значення, щоб формат комірки застосовувався самим Excel або Google Sheets.
- `/health` повертає `"version": "5.3.0"`.


## Версія 5.4.0

- Комірка `L34` отримує підсумкову масу брутто як звичайне числове значення у кілограмах.
- Приклад: `cargo.gross_weight = "18.900"` → у `L34` записується число `18900`.
- Формат відображення залишається відповідальністю XLSX-шаблону.


## Версія 5.5.0

- Комірки `L29:L33` отримують масу брутто кожної товарної позиції як звичайне числове значення у кілограмах.
- Приклад: `cargo.items[0].gross_weight = "18.900"` → у рядок записується число `18900`.
- Для однієї позиції n8n передає загальну масу брутто в її товарний рядок.
- Для кількох позицій n8n передає окремо розпізнану масу кожного рядка.


## Версія 5.6.0

Усі значення маси брутто у таблиці ТТН записуються у тоннах:

- товарні рядки `L29:L33`;
- загальна маса `L34`.

Приклади:

- `16.100` т → числове значення `16.1`;
- `2.800` т → числове значення `2.8`;
- `18.900` т → числове значення `18.9`.

Шаблон Excel відображає значення з потрібною кількістю знаків після коми.


## Версія 5.7.0

Форматування полів:

- `A7:E7` — назва «Автомобіль» без підкреслення, вставлене значення підкреслене;
- `F7:G7` — назва «Причіп/напівпричіп» без підкреслення, вставлене значення підкреслене;
- `H7:J7` — назва «Вид перевезень» без підкреслення, вставлене значення підкреслене;
- `F19`, `G19`, `I19` — значення довжини, ширини та висоти підкреслені;
- підписи габаритів у шаблоні не змінюються.

## Версія 5.8.0

Виправлено різний формат заявки на різних комп’ютерах.

Причина була в API:
- `/fill-application` формував XLSX;
- `/fill-application-docx` формував DOCX.

Різні або закешовані версії адмінки могли викликати різні endpoint-и.

Тепер:
- `POST /fill-application` — завжди DOCX;
- `POST /fill-application-docx` — DOCX, сумісний альтернативний endpoint;
- `POST /fill-application-xlsx` — XLSX лише для явного запиту Excel.

Для DOCX відповідь завжди містить:
- розширення `.docx`;
- MIME type `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.



## Версія 5.9.0

Зміни у формуванні договору-заявки:

- поле `truck_type` містить тільки дані автомобіля, без типу/моделі/номера причепа;
- `service_period` завжди формується як перша дата та дата `+2 календарні дні`;
- `unloading_datetime` завжди отримує дату `+2 календарні дні` від першої дати періоду;
- якщо в `unloading_datetime` був час, він зберігається;
- правила застосовуються на сервері та мають пріоритет над старими значеннями із закешованої адмінки;
- `/health` повертає `"version": "5.9.0"`.


## Version 5.9.1

Application generation update:

- `transport_period` (field «СТРОК ПЕРЕВЕЗЕННЯ») is now generated by the same rule as `service_period`;
- the second date is always the first date + 2 calendar days;
- stale values sent by cached admin clients are overridden on the server.

## Version 5.9.2

Application truck type correction:

- field «ТИП АВТОМОБІЛЯ» now prioritizes only flat `truck_type` from the admin panel;
- `truck_body_type` is no longer appended to the application field;
- trailer type, body, model and plate are never added to this field;
- old combined strings are cleaned as a compatibility fallback;
- `/health` returns `"version": "5.9.2"`.

## Version 5.9.3

Root correction for application field `{{truck_type}}`:

- uses only `truck_type`;
- never falls back to `trailer_type`;
- never appends `truck_body_type`;
- never splits values by `/`;
- abbreviation `н/пр` can no longer produce the fragment `пр тентований`;
- `/health` returns `"version": "5.9.3"`.

## Version 5.9.4

TTN static-label style correction:

- `A5` — «Місце складання»;
- `A11` — «Вантажовідправник»;
- `A13` — «Вантажоодержувач»;
- `A21` — «Усього відпущено на загальну суму»;
- `A23` — «Супровідні документи на вантаж».

Railway now explicitly removes font underline and bottom border from these
static labels and restores left alignment after all TTN values are written.
The underline/bottom border of the adjacent value fields is preserved.
`/health` returns `"version": "5.9.4"`.

## Version 5.9.5

TTN XLSX formatting corrections:

- `F19`, `G19`, `I19` now use an explicit thin bottom border for
  length, width, and height values;
- the dimension fields no longer rely on font underline;
- `J21` («у тому числі ПДВ») is explicitly normalized without
  font underline or a bottom border;
- the value field `K21:L21` keeps its bottom border.

`/health` returns `"version": "5.9.5"`.

## Version 5.9.6

Dimension underline correction:

- `F19` — length;
- `G19` — width;
- `I19` — height.

Each field now receives both font underline and a thin bottom border.
This keeps the underline visible across Excel renderers.

The v5.9.5 correction for `J21` («у тому числі ПДВ») remains:
the static label is not underlined, while `K21:L21` keeps the value line.

`/health` returns `"version": "5.9.6"`.

## Version 5.9.7

Dimension field correction:

- `F19`, `G19`, and `I19` keep font underline for the numeric values;
- the thin bottom cell border is removed from all three cells;
- no long line is drawn across the full cell width;
- the `J21` VAT-label correction from v5.9.5 remains unchanged.

`/health` returns `"version": "5.9.7"`.

