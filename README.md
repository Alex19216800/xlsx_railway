# TTN XLSX Service v3

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
