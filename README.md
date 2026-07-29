# HR Portal API v1

Окремий read-only API для порталу «Овопрайм». Він читає актуальну штатну
структуру з PostgreSQL `hr_portal`, але не може додавати, змінювати або видаляти
дані.

## Що вже реалізовано

- `GET /health` — перевірка, що HTTP-сервер API працює.
- `GET /ready` — окрема перевірка з'єднання з PostgreSQL.
- `GET /api/departments` — підрозділи, кількість посад, працівників і вакансій.
- `GET /api/employees` — список працівників із пошуком, фільтрами й пагінацією.
- `GET /api/employees/:id` — картка одного працівника.
- API-ключ у заголовку `X-API-Key`.
- CORS лише для дозволених адрес порталу.
- обмеження частоти запитів, безпечні HTTP-заголовки та приховування секретних
  заголовків у логах.
- Dockerfile і конфігурація для Railway.

У цій версії навмисно немає `POST`, `PATCH` або `DELETE`.

## Крок 1. Створити користувача бази тільки для читання

1. Відкрийте в DBeaver підключення до бази `hr_portal` під користувачем
   `postgres`.
2. Відкрийте файл `sql/01_create_hr_portal_reader.sql.txt`.
3. Згенеруйте у менеджері паролів випадковий пароль довжиною щонайменше
   32 символи.
4. Замініть у файлі
   `REPLACE_WITH_A_RANDOM_32_PLUS_CHARACTER_PASSWORD` на цей пароль.
5. Виконайте весь SQL-файл.
6. Не надсилайте пароль у чат і не зберігайте копію SQL-файлу з підставленим
   паролем.

Останні два запити у файлі перевірять результат. Для ролі мають бути вимкнені
права адміністратора, а для таблиць: `can_select = true`, `can_write = false`.

## Крок 2. Завантажити проєкт у GitHub

Створіть приватний репозиторій, наприклад `hr-portal-api`, і завантажте в нього
вміст цієї папки. Файл `.env` із секретами не створюйте в репозиторії та не
комітьте.

## Крок 3. Розгорнути API на Railway

1. У потрібному Railway-проєкті натисніть **New → GitHub Repo**.
2. Виберіть приватний репозиторій `hr-portal-api`.
3. У сервісі відкрийте **Variables** і додайте:

```text
NODE_ENV=production
DATABASE_URL=postgresql://hr_portal_reader:YOUR_PASSWORD@shortline.proxy.rlwy.net:48195/hr_portal
DATABASE_SSL=true
DATABASE_SSL_REJECT_UNAUTHORIZED=false
DATABASE_POOL_MAX=5
API_ACCESS_TOKEN=YOUR_RANDOM_32_PLUS_CHARACTER_TOKEN
ALLOWED_ORIGINS=https://hr-portal-yaico.bridge182.chatgpt.site
LOG_LEVEL=info
```

Для `DATABASE_URL` використайте пароль `hr_portal_reader`. Якщо пароль містить
символи на кшталт `@`, `:`, `/`, `?` або `#`, їх треба URL-кодувати. Безпечний
простий варіант — згенерувати довгий пароль з літер і цифр.

`API_ACCESS_TOKEN` — інший незалежний випадковий секрет. Не використовуйте для
нього пароль бази.

Railway сам задає змінну `PORT`, тому вручну її додавати не треба.

## Крок 4. Додати адресу API і перевірити

Після успішного розгортання у Railway відкрийте **Settings → Networking →
Generate Domain**.

Перевірка, що API запущений:

```text
https://YOUR-API-DOMAIN/health
```

Очікувана відповідь:

```json
{
  "status": "ok"
}
```

Перевірка PostgreSQL:

```text
https://YOUR-API-DOMAIN/ready
```

Очікується `status: "ready"`, `database: "connected"`,
`databaseName: "hr_portal"` і `databaseUser: "hr_portal_reader"`.

Перевірка працівників у PowerShell:

```powershell
$headers = @{ "X-API-Key" = "YOUR_RANDOM_API_ACCESS_TOKEN" }
Invoke-RestMethod `
  -Uri "https://YOUR-API-DOMAIN/api/employees?limit=5" `
  -Headers $headers
```

У полі `pagination.total` для поточного імпорту очікується `65`.

## Параметри списку працівників

```text
GET /api/employees?query=Ковальова
GET /api/employees?department=Бухгалтерія
GET /api/employees?status=active
GET /api/employees?limit=25&offset=0
```

Параметри можна поєднувати. Максимальний `limit` — 100.

## Важливі правила безпеки

- API підключається лише під `hr_portal_reader`, не під `postgres` і не під
  власником бази `hr_portal_app`.
- Паролі співробітників, RDP, BitLocker і ліцензійні ключі цей API не читає й
  не зберігає.
- `API_ACCESS_TOKEN` не можна вставляти у JavaScript статичної вебсторінки:
  відвідувач браузера зможе його побачити. На наступному етапі портал
  звертатиметься до API через серверний проксі або власну серверну частину.
- Для майбутнього редагування даних буде окрема роль і окремі перевірені
  маршрути з журналом аудиту.

## Локальна перевірка

Потрібні Node.js 22+ і власний файл `.env` зі значеннями за прикладом
`.env.example`.

```bash
npm ci
npm run typecheck
npm run build
npm start
```

Для перенесення на власний хостинг достатньо Docker-сумісного середовища та тих
самих змінних оточення.
