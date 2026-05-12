# Apps Script — серверная часть P0 фиксов

Этот сниппет добавляет в твой Google Apps Script:
1. **Серверную пере-валидацию теста** — клиент больше не верится на слово.
2. **Cloudflare Turnstile siteverify** — отшивает ботов и DevTools-подделки без валидного токена.
3. **Сравнение клиентского score с серверным** — флаг при расхождении (попытка обмана).

---

## Шаг 1. Получи Turnstile secret key

1. Заходишь на https://dash.cloudflare.com → Turnstile.
2. Создаёшь сайт: name «HC онбординг», hostnames — `albertkaunitz.github.io`.
3. Получаешь **Site key** (публичный, идёт в index.html) и **Secret key** (приватный, остаётся в Apps Script).

---

## Шаг 2. В Apps Script → Project Settings → Script Properties добавь:

| Property | Value |
|---|---|
| `TURNSTILE_SECRET` | твой Secret key с Cloudflare |
| `ANSWERS_KEY` | (см. шаг 3 — JSON с правильными ответами) |

---

## Шаг 3. Сформируй ANSWERS_KEY

Скопируй массив `BLOCKS` из старого index.html (до фикса). Преврати в плоскую структуру:

```json
{
  "0_0": [2],
  "0_1": [1],
  "0_2": [0, 3]
}
```

Где ключ — `{blockIdx}_{questionIdx}`, значение — массив правильных индексов опций (как в `q.correct`).

Сохрани этот JSON в Script Property `ANSWERS_KEY` строкой.

---

## Шаг 4. Замени `doPost` в Apps Script на следующий код

```javascript
function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  const TURNSTILE_SECRET = props.getProperty('TURNSTILE_SECRET');
  const ANSWERS_KEY = JSON.parse(props.getProperty('ANSWERS_KEY') || '{}');

  const p = e.parameter;

  // ═══ 1. ПРОВЕРКА TURNSTILE ═══
  if (TURNSTILE_SECRET && p.cf_token) {
    const verifyResp = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      method: 'post',
      payload: { secret: TURNSTILE_SECRET, response: p.cf_token },
      muteHttpExceptions: true
    });
    const verifyData = JSON.parse(verifyResp.getContentText());
    if (!verifyData.success) {
      return jsonResponse({ ok: false, error: 'turnstile_failed', codes: verifyData['error-codes'] || [] });
    }
  } else if (TURNSTILE_SECRET) {
    return jsonResponse({ ok: false, error: 'turnstile_missing' });
  }

  // ═══ 2. ПЕРЕ-ВАЛИДАЦИЯ ОТВЕТОВ (только для attestation_summary) ═══
  let serverPercent = null, serverPass = null, mismatch = false;
  if (p.type === 'attestation_summary' && p.raw_answers) {
    try {
      const rawAnswers = JSON.parse(p.raw_answers);
      let correct = 0, total = 0;
      Object.keys(ANSWERS_KEY).forEach(key => {
        total++;
        const userAns = (rawAnswers[key] || []).slice().sort().join(',');
        const rightAns = ANSWERS_KEY[key].slice().sort().join(',');
        if (userAns === rightAns) correct++;
      });
      serverPercent = total > 0 ? Math.round((correct / total) * 100) : 0;
      serverPass = serverPercent >= 90;

      const clientPct = parseInt(p.percent || '0', 10);
      if (Math.abs(clientPct - serverPercent) > 1) mismatch = true;
    } catch(err) { mismatch = true; }
  }

  // ═══ 3. ЗАПИСЬ В SHEET ═══
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.type === 'portfolio' ? 'Портфолио' : 'Аттестация');
  if (!sheet) return jsonResponse({ ok: false, error: 'sheet_not_found' });

  const row = [
    new Date(), p.attempt_id || '', p.fio || '', p.telegram || '', p.email || '',
    p.status || '', p.contract || '', p.percent || '', p.passfail || '',
    serverPercent !== null ? serverPercent : '',
    serverPass !== null ? (serverPass ? 'PASS' : 'FAIL') : '',
    mismatch ? '⚠️ MISMATCH' : '',
    p.block1 || '', p.block2 || '', p.block3 || '', p.block4 || '',
    p.timeMin || '', p.gdpr || '', p.error_count || '', p.type || ''
  ];
  sheet.appendRow(row);

  return jsonResponse({
    ok: true,
    server_percent: serverPercent,
    server_passfail: serverPass !== null ? (serverPass ? 'PASS' : 'FAIL') : null,
    mismatch: mismatch
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
```

---

## Шаг 5. Деплой

1. Apps Script → Deploy → Manage Deployments → New version.
2. URL может остаться прежним при обновлении той же версии.
3. Убедись, что `WEBAPP_URL` в `index.html` (строки 1314, 1973) совпадает.

---

## Шаг 6. (Опционально) Активация серверного scoring в клиенте

После шага 5 в `index.html` строка ~1991:

```javascript
const USE_SERVER_SCORING = false;  // меняешь на true
```

**Важно:** даже с `false` сервер уже пишет реальный score и ставит `⚠️ MISMATCH` при попытке обмана через DevTools. То есть защита работает с момента деплоя Apps Script, даже без флипа флага.

> Полная реализация `USE_SERVER_SCORING=true` (чтение ответа от сервера через `fetch` вместо iframe-form) — задача следующего захода.

---

## Проверка работы

1. Открой `https://albertkaunitz.github.io/hc-test/` в режиме инкогнито.
2. Пройди тест нормально → должно дойти в таблицу, `server_percent = client percent`, `mismatch = пусто`.
3. Попробуй через DevTools поправить ответы → должен прийти `⚠️ MISMATCH`.
4. Попробуй сабмитнуть прямым `curl` без `cf_token` → должен прийти отказ `turnstile_missing`.

---

## Что НЕ покрыто этим сниппетом

- `submitPortfolio` (`type=portfolio`) — пока без Turnstile (отдельная форма). Добавим если начнётся флуд.
- Rate-limit на уровне Apps Script — не реализован (можно через `CacheService`).
- Бэкап `ANSWERS_KEY` — держи копию JSON в безопасном месте.
