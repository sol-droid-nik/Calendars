# Work Calendars (GitHub Pages)

Шаги:
1) Создай репозиторий на GitHub (например, `work-calendars`) и включи GitHub Pages: Settings → Pages → Source: **GitHub Actions**.
2) Загрузите файлы из этого архива (кнопка **Add file → Upload files**).
3) Помести Excel `Työvuorot vuosi 2025.xlsx` в папку `data/` (как у начальницы: вкладки vko 1, vko 2, ... с заголовками `MA 10.3`, `TI 11.3`, ...).
4) Подожди 1–3 минуты — GitHub Actions соберёт сайт и опубликует его.
5) Ссылка на сайт будет вида: `https://ТВОЁ_ИМЯ.github.io/ИМЯ_РЕПО/` → там список персональных ссылок на .ics.
6) Сотрудники подписываются на свой URL (iPhone: Settings → Calendar → Accounts → Add Subscribed Calendar; Google Calendar: web → Other calendars → From URL).

По умолчанию делаем календари на ближайшие 8 недель. Если хочешь изменить — редактируй `scripts/build_calendars.py`.

Дата сборки архива: 2025-11-01T02:37:42Z
