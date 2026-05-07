# Manual SmartFlow Report

Исправление:

- бот больше не присылает SmartFlow-отчёт автоматически после запуска / redeploy;
- админ сам открывает отчёт, когда нужно;
- кнопка отчёта остаётся в админском меню: `📊 SmartFlow отчёт`;
- после ручного отчёта админское меню не пропадает, потому что сообщение отправляется с `admin_keyboard()`.

Что поставить в Railway:

```env
OWNER_REPORT_ENABLED=false
```

Что заменить на GitHub:

- `main.py`
- `.env.example`
- `README.md`
- `19_MANUAL_SMARTFLOW_REPORT.md`
```
