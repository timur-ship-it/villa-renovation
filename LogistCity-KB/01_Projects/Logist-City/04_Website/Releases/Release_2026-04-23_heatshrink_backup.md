---
type: site_release
release_id: rel-20260423-143551
date: 2026-04-23 14:35
git_commit: "pending"
deployed_url: "https://timur-ship-it.github.io/logist-city/"
backup_path: "/Users/Timur/Documents/New project/output/site_archives/logist-city_post_heatshrink_ui_20260423_143551.zip"
rollback_note: "Восстановить index-файлы из архива и перезалить сайт."
---

# Site Release - rel-20260423-143551

## What changed
- Подготовлен бэкап актуальной версии по термоусадке.
- В бэкап включены файлы:
  - `output/local_preview_site/index.html`
  - `docs/order-form/index.html`
  - `/tmp/logist-city/heatshrink-preview-v2.html`

## Backup integrity
- SHA256: `6d019d1d473309f9f44c42fb1bf85a96e69818e7a9b6b47112c3e42c89dd3c35`
- Size: `154274` bytes
- Created at: `2026-04-23 14:35:51`

## Rollback steps
- Распаковать архив `logist-city_post_heatshrink_ui_20260423_143551.zip`.
- Вернуть файлы `index_local_preview_site.html` и `index_docs_order_form.html` в соответствующие `index.html`.
- При необходимости взять `heatshrink-preview-v2.html` как reference-версию из архива.
