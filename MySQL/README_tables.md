# tables.sql

> **ВНИМАНИЕ:** Данный код не подлежит копированию или использованию без разрешения автора (Все права защищены).
> Данный код не предназначен для использования или распространения без письменного разрешения автора.

## Назначение

1. Создает таблицы: `users`, `notes`, `Today`, `ThreeDaysAgo`, `OneWeekAgo`, `TwoWeeksAgo`, `OneMonthAgo`, `PreYear`, `AllTime`, `debug_table`, `password`.
2. Добавляет тестовые данные в таблицу `users`.
3. Настраивает внешние ключи (например, связь `notes.id` → `users.id_user`).

## Основные моменты

- **users**: таблица с перечнем пользователей (id_user, name).
- **notes**: хранит «заметки» (id_note, id, name, num_sum_read, num_sum_send, …).
- **Today**, **ThreeDaysAgo**, **OneWeekAgo**, … : сводные таблицы, в которых агрегируются данные за разные периоды.
- **AllTime**: общая таблица с накопительными данными без ограничения по дате.
- **debug_table**: вспомогательная таблица для хранения отладочных сообщений.
- **password**: таблица, в которой может храниться имя и пароль (строго для внутреннего использования).

**Автор: sambuka_lx**

Для связи: https://t.me/Sambuka_lx
