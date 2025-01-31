# dependent_trigger.sql

> **ВНИМАНИЕ:** Данный код не подлежит копированию или использованию без разрешения автора (Все права защищены).
> Данный код не предназначен для использования или распространения без письменного разрешения автора.

## Назначение

- Создает триггер `transfer_data_cascading` (AFTER UPDATE ON `Today`).
- При обновлении таблицы `Today` (например, когда меняется дата `DateReference`), 
  триггер проверяет «старые» записи и **каскадно** переносит их:
  1. Из `Today` → `ThreeDaysAgo`.
  2. Из `ThreeDaysAgo` → `OneWeekAgo`.
  3. Из `OneWeekAgo` → `TwoWeeksAgo`.
  4. Из `TwoWeeksAgo` → `OneMonthAgo`.
  5. Из `OneMonthAgo` → `PreYear`.
  6. Из `PreYear` → удаление, если дата слишком старая.

## Ключевые моменты

- Проверяет наличие записей через `IF EXISTS (SELECT 1 FROM <table> WHERE DateReference = ... ) THEN ...`.
- Удаляет записи из предыдущей таблицы, после успешной вставки в следующую.
- Оперирует датами: `CURDATE() - INTERVAL 3 DAY`, `- INTERVAL 4 DAY`, `- INTERVAL 1 WEEK`, и так далее.

**Автор: sambuka_lx**

Для связи: https://t.me/Sambuka_lx
