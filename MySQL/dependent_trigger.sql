-- Все права защищены (c) 2024
-- Данный скрипт создает триггер transfer_data_cascading, который переносит записи
-- из одной таблицы в другую при достижении определенной «возрастной» метки по дате.
-- Код не подлежит копированию и распространению без согласия автора.

DELIMITER //
CREATE TRIGGER transfer_data_cascading
AFTER UPDATE ON Today 
FOR EACH ROW
BEGIN
    IF EXISTS (SELECT 1 FROM Today WHERE DateReference = CURDATE() - INTERVAL 3 DAY) THEN
        INSERT INTO ThreeDaysAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
        SELECT id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference
        FROM Today
        WHERE DateReference = CURDATE() - INTERVAL 4 DAY;

        DELETE FROM Today WHERE DateReference = CURDATE() - INTERVAL 4 DAY;
    END IF;

    IF EXISTS (SELECT 1 FROM ThreeDaysAgo WHERE DateReference = CURDATE() - INTERVAL 4 DAY) THEN
        INSERT INTO OneWeekAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
        SELECT id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference
        FROM ThreeDaysAgo
        WHERE DateReference = CURDATE() - INTERVAL 4 DAY;

        DELETE FROM ThreeDaysAgo WHERE DateReference = CURDATE() - INTERVAL 4 DAY;
    END IF;
    
    IF EXISTS (SELECT 1 FROM OneWeekAgo WHERE DateReference = CURDATE() - INTERVAL 1 WEEK) THEN
        INSERT INTO TwoWeeksAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
        SELECT id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference
        FROM OneWeekAgo
        WHERE DateReference = CURDATE() - INTERVAL 1 WEEK;

        DELETE FROM OneWeekAgo WHERE DateReference = CURDATE() - INTERVAL 1 WEEK;
    END IF;

    IF EXISTS (SELECT 1 FROM TwoWeeksAgo WHERE DateReference = CURDATE() - INTERVAL 16 DAY) THEN
        INSERT INTO OneMonthAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
        SELECT id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference
        FROM TwoWeeksAgo
        WHERE DateReference = CURDATE() - INTERVAL 16 DAY;

        DELETE FROM TwoWeeksAgo WHERE DateReference = CURDATE() - INTERVAL 16 DAY;
    END IF;



    IF EXISTS (SELECT 1 FROM OneMonthAgo WHERE DateReference <= CURDATE() - INTERVAL 10 MONTH) THEN
        INSERT INTO PreYear (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
        SELECT id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference
        FROM OneMonthAgo
        WHERE DateReference <= CURDATE() - INTERVAL 10 MONTH;

        DELETE FROM OneMonthAgo WHERE DateReference <= CURDATE() - INTERVAL 10 MONTH;
    END IF;


    IF EXISTS (SELECT 1 FROM PreYear WHERE DateReference <= CURDATE() - INTERVAL 1 YEAR) THEN
        DELETE FROM PreYear WHERE DateReference <= CURDATE() - INTERVAL 1 YEAR;
    END IF;

END;
//
DELIMITER ;
