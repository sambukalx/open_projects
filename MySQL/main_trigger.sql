-- Все права защищены (c) 2024
-- Данный скрипт создает триггер main_trigger, который срабатывает при добавлении записей в notes
-- и распределяет данные по различным сводным таблицам в зависимости от даты.
-- Код не подлежит копированию и распространению без согласия автора.

DELIMITER //

CREATE TRIGGER main_trigger 
AFTER INSERT ON notes
FOR EACH ROW
BEGIN
    DECLARE TimeRead DOUBLE;
    DECLARE TimeSend DOUBLE;

    SET TimeRead = NEW.num_sum_read * 0.12 / 60;
    SET TimeSend = NEW.num_sum_send * 0.4 / 60;


    SET @start_of_week := CURDATE() - INTERVAL WEEKDAY(CURDATE()) DAY;
    SET @end_of_week := @start_of_week + INTERVAL 6 DAY;


    SET @start_of_last_two_week := SUBDATE(CURDATE(), INTERVAL WEEKDAY(CURDATE()) + 14 DAY);
    SET @end_of_last_two_week := SUBDATE(@start_of_last_two_week, INTERVAL 1 DAY);

    SET @start_of_month := CURDATE() - INTERVAL DAY(CURDATE()) - 1 DAY;
    SET @end_of_month := LAST_DAY(CURDATE());

    SET @start_of_year := MAKEDATE(YEAR(CURDATE()), 1);
    SET @end_of_year := MAKEDATE(YEAR(CURDATE()) + 1, 1) - INTERVAL 1 DAY;


    IF DATE(NEW.created_at) = CURDATE() THEN
        IF EXISTS (SELECT * FROM Today WHERE id = NEW.id) THEN
            UPDATE Today
            SET ReadSum = ReadSum + NEW.num_sum_read,
                SendSum = SendSum + NEW.num_sum_send,
                TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
                TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60),
                DateReference = NEW.created_at
            WHERE id = NEW.id;
        ELSE
            INSERT INTO Today (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
            VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
        END IF;
    END IF;

    IF DATEDIFF(CURDATE(), DATE(NEW.created_at)) <= 3 THEN
        IF EXISTS (SELECT * FROM ThreeDaysAgo WHERE id = NEW.id) THEN
            UPDATE ThreeDaysAgo
            SET ReadSum = ReadSum + NEW.num_sum_read,
                SendSum = SendSum + NEW.num_sum_send,
                TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
                TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60),
                DateReference = NEW.created_at
            WHERE id = NEW.id;
        ELSE
            INSERT INTO ThreeDaysAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
            VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
        END IF;
    END IF;


    IF NEW.created_at >= @start_of_week AND NEW.created_at <= @end_of_week THEN
        IF EXISTS (SELECT * FROM OneWeekAgo WHERE id = NEW.id AND DateReference BETWEEN @start_of_week AND @end_of_week) THEN
            UPDATE OneWeekAgo
            SET ReadSum = ReadSum + NEW.num_sum_read,
                SendSum = SendSum + NEW.num_sum_send,
                TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
                TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60),
                DateReference = NEW.created_at
            WHERE id = NEW.id AND DateReference BETWEEN @start_of_week AND @end_of_week;
        ELSE
            INSERT INTO OneWeekAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
            VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
        END IF;
    END IF;


    IF NEW.created_at >= @start_of_last_two_week AND NEW.created_at <= @end_of_last_two_weekk THEN
        IF EXISTS (SELECT * FROM TwoWeeksAgo WHERE id = NEW.id) THEN
            UPDATE TwoWeeksAgo
            SET ReadSum = ReadSum + NEW.num_sum_read,
                SendSum = SendSum + NEW.num_sum_send,
                TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
                TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60),
                DateReference = NEW.created_at
            WHERE id = NEW.id;
        ELSE
            INSERT INTO TwoWeeksAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
            VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
        END IF;
    END IF;



    IF NEW.created_at >= @start_of_month AND NEW.created_at <= @end_of_month THEN
        IF EXISTS (SELECT * FROM OneMonthAgo WHERE id = NEW.id AND DateReference BETWEEN @start_of_month AND @end_of_month) THEN
            UPDATE OneMonthAgo
            SET ReadSum = ReadSum + NEW.num_sum_read,
                SendSum = SendSum + NEW.num_sum_send,
                TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
                TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60),
                DateReference = NEW.created_at
            WHERE id = NEW.id AND DateReference BETWEEN @start_of_month AND @end_of_month;
        ELSE
            INSERT INTO OneMonthAgo (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
            VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
        END IF;
    END IF;


    IF NEW.created_at >= @start_of_year AND NEW.created_at <= @end_of_year THEN
        IF EXISTS (SELECT * FROM PreYear WHERE id = NEW.id AND DateReference BETWEEN @start_of_year AND @end_of_year) THEN
            UPDATE PreYear
            SET ReadSum = ReadSum + NEW.num_sum_read,
                SendSum = SendSum + NEW.num_sum_send,
                TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
                TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60),
                DateReference = NEW.created_at
            WHERE id = NEW.id AND DateReference BETWEEN @start_of_year AND @end_of_year;
        ELSE
            INSERT INTO PreYear (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
            VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
        END IF;
    END IF;


    IF EXISTS (SELECT * FROM AllTime WHERE id = NEW.id) THEN
        UPDATE AllTime
        SET ReadSum = ReadSum + NEW.num_sum_read,
            SendSum = SendSum + NEW.num_sum_send,
            TimeRead = TimeRead + (NEW.num_sum_read * 0.12 / 60),
            TimeSend = TimeSend + (NEW.num_sum_send * 0.4 / 60)
        WHERE id = NEW.id;
    ELSE
        INSERT INTO AllTime (id, name, ReadSum, SendSum, TimeRead, TimeSend, DateReference)
        VALUES (NEW.id, NEW.name, NEW.num_sum_read, NEW.num_sum_send, TimeRead, TimeSend, NEW.created_at);
    END IF;

END;
//
DELIMITER ;
