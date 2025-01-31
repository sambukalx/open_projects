-- Все права защищены (c) 2024
-- Данный скрипт создает таблицы и первичные данные для нашего MySQL-проекта.
-- Код не подлежит копированию и распространению без согласия автора.

CREATE TABLE users (
    id_user INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(40) NOT NULL
);

CREATE TABLE notes (
    id_note INT AUTO_INCREMENT PRIMARY KEY,
    id INT,
    name VARCHAR(40) NOT NULL,
    num_sum_read INT,
    num_sum_send INT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, 
    FOREIGN KEY (id) REFERENCES users(id_user)
);

INSERT INTO users (name) VALUES
    ('Lera'),
    ('Kira'),
    ('Yana'),
    ('Lena'),
    ('Natasha'),
    ('Nasiba'),
    ('Test');

CREATE TABLE Today (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    PRIMARY KEY (id, DateReference), 
    FOREIGN KEY (id) REFERENCES users(id_user)
);

CREATE TABLE ThreeDaysAgo (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    TimeEN DATE,
    PRIMARY KEY (id, DateReference),
    FOREIGN KEY (id) REFERENCES users(id_user)
);

CREATE TABLE OneWeekAgo (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    TimeEN DATE,
    PRIMARY KEY (id, DateReference),
    FOREIGN KEY (id) REFERENCES users(id_user)
);

CREATE TABLE TwoWeeksAgo (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    TimeEN DATE,
    PRIMARY KEY (id, DateReference),
    FOREIGN KEY (id) REFERENCES users(id_user)
);

CREATE TABLE OneMonthAgo (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    TimeEN DATE,
    PRIMARY KEY (id, DateReference),
    FOREIGN KEY (id) REFERENCES users(id_user)
);

CREATE TABLE PreYear (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    TimeEN DATE,
    PRIMARY KEY (id, DateReference),
    FOREIGN KEY (id) REFERENCES users(id_user)
);


CREATE TABLE AllTime (
    id INT,
    name VARCHAR(40),
    ReadSum FLOAT,
    SendSum INT,
    TimeRead DECIMAL(10, 2),
    TimeSend DECIMAL(10, 2),
    DateReference TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    TimeST DATE,
    TimeEN DATE,
    PRIMARY KEY (id, DateReference),
    FOREIGN KEY (id) REFERENCES users(id_user)
);

CREATE TABLE debug_table (
    debug_message VARCHAR(255)
);

CREATE TABLE password (
    name VARCHAR(40) NOT NULL,
    pas VARCHAR(40) NOT NULL
);
