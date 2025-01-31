const bitrixWebhookUrl = '';
const userWebhookUrl = '';
const userGetUrl = '';

const maxDuplicateSubmissions = 3;
const blockTimeMinutes = 30;
const correctPassword = "";
const maxAttempts = 3;
let failedAttempts = parseInt(localStorage.getItem("failedAttempts")) || 0;

// Получение списка всех сотрудников и сохранение в localStorage
async function fetchAllUsers() {
    let allUsers = [];
    let start = 0;

    try {
        while (true) {
            const response = await fetch(`${userGetUrl}?start=${start}`);
            const result = await response.json();

            if (result.result) {
                allUsers = allUsers.concat(result.result);
                console.log(`Загружено ${allUsers.length} пользователей`);
            } else {
                console.error("Ошибка загрузки пользователей: ", result.error_description || "Неизвестная ошибка");
                break;
            }

            if (result.total && result.result.length < 50) {
                break;
            }

            start += 50;
        }

        localStorage.setItem("bitrix_users", JSON.stringify(allUsers));
        console.log("Все пользователи успешно загружены и сохранены в localStorage.");
    } catch (error) {
        console.error("Ошибка при получении списка пользователей:", error);
    }
}

// Получение ФИО пользователя по ID
function getUserNameById(userId) {
    const users = JSON.parse(localStorage.getItem("bitrix_users")) || [];
    console.log("Список пользователей:", users);
    console.log("Ищем пользователя с ID:", userId);
    const user = users.find(user => user.ID == userId);
    return user ? `${user.NAME} ${user.LAST_NAME}` : "Неизвестный пользователь";
}

// Загрузка и сохранение ID ответственного пользователя
function loadUserId() {
    let userId = localStorage.getItem("bitrix_user_id");

    if (!userId) {
        userId = prompt("Введите ID ответственного пользователя:");
        if (userId) {
            localStorage.setItem("bitrix_user_id", userId);
        } else {
            alert("Необходимо ввести ID!");
            return;
        }
    }

    const userName = getUserNameById(userId);
    document.getElementById("userId").textContent = `${userId} (${userName})`;
}

function disableButtonTemporarily() {
    const button = document.getElementById("sendToBitrix");
    button.disabled = true;
    button.textContent = "Подождите...";
    setTimeout(() => {
        button.disabled = false;
        button.textContent = "Отправить в Битрикс24";
    }, 10000);
}

function checkDuplicateSubmission(data) {
    const lastSubmission = JSON.parse(localStorage.getItem("last_submission")) || {};
    const submissionCount = parseInt(localStorage.getItem("submission_count")) || 0;

    if (JSON.stringify(lastSubmission) === JSON.stringify(data)) {
        if (submissionCount + 1 >= maxDuplicateSubmissions) {
            alert(`Вы отправили одинаковые данные 3 раза. Расширение заблокировано на ${blockTimeMinutes} минут.`);
            const blockUntil = Date.now() + blockTimeMinutes * 60 * 1000;
            localStorage.setItem("blockedUntil", blockUntil);
            return false;
        }
        localStorage.setItem("submission_count", submissionCount + 1);
    } else {
        localStorage.setItem("last_submission", JSON.stringify(data));
        localStorage.setItem("submission_count", 1);
    }

    return true;
}

function isBlocked() {
    const blockedUntil = localStorage.getItem("blockedUntil");
    if (blockedUntil && Date.now() < parseInt(blockedUntil)) {
        alert("Расширение заблокировано. Попробуйте позже.");
        return true;
    }
    return false;
}

async function sendDataToBitrix(data) {
    if (isBlocked()) return;

    if (!checkDuplicateSubmission(data)) return;

    disableButtonTemporarily();
    
    const userId = localStorage.getItem("bitrix_user_id");
    if (!userId) {
        alert("Ошибка: ID пользователя не установлен!");
        return;
    }

    const userName = getUserNameById(userId);

    const payload = {
        fields: {
            TITLE: `${data.companyName} GOSBASE ГОСБАЗА`,
            UF_CRM_INN: data.inn,
            COMMENTS: `Руководитель: ${data.director}`,
            SOURCE_DESCRIPTION: `Отправитель: ${userName}`,
            COMPANY_TITLE: data.companyName,
            PHONE: data.phones.map(phone => ({ VALUE: phone, VALUE_TYPE: "WORK" })),
            EMAIL: data.emails.map(email => ({ VALUE: email, VALUE_TYPE: "WORK" })),
            SOURCE_ID: "UC_GKMJY7",
            STATUS_ID: "NEW",
            ASSIGNED_BY_ID: userId
        }
    };

    console.log("Отправка лида в Битрикс24 с данными:", payload);

    try {
        const response = await fetch(bitrixWebhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const result = await response.json();
        if (result.result) {
            alert("Лид успешно создан!");

            // Показываем временное сообщение и запускаем таймер
            document.getElementById("leadResponsible").textContent = "Ожидание...";
            setTimeout(() => {
                checkLeadResponsible(result.result);
            }, 10000);
        } else {
            alert("Ошибка при создании: " + (result.error_description || "Неизвестная ошибка"));
        }
    } catch (error) {
        alert("Ошибка сети: " + error.message);
    }
}

// Функция для получения ответственного за лид по его ID
async function checkLeadResponsible(leadId) {
    try {
        const response = await fetch(``);
        const result = await response.json();

        if (result.result && result.result.ASSIGNED_BY_ID) {
            const responsibleName = getUserNameById(result.result.ASSIGNED_BY_ID);
            document.getElementById("leadResponsible").textContent = responsibleName || "Неизвестно";
            console.log(`Ответственный за лид ${leadId}: ${responsibleName}`);
        } else {
            document.getElementById("leadResponsible").textContent = "Ошибка получения данных";
        }
    } catch (error) {
        document.getElementById("leadResponsible").textContent = "Ошибка сети";
        console.error("Ошибка при получении ответственного:", error);
    }
}

// Обновление ID пользователя с проверкой пароля
function updateUserId() {
    if (isBlocked()) {
        alert("Вы превысили количество попыток. Попробуйте позже.");
        return;
    }

    const enteredPassword = document.getElementById("passwordInput").value;
    if (enteredPassword !== correctPassword) {
        failedAttempts++;
        localStorage.setItem("failedAttempts", failedAttempts);

        if (failedAttempts >= maxAttempts) {
            alert("Вы превысили количество попыток! Блокировка на 1 час.");
            const blockTime = Date.now() + 60 * 60 * 1000;
            localStorage.setItem("blockedUntil", blockTime);
            isBlocked = true;
        } else {
            alert(`Неверный пароль. Осталось попыток: ${maxAttempts - failedAttempts}`);
        }
        return;
    }

    failedAttempts = 0;
    localStorage.setItem("failedAttempts", failedAttempts);

    const newUserId = document.getElementById("newUserId").value;
    if (newUserId) {
        localStorage.setItem("bitrix_user_id", newUserId);
        document.getElementById("userId").textContent = `${newUserId} (${getUserNameById(newUserId)})`;
        alert("ID пользователя обновлен!");
    } else {
        alert("Введите новый ID!");
    }
}

// Открытие настроек
function openSettings() {
    document.getElementById("settings").style.display = "block";
}

document.addEventListener("DOMContentLoaded", async function () {
    await fetchAllUsers();
    loadUserId();

    chrome.tabs.query({ active: true, currentWindow: true }, function(tabs) {
        chrome.tabs.sendMessage(tabs[0].id, { action: "getExtractedData" }, (response) => {
            if (chrome.runtime.lastError || !response) {
                document.getElementById("companyName").textContent = "Ошибка загрузки данных";
                return;
            }

            document.getElementById("companyName").textContent = response.companyName;
            document.getElementById("inn").textContent = response.inn;
            document.getElementById("director").textContent = response.director;
            document.getElementById("phones").textContent = response.phones.join(', ');
            document.getElementById("emails").textContent = response.emails.join(', ');

            document.getElementById("sendToBitrix").addEventListener("click", function() {
                sendDataToBitrix(response);
            });
        });
    });

    document.getElementById("updateUserId").addEventListener("click", updateUserId);
    document.getElementById("openSettings").addEventListener("click", openSettings);
});
