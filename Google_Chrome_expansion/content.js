chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "getExtractedData") {
        const companyName = document.querySelector('h3.fw-semi-bold')?.innerText || 'Не найдено';
        const inn = document.querySelector('a[data-inn]')?.getAttribute('data-inn') || 'Не найдено';
        const director = document.querySelector('div[label="Руководитель"] + .card-value span')?.innerText || 'Не найдено';

        let phones = Array.from(document.querySelectorAll('.label-phone span.text-nowrap')).slice(0, 5).map(el => el.innerText.trim());
        let emails = Array.from(document.querySelectorAll('#emails a[href^="mailto:"]')).slice(0, 5).map(el => el.innerText.trim());

        sendResponse({
            companyName,
            inn,
            director,
            phones,
            emails
        });
    }
});
