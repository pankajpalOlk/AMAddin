var initialHelpCard = {
    "type": "AdaptiveCard",
    "id": "427FB9DB-E2E6-4A26-B684-454D6B62A731",
    "correlationId": "9b7aedb0-833f-4bde-b9b4-31f69ca14dc0",
    "originator": "0eb3a855-e2d4-4bc9-8038-b22d614e4788",
    "version": "1.0",
    "padding": "Default",
    "body": [
        {
            "text": "What can we help you with?",
            "wrap": true,
            "id": "3095c4b4-03a5-dd07-7915-98d7c7fef7a9",
            "type": "TextBlock"
        },
        {
            "actions": [
                {
                    "method": "POST",
                    "url": "https://amsupporthack.azurewebsites.net/api/PartnerOnboarding?code=9d97sV8aptFv/CuPd7iObLvG52jhAUFNXT4allUzmXBkNmwW3uQI7w==",
                    "title": "Partner Onboarding",
                    "isPrimary": false,
                    "type": "Action.Http",
                    "id": "4d1029b7-2d80-8088-8f6f-48532e1b286f"
                },
                {
                    "method": "POST",
                    "url": "https://amsupporthack.azurewebsites.net/api/Rendering?code=dEW2WK7jALVQeBXBxVb3tg2dRPWal2Tzr9OQYbL4YCzN0wJaJ/Rshw==",
                    "title": "Rendering",
                    "isPrimary": false,
                    "type": "Action.Http",
                    "id": "3efd62b1-8fb9-b965-6fcc-a7a4491ee841"
                },
                {
                    "method": "POST",
                    "url": "https://amsupporthack.azurewebsites.net/api/ActionExecution?code=vitS0zijbtoiqowrATHYHSm6er/3ASJozT0GoNZTHyMfBtKcGLZyPA==",
                    "title": "Action Execution",
                    "isPrimary": false,
                    "type": "Action.Http",
                    "id": "6487a969-45cf-071f-f2a4-6052e46c16e8"
                },
                {
                    "method": "POST",
                    "url": "https://amsupporthack.azurewebsites.net/api/Others?code=QRdK6GiHRrv/96q0IOKLlboan89spKuS7P0TIjTIT0KO/ssQbZFx8Q==",
                    "title": "Others",
                    "isPrimary": false,
                    "type": "Action.Http",
                    "id": "6487a969-45cf-071f-f2a4-6052e46c16e8"
                }
            ],
            "id": "1097f031-1e62-fa96-9010-0ef9553f2068",
            "type": "ActionSet"
        }
    ],
    "autoInvokeAction": null,
    "autoInvokeOptions": null,
    "constrainWidth": true,
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "@type": "AdaptiveCard",
    "@context": "http://schema.org/extensions"
};