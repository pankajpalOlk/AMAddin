var initialHelpCard = {

    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",

    "type": "AdaptiveCard",

    "version": "1.0",

    "body": [

        {

            "type": "Container",

            "id": "6343565f-52eb-c87b-c815-b772a86f1aca",

            "padding": "Default",

            "items": [

                {

                    "type": "TextBlock",

                    "id": "d9c56323-fe1a-e090-1df5-311743b856cd",

                    "text": "Actionable Messages Support Team",

                    "wrap": true

                }

            ],

            "spacing": "None",

            "style": "emphasis"

        },

        {

            "type": "Container",

            "id": "03a273b0-e511-5f0d-3a44-1eddaa03a8a5",

            "padding": "Default",

            "spacing": "None",

            "items": [

                {

                    "type": "TextBlock",

                    "id": "231d74e7-5c53-6ec8-102d-718dc936bc09",

                    "text": "What can we help you with today?",

                    "wrap": true,

                    "size": "Large",

                    "weight": "Bolder"

                }

            ]

        },

        {

            "type": "Container",

            "id": "885220a9-5ab1-95dd-5b66-20f42c452fa9",

            "padding": "Default",

            "items": [

                {

                    "actions": [

                        {

                            "method": "POST",

                            "url": "https://amsupporthack.azurewebsites.net/api/Rendering?code=dEW2WK7jALVQeBXBxVb3tg2dRPWal2Tzr9OQYbL4YCzN0wJaJ/Rshw==",

                            "title": "Rendering",

                            "type": "Action.Http",

                            "id": "3efd62b1-8fb9-b965-6fcc-a7a4491ee841"

                        },

                        {

                            "method": "POST",

                            "url": "https://amsupporthack.azurewebsites.net/api/ActionExecution?code=vitS0zijbtoiqowrATHYHSm6er/3ASJozT0GoNZTHyMfBtKcGLZyPA==",

                            "title": "Action Execution",

                            "type": "Action.Http",

                            "id": "6487a969-45cf-071f-f2a4-6052e46c16e8"

                        },

                        {

                            "type": "Action.DisplayMessageForm",

                            "itemId": "abcdef",

                            "title": "Email us for more help",

                            "isPrimary": false

                        }

                    ],

                    "id": "1097f031-1e62-fa96-9010-0ef9553f2068",

                    "type": "ActionSet"

                }

            ],

            "spacing": "None"

        }

    ],

    "padding": "None"

};