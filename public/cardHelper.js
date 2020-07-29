// Show initial help card loaded from card.js
$('#sendMail').off().click(function(){
    if (initialHelpCard) {
        window.nativeactionHandler = function() {
            
            if (nativeAction_htmlBody) {
                try {
                    Office.context.mailbox.displayNewMessageForm(
                        {
                            toRecipients: ['onboardoam@microsoft.com'],
                            subject: 'Actionable Message Issue Report for message ' + Office.context.mailbox.item.internetMessageId,
                            htmlBody: nativeAction_htmlBody
                        });
                } catch (e) {
                    if (app && app.showNotification) {
                        app.showNotification("Send mail failed", "Please copy the diagnostics content and send it to onboardoam@microsoft.com.");
                    } else {
                        window.amCardRenderer.clientManagerInstance.displaySnackMessage("Send mail failed : Please copy the diagnostics content and send it to onboardoam@microsoft.com.", false);
                    }
                }
            } else {
                if (app && app.showNotification) {
                    app.showNotification("Send mail not supported", "Please copy the diagnostics content and send it to onboardoam@microsoft.com.");
                } else {
                    window.amCardRenderer.clientManagerInstance.displaySnackMessage("Send mail failed: Please copy the diagnostics content and send it to onboardoam@microsoft.com.", false);
                }
            }
            
            window.amCardRenderer.updateActionStatus();
        }
        window.amCardRenderer = new CardRenderHelper.CardRender("#actionable-message");
        window.amCardRenderer.renderCard(initialHelpCard);
    }
});