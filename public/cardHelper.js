// Show initial help card loaded from card.js
$('#sendMail').off().click(function(){
    if (initialHelpCard) {
        window.nativeactionHandler = function() {
            console.log('native action handler called');
            window.amCardRenderer.updateActionStatus();
        }
        window.amCardRenderer = new CardRenderHelper.CardRender("#actionable-message");
        window.amCardRenderer.renderCard(initialHelpCard);
    }
});