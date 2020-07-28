// Show initial help card loaded from card.js
$('#sendMail').off().click(function(){
    if (initialHelpCard) {
        window.amCardRenderer = new CardRenderHelper.CardRender("#actionable-message");
        window.amCardRenderer.renderCard(initialHelpCard);
    }
});