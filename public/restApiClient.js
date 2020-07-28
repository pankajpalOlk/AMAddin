
function RestApiClient(restUrl, user) {
    if (restUrl) {
        this.restUrl = restUrl;
    } else {
        this.restUrl = "https://outlook.office.com/api";
    }

    this.user = user;
}

RestApiClient.prototype.loadProperties = function (itemId, token, callback) {

    var self = this;

    var odataId = self.ewsIdToDataId(itemId);

    var restApiUrl = self.restUrl + '/v2.0/me/messages/' + odataId + '/';
    restApiUrl += "?$expand = SingleValueExtendedProperties($filter = "
        + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name EntityDocument')"
        + " or "
        + "(PropertyId eq 'Boolean {00062008-0000-0000-c000-000000000046} Name EntityExtractionSuccess')"
        + " or "
        + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name EntityExtractionServiceDiagnosticContext')"
        + " or "
        + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name ExplicitMessageCard')"
        + " or "
        + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name ActionExecutionHttpTrace')"
        + " or "
        + "(PropertyId eq 'String {00062008-0000-0000-c000-000000000046} Name TeeVersion')"
        + " or "
        + "(PropertyId eq 'String 0x007D')"
        + "),Extensions($filter=(Id eq 'Com.Microsoft.Graph.MessageCard'))";

    $.ajax({
        type: "GET",
        url: restApiUrl,
        dataType: 'json',
        headers: {
            "Authorization": "Bearer " + token,
            "Accept": "application/json; odata.metadata=none",
            "Prefer": "outlook.allow-unsafe-html"
        },
        cache: false,
        error: function (jqXHR, textStatus, errorThrown) {
            var errorMsg;
            if (textStatus === "error" && jqXHR.status === 0) {
                // Mac Outlook usually ends in this status
                errorMsg = "0"
            }else{
                errorMsg = "textStatus: " + textStatus + '\nerrorThrown: ' + errorThrown + "\nState: " + jqXHR.state() + "\njqXHR: " + JSON.stringify(jqXHR, null, 2);
            }

            callback(null, null, null, null, null, null, null, null, null, errorMsg);
        },
        success: function (odata) {
            var messageCard = null;
            var adaptiveCard = null;
            var diagnostics = null;

            if (odata.Extensions && odata.Extensions.length > 0) {

                var cardExt = odata.Extensions[0];
                if (cardExt.MessageCardSerialized) {
                    messageCard = $.parseJSON(cardExt.MessageCardSerialized);
                }

                if (cardExt.AdaptiveCardSerialized) {
                    adaptiveCard = $.parseJSON(cardExt.AdaptiveCardSerialized);
                }

                if (cardExt.DeveloperDiagnosticsSerialized) {
                    diagnostics = $.parseJSON(cardExt.DeveloperDiagnosticsSerialized);
                }
            }

            var svp = odata.SingleValueExtendedProperties;
            if (svp && svp.length > 0) {
                for (var prop in svp) {
                    var id = svp[prop].PropertyId;
                    var value = svp[prop].Value;
                    switch (id) {
                        case 'String {00062008-0000-0000-c000-000000000046} Name EntityDocument':
                            var entityDocument = JSON.parse(value);
                            break;
                        case 'Boolean {00062008-0000-0000-c000-000000000046} Name EntityExtractionSuccess':
                            var entityExtractionSuccess = JSON.parse(value);
                            break;
                        case 'String {00062008-0000-0000-c000-000000000046} Name EntityExtractionServiceDiagnosticContext':
                            var entityExtractionDiagnostics = JSON.parse(value);
                            break;
                        case 'String {00062008-0000-0000-c000-000000000046} Name ExplicitMessageCard':
                            var explicitMessageCard = value;
                            break;
                        case 'String {00062008-0000-0000-c000-000000000046} Name ActionExecutionHttpTrace':
                            var actionExecutionHttpTrace = JSON.parse(value);
                            break;
                        case 'String {00062008-0000-0000-c000-000000000046} Name TeeVersion':
                            var teeVersion = value;
                            break;
                        case 'String 0x7d':
                            var headers = value;
                            break;
                    }
                }
            }

            var bodyHtml = null;

            if (odata.Body.ContentType === "HTML") {

                bodyHtml = odata.Body.Content;
            }

            callback(bodyHtml, messageCard, adaptiveCard, diagnostics, explicitMessageCard, actionExecutionHttpTrace, entityExtractionSuccess, entityDocument, entityExtractionDiagnostics, headers, null);
        }
    });
};

RestApiClient.prototype.ewsIdToDataId = function (id) {
    return Office.context.mailbox.convertToRestId(
        id,
        Office.MailboxEnums.RestVersion.v2_0
    );
};