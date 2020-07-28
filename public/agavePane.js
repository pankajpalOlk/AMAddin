

(function () {

    "use strict";

    // Max htmlBody length on calling displayNewMessageForm
    var MAX_HTML_LENGTH = 31 * 1024;

    var MESSAGE_CARD_MIN_OUTLOOK_VERSION = 8431;

    var ADAPTIVE_CARD_MIN_OUTLOOK_VERSION = 9330;

    var OUTLOOKIOS = "OutlookIOS";

    var OUTLOOKANDROID = "OutlookAndroid";

    var OUTLOOKWEB = "OutlookWebApp";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {

        $(document).ready(function () {

            app.initialize();

            // Set up ItemChanged event
            if (Office.context.mailbox.addHandlerAsync !== undefined) {
                Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, emailItemChanged);
            }

            initialize();

            emailItemChanged(null);

        });
    };

    function initialize() {

        // Initialization code goes here

        //AWT.initialize("5e3990d926a54076a97ec4299826ea03-68e14c6b-4173-44a3-9b9e-5efcbfcf5419-6804");
        //AWT.setContext("UserAgent", navigator.userAgent);

        window.addEventListener('error', function (event) {
            // AWT.logEvent({
            //     name: "Error",
            //     properties: {
            //         "HostName": Office.context.mailbox.diagnostics.hostName,
            //         "HostVersion": Office.context.mailbox.diagnostics.hostVersion,
            //         "Message": event.message,
            //         "File": event.filename,
            //         "Line": event.lineno,
            //         "Column": event.colno
            //     }
            // });
        });

        window.addEventListener("beforeunload", function (event) {
            try {
                //AWT.flushAndTeardown();
            } catch (e) {
                // ingore any error
            }
        });

        $('#showAdditional').click(function () {
            $('#dignosticsArea').toggle();
        });

        $('.toggleTitleBar').click(function () {
            $(this).parent().children('.messageDetailClosed, .messageDetailOpen').toggle();
            $(this).children('.chevron').toggleClass("titleClosed titleOpen");
        });
        var hostName = Office.context.mailbox.diagnostics.hostName;
        var hostVersion = Office.context.mailbox.diagnostics.hostVersion;
        var isMacOutlook = hostVersion.indexOf('(') !== -1;

        if (hostName === OUTLOOKIOS || hostName === OUTLOOKANDROID || isMacOutlook) {
            $("#clientNotSupportedMessage").show();
        }

        if (hostName === OUTLOOKIOS || hostName === OUTLOOKANDROID) {
            $(".actioncopy").hide();
        }

        if (hostName === "Outlook") {
            var versionSplits = hostVersion.split('.');
            var mainVersion = versionSplits[0];
            var build = versionSplits[2];

            if (mainVersion < 16) {
                $("#outlook2013Message").show();
            } else if (build < MESSAGE_CARD_MIN_OUTLOOK_VERSION) {
                $("#outlookVersionMessage").show();
            } else if (build < ADAPTIVE_CARD_MIN_OUTLOOK_VERSION) {
                $("#adaptiveCardNotSupportedMessage").show();
            }
        }
    }

    function buildTextBody(messageCard, adaptiveCard, combinedDiagnostics, entityExtractionDiagnostics) {
        var body = 'Sending us this email message means you agree to the Terms of Use (https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx) and Privacy Statement (https://privacy.microsoft.com/privacystatement)\n';
        body += '\n(Please review the information below to make sure it doesn\'t include any personal information.)\n';

        if (adaptiveCard && messageCard) {
            body += '\nBoth Adaptive Card and Message Card found\n';
        } else if (adaptiveCard) {
            body += '\nAdaptive Card found\n';
        } else if (messageCard) {
            body += '\nMessage Card found\n';
        } else {
            body += '\nNo card found\n';
        }

        var jsonFormatted = JSON.stringify(combinedDiagnostics, null, 4);

        body += '\nDiagnostics:\n';

        body += jsonFormatted;

        return body;
    }

    function buildBody(messageCard, adaptiveCard, combinedDiagnostics, entityExtractionDiagnostics) {
        var $body = $('<div>');
        $body.html('<p>Sending us this email message means you agree to the <a href="https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx">Terms of Use</a> and <a href="https://privacy.microsoft.com/privacystatement">Privacy Statement</a>.</p>'
            + '<p><b>(Please review the information below to make sure it doesn\'t include any personal information.)</b></p>'
            + '<br />');

        var jsonFormatted;

        if (adaptiveCard && messageCard) {
            $body.append($('<p>').text('Both Adaptive Card and Message Card found'));
        } else if (adaptiveCard) {
            $body.append($('<p>').text('Adaptive Card found'));
        } else if (messageCard) {
            $body.append($('<p>').text('Messsage Card found'));
        } else {
            $body.append($('<p>').text('No card found'));
        }

        jsonFormatted = JSON.stringify(combinedDiagnostics, null, 4);
        $body.append($('<p>').text('Diagnostics:'));
        $body.append($('<pre>').text(jsonFormatted));

        jsonFormatted = JSON.stringify(entityExtractionDiagnostics, null, 4);
        $body.append($('<p>').text('Entity Extraction Diagnostics:'));
        $body.append($('<pre>').text(jsonFormatted));

        if (adaptiveCard) {
            jsonFormatted = JSON.stringify(adaptiveCard, null, 4);
            $body.append($('<p>').text('Adaptive Card:'));
            $body.append($('<pre>').text(jsonFormatted));
        }

        if (messageCard) {
            jsonFormatted = JSON.stringify(messageCard, null, 4);
            $body.append($('<p>').text('Message Card:'));
            $body.append($('<pre>').text(jsonFormatted));
        }

        var bodyHtml = $body.html();

        if (bodyHtml.length > MAX_HTML_LENGTH){
            return bodyHtml.substr(0, MAX_HTML_LENGTH);
        } else {
            return bodyHtml;
        }
    }

    function addCard(name, card) {
        if (card) {
            var cardFormatted = JSON.stringify(card, null, 4);

            var $block = $('#' + name + "Block");

            $block.find('.actioncopy').off().click(function (event) {
                event.stopPropagation();
                copyToClipboard(cardFormatted);
            });

            $block.find('pre').text(cardFormatted);
            $block.show();
        }
    }

    function addHttpTrace(actionExecutionHttpTrace) {
        if (!actionExecutionHttpTrace) {
            return;
        }

        var $block = $('#actionHistory');
        for (var i = 0; i < actionExecutionHttpTrace.length; i++) {
            var httpTrace = actionExecutionHttpTrace[i];
            if (!httpTrace) {
                continue;
            }

            var $recordBlock = $block.find('#historyRecordTemplate').clone();
            $recordBlock.prop("id", "historyRecord" + i);
            $recordBlock.find(".performedAt").text(httpTrace.performedAt);
            $recordBlock.find(".errorMessage").text(httpTrace.errorMessage);
            $recordBlock.find(".requestMethod").text(httpTrace.requestMethod || "");
            $recordBlock.find(".requestUrl").text(httpTrace.requestUrl || "");
            $recordBlock.find(".requestBody").text(httpTrace.requestBody || "");
            if (httpTrace.responseStatus) {
                $recordBlock.find(".responseStatus").text(httpTrace.responseStatus || "");
            } else {
                $recordBlock.find(".responseStatusBlock").hide();
            }
            $recordBlock.find(".responseStatus").text(httpTrace.responseStatus || "");
            $recordBlock.find(".responseHeaders").text(httpTrace.responseHeaders || "");
            $recordBlock.find(".responseBody").text(httpTrace.responseBody || "");

            $recordBlock.show();

            $block.find("#historyContainer").append($recordBlock);

            $block.show();
        }

        $block.find('.actioncopy').off().click(function (event) {
            event.stopPropagation();
            var str = $('#historyContainer').text();
            // Trim heading blank spaces for each line
            str = str.replace(/^\s+/gm, '');
            copyToClipboard(str);
        });
    }

    function enrichDiagnostics(htmlBody, stampingDiagnostics, explicitMessageCard, entityExtractionSuccess, entityDocument, headers){
        var combinedDiagnostics = {};

        if (stampingDiagnostics) {
            combinedDiagnostics["ActionableMessageStamping"] = stampingDiagnostics;
        }

        combinedDiagnostics['CardEnabledForMessage'] = explicitMessageCard === 'True';

        combinedDiagnostics["ClientName"] = Office.context.mailbox.diagnostics.hostName;

        combinedDiagnostics["ClientVersion"] = Office.context.mailbox.diagnostics.hostVersion;

        var internetMessageId = Office.context.mailbox.item.internetMessageId;
        combinedDiagnostics["InternetMessageId"] = internetMessageId;

        combinedDiagnostics["EntityExtractionSuccess"] = entityExtractionSuccess === true;

        if (combinedDiagnostics.ActionableMessageStamping && combinedDiagnostics.ActionableMessageStamping.Errors && combinedDiagnostics.ActionableMessageStamping.Errors.length > 0) {
            $("#stampingError .messageDetailOpen").empty();
            $.each(combinedDiagnostics.ActionableMessageStamping.Errors, function (idx, err) {
                $("#stampingError .messageDetailOpen").append($('<p>').text(err));
            });
            $("#stampingError").show();
        }

        var messageCardParsed = false;
        var adaptiveCardParsed = false;

        if (entityDocument) {
            var originatorShouldExist = false;
            var originator = null;
            for (var i = 0; i < entityDocument.length; i++) {
                var entity = entityDocument[i];

                switch (entity.type) {
                    case "MessageCard":
                        combinedDiagnostics["MessageCardPayloadParsed"] = true;
                        messageCardParsed = true;
                        originatorShouldExist = true;
                        originator = entity.entities[0].originator;
                        break;
                    case "AdaptiveCard":
                        combinedDiagnostics["AdaptiveCardPayloadParsed"] = true;
                        adaptiveCardParsed = true;
                        originatorShouldExist = true;
                        originator = entity.entities[0].originator;
                        break;
                    case "SignedMessageCard":
                        combinedDiagnostics["SignedMessageCard"] = true;
                        messageCardParsed = true;
                        originatorShouldExist = false;
                        break;
                    case "SignedAdaptiveCard":
                        combinedDiagnostics["SignedAdaptiveCard"] = true;
                        adaptiveCardParsed = true;
                        originatorShouldExist = false;
                        break;
                    default:
                        break;
                }
            }

            if (originator && originator.length > 0) {
                combinedDiagnostics['Originator'] = originator;
            } else if (originatorShouldExist) {
                combinedDiagnostics['Warning'] = 'Originator not set.';
                $("#originatorNotSetMessage").show();
            }
        } else {
            combinedDiagnostics['Error'] = 'EntityDocument does not exist.';
        }

        if (!entityExtractionSuccess) {
            $("#entityExtractionNotRun").show();
        }

        if (!adaptiveCardParsed && htmlBody) {
            var result = checkCardPayload(htmlBody, "adaptivecard");
            combinedDiagnostics['AdaptiveCardPayload'] = result;
            if (entityDocument && result.found) {
                $("#adaptiveCardPayloadError").show();
            }
        }

        if (!messageCardParsed && htmlBody) {
            result = checkCardPayload(htmlBody, "messagecard");
            combinedDiagnostics['MessageCardPayload'] = result;

            // Note: SignedAdaptiveCard maybe in the message card tag, do not show the error message in this case
            if (entityDocument && result.found && result.type.toLowerCase().indexOf("messagecard") !== -1) {
                $("#messageCardPayloadError").show();
            }
        }

        if (headers) {
            var authHeader = getAuthHeader(headers);
            combinedDiagnostics['AuthHeader'] = authHeader;

            if (!authPass(authHeader)) {
                $("#spfDkimFailMessage").show();
            }
        }

        $('#diagnosticsBlock').find('pre').JSONView(combinedDiagnostics);
        $('#diagnosticsBlock').show();

        return combinedDiagnostics;
    }

    function authPass(authHeader) {

        // 1. If “X-MS-Exchange-Organization-AuthAs” exists and the value is “Internal”, then return pass.
        if (authHeader.authAs && authHeader.authAs === "Internal") {
            return true;
        }

        // If “Authentication-Results” header is missing, then return pass.
        if (authHeader.results === undefined) {
            return true;
        }

        if (!authHeader.results) {
            return false;
        }

        // 3.a. “dmarc=fail”, return fail;
        if (authHeader.results.indexOf('dmarc=fail') !== -1) {
            return false;
        }

        // 3.b. “dmarc=pass” or “dmarc=bestguesspass”, return pass;
        if (authHeader.results.indexOf('dmarc=pass') !== -1 || authHeader.results.indexOf('dmarc=bestguesspass') !== -1) {
            return true;
        }

        // 3.c. “dkim=fail”, return fail;
        if (authHeader.results.indexOf('dkim=fail') !== -1) {
            return false;
        }

        // 3.d. "spf=pass" or "dkim=pass", return pass;
        if (authHeader.results.indexOf('spf=pass') !== -1 || authHeader.results.indexOf('dkim=pass') !== -1) {
            return true;
        }

        return false;
    }

    function getAuthHeader(headers) {
        var authHeader = {};

        // Header parsing logic borrowed from https://github.com/stephenegriffin/MHA/blob/master/Scripts/Headers.js
        var lines = headers.split(/[\n\r]+/);

        var headerList = [];
        var iNextHeader = 0;
        // Unfold lines
        for (var iLine = 0; iLine < lines.length; iLine++) {
            var line = lines[iLine];
            // Skip empty lines
            if (line === "") continue;

            // Recognizing a header:
            // - First colon comes before first white space.
            // - We're not strictly honoring white space folding because initial white space
            // - is commonly lost. Instead, we heuristically assume that space before a colon must have been folded.
            // This expression will give us:
            // match[1] - everything before the first colon, assuming no spaces (header).
            // match[2] - everything after the first colon (value).
            var match = line.match(/(^[\w-.]*?): ?(.*)/);

            // There's one false positive we might get: if the time in a Received header has been
            // folded to the next line, the line might start with something like "16:20:05 -0400".
            // This matches our regular expression. The RFC does not preclude such a header, but I've
            // never seen one in practice, so we check for and exclude 'headers' that
            // consist only of 1 or 2 digits.
            if (match && match[1] && !match[1].match(/^\d{1,2}$/)) {
                headerList[iNextHeader] = line;
                iNextHeader++;
            } else {
                if (iNextHeader > 0) {
                    // Tack this line to the previous line
                    // All folding whitespace should collapse to a single space
                    line = line.replace(/^[\s]+/, "");
                    if (!line) continue;
                    var separator = headerList[iNextHeader - 1] ? " " : "";
                    headerList[iNextHeader - 1] += separator + line;
                } else {
                    // If we didn't have a previous line, go ahead and use this line
                    if (line.match(/\S/g)) {
                        headerList[iNextHeader] = line;
                        iNextHeader++;
                    }
                }
            }
        }

        for (var iHeader = 0; iHeader < headerList.length; iHeader++) {
            line = headerList[iHeader];

            if (line.startsWith('X-MS-Exchange-Organization-AuthAs:') && authHeader.authAs === undefined) {
                authHeader.authAs = line.replace('X-MS-Exchange-Organization-AuthAs: ', '');
                continue;
            }

            if (line.startsWith('Authentication-Results:') && authHeader.results === undefined) {
                authHeader.results = line.replace('Authentication-Results: ', '');
                continue;
            }
        }

        return authHeader;
    }

    function checkCardPayload(html, cardType) {
        var found = false;
        var type = null;
        var adaptiveCardTagRegex = /<script\s+?type=\"application\/adaptivecard\S+?json\".*?>([^]*?)<\/script>/mi;
        var messageCardTagRegex = /<script\s+?type=\"application\/ld\S+?json\".*?>([^]*?)<\/script>/mi;
        var adaptiveCardTypeRegex = /\"type"\:\s*\"(\w+?)\"/mi;
        var messageCardTypeRegex = /\".?type"\:\s*\"(\w+?)\"/mi;

        var cardTagRegex = adaptiveCardTagRegex;
        var typeRegex = adaptiveCardTypeRegex;
        if (cardType.toLowerCase() === "messagecard") {
            cardTagRegex = messageCardTagRegex;
            typeRegex = messageCardTypeRegex;
        }

        var match = html.match(cardTagRegex);
        if (match) {
            var scriptText = match[1];

            match = scriptText.match(typeRegex);

            if (match) {
                type = match[1];
            }

            if (cardType === "messagecard") {
                // script tag with application/ld+json is also used for list view actions, tighten found logic
                if (type) {
                    var typeLower = type.toLowerCase();
                    found = typeLower === "messagecard" || typeLower === "adaptivecard" || typeLower === "signedmessagecard" || typeLower === "signedadaptivecard";
                }
            } else {
                found = true;
            }
        }

        var signedCardMicrodataRegex = /<\w+\s+?itemprop=\"@type\".*?content=\"(\w+?)\".*?>/mi;
        match = html.match(signedCardMicrodataRegex);
        if (match && match[1].toLowerCase().indexOf(cardType) !== -1) {
            found = true;
            type = match[1] + "_MicroData";
        }

        var result = { found: found, type: type };
        return result;
    }

    function copyToClipboard(str){
        var el = document.createElement('textarea');
        el.value = str;
        el.setAttribute('readonly', '');
        el.style.position = 'absolute';
        el.style.left = '-9999px';
        document.body.appendChild(el);
        el.select();
        document.execCommand('copy');
        document.body.removeChild(el);
    }

    function getTokenContext(callback) {

        Office.context.mailbox.getCallbackTokenAsync(
            { isRest: true },
            function (tr) {
                if (tr.status === 'succeeded') {
                    var context = {
                        user: Office.context.mailbox.userProfile.emailAddress,
                        item: Office.context.mailbox.item.itemId,
                        token: tr.value
                    };
                    callback(context, null);
                }
                else {
                    callback(null, tr.status);
                }
            });
    }

    function emailItemChanged(eventArgs) {
        // Reset the status
        $('#errorMessage').text('Loading...');
        $('#errorMessage').show();
        $('#main').hide();
        $('#blocksContainer').find('.messageBlock').hide();

        getTokenContext(function (context, error) {
            if (error) {
                $('#errorMessage').text(error);
            } else {
                var client = new RestApiClient(Office.context.mailbox.restUrl, context.user);

                client.loadProperties(context.item, context.token, function (htmlBody, messageCard, adaptiveCard, diagnostics, explicitMessageCard, actionExecutionHttpTrace, entityExtractionSuccess, entityDocument, entityExtractionDiagnostics, headers, err) {
                    if (err === "0") {
                        $('#errorMessage').hide();
                        $('#restError').show();
                        return;
                    }

                    if (err) {
                        $('#errorMessage').text(err);
                        return;
                    }

                    if (messageCard === null && adaptiveCard === null) {
                        if (!err) {
                            $('#cardNotFound').show();

                            // Show Html body if no card found
                            var $block = $('#htmlBodyBlock');

                            $block.find('.actioncopy').off().click(function (event) {
                                event.stopPropagation();
                                copyToClipboard(htmlBody);
                            });

                            $block.find('pre').text(htmlBody);
                            $block.show();
                        }
                    }

                    $('#errorMessage').hide();
                    $('#main').show();

                    addCard('messageCard', messageCard);
                    addCard('adaptiveCard', adaptiveCard);

                    addHttpTrace(actionExecutionHttpTrace);

                    var combinedDiagnostics = enrichDiagnostics(htmlBody, diagnostics, explicitMessageCard, entityExtractionSuccess, entityDocument, headers);

                    var hostName = Office.context.mailbox.diagnostics.hostName;
                    var hostVersion = Office.context.mailbox.diagnostics.hostVersion;
                    var mainVersion = hostVersion.split('.')[0];

                    var useSendMailLink = mainVersion < 16 || hostName === OUTLOOKIOS || hostName === OUTLOOKANDROID || hostName === OUTLOOKWEB || !Office.context.mailbox.displayNewMessageForm;

                    // if (useSendMailLink) {
                    //     $('#sendMailLink').show();
                    //     $('#sendMail').hide();
                    // } else {
                    //     $('#sendMailLink').hide();
                    //     $('#sendMail').show();
                    // }
                    // var mailLinkParams = {
                    //     subject: 'Actionable Message Issue Report for message ' + Office.context.mailbox.item.internetMessageId,
                    //     body: buildTextBody(messageCard, adaptiveCard, combinedDiagnostics, entityExtractionDiagnostics)
                    // };

                    // var paramsEncoded = $.param(mailLinkParams).replace(/\+/g, "%20");

                    // $('#sendMailLink').attr('href', 'mailto:onboardoam@microsoft.com?' + paramsEncoded);

                    // // Send mail link handling for android, with will show "unable to open" page for mailto link in iframe
                    // if (hostName === OUTLOOKANDROID) {
                    //     $('#sendMailLink').attr('target', '_blank');
                    // }

                    // $('#sendMail').off().click(function () {
                    //     if (Office.context.mailbox.displayNewMessageForm !== undefined) {
                    //         try {
                    //             Office.context.mailbox.displayNewMessageForm(
                    //                 {
                    //                     toRecipients: ['onboardoam@microsoft.com'],
                    //                     subject: 'Actionable Message Issue Report for message ' + Office.context.mailbox.item.internetMessageId,
                    //                     htmlBody: buildBody(messageCard, adaptiveCard, combinedDiagnostics, entityExtractionDiagnostics)
                    //                 });
                    //         } catch (e) {
                    //             app.showNotification("Send mail failed", "Please copy the diagnostics content and send it to onboardoam@microsoft.com.");
                    //         }
                    //     } else {
                    //         app.showNotification("Send mail not supported", "Please copy the diagnostics content and send it to onboardoam@microsoft.com.");
                    //     }
                    // });

                    // AWT.logEvent({
                    //     name: "DiagnosticsShown",
                    //     properties: {
                    //         "HostName": hostName,
                    //         "HostVersion": hostVersion,
                    //         "AdaptiveCard": adaptiveCard !== null,
                    //         "MessageCard": messageCard !== null,
                    //         "EntityDocument": entityDocument !== null,
                    //         "AddinHost": document.location.host
                    //     }
                    // });
                });
            }
        });
    }
})();



