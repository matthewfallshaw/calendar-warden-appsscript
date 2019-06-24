"use strict";
var ACTION_MAKE_FREE = 'makefree';
var ACTION_MAKE_BUSY = 'makebusy';
var NAG_PROPERTY_PREFIX = '_NAG:';
var TAG_MAKE_FREE = 'free-calendar-event';
var TAG_MAKE_BUSY = 'busy-calendar-event';
var TAG_MADE_FREE = 'calendar-event-made-free';
var TAG_MADE_BUSY = 'calendar-event-made-busy';
var EventWithCalendarId = /** @class */ (function () {
    function EventWithCalendarId(calendarId, event) {
        workaroundAssign(this, event); // typescript happily emits Object.assign, which fails on GAS
        this.calendarId = calendarId;
        // TODO: make an `assert_type` wrapper for this pattern
        if (typeof event.id === 'string') {
            this.id = event.id; // TODO: already assigned by workaroundAssign, but this to cause typescript
            // to see an assignment
        }
        else {
            throw new Error('Error: Event has no id. ' + JSON.stringify(event));
        }
    }
    return EventWithCalendarId;
}());
/**
 * Web App server. Handle GET requests (from links in generated emails).
 *
 * @param {IRequest} request
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(request) {
    // Should really be a POST, but that's hard to do from an email
    var params = request.parameter;
    if (!validParams(params)) {
        return responseInvalidParams(request);
    }
    var event = getCalendarEvent(params.calendarId, params.eventId);
    if (!event) {
        return responseEventNotFound(request);
    }
    try {
        var report = makeEventBe(params.action, event);
        return responseSuccess(report);
    }
    catch (err) {
        return responseError(request, event, err);
    }
}
function validParams(params) {
    return (params.calendarId && params.eventId && params.action && true || false);
}
function responseInvalidParams(request) {
    var message = '<p>Missing parameters:</p>\n<ol>\n';
    if (!request.parameter.calendarId) {
        message += '<li>Missing calendarId.</li>\n';
    }
    if (!request.parameter.eventId) {
        message += '<li>Missing eventId.</li>\n';
    }
    if (!request.parameter.action) {
        message += '<li>Missing action.</li>\n';
    }
    message += '</ol>';
    var template = standardTemplate();
    template.status = 'error';
    template.message = message;
    template.detail = json_out(request);
    return responseStandard(template);
}
function responseEventNotFound(request) {
    var template = standardTemplate();
    template.status = 'error';
    template.message = '<p>Event not found.</p>\n';
    template.detail = json_out(request);
    return responseStandard(template);
}
function responseError(request, event, err) {
    var template = standardTemplate();
    template.status = 'error';
    template.message = '<p>' + err + '</p>\n';
    template.detail = json_out({ request: request, event: event, error: err });
    return responseStandard(template);
}
function responseSuccess(report) {
    var event = report.event;
    var message = '<p><a target="_parent" href="' + event.htmlLink + '">Event</a> ' +
        report.message + '.</p>\n';
    var template = standardTemplate();
    template.status = 'success';
    template.message = message;
    template.detail = json_out(event);
    return responseStandard(template);
}
function standardTemplate() {
    var template = HtmlService.createTemplateFromFile('Index');
    return template;
}
function responseStandard(template) {
    var output = template.evaluate();
    output.setTitle('Calendar Warden');
    output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    output.setFaviconUrl(googleCalendarFavicon());
    return output;
}
function googleCalendarFavicon() {
    return 'https://calendar.google.com/googlecalendar/images/favicon_v2014_' + (new Date()).getDate() + '.ico';
}
/**
 * The next 50 upcoming events in a calendar.
 *
 * @param {string} calendarId
 * @returns {EventWithCalendarId[]}
 */
function upcomingEvents(calendarId) {
    return listCalendarEvents(calendarId);
}
function eventsEmailSubject(events) {
    var subject = events.map(function (event) {
        var match = event.summary && event.summary.match(/^.{0,25}\w{0,15}/);
        var summary = match && match[0] || 'Unnamed event';
        if (event.summary && event.summary.length || 0 > summary.length) {
            return summary + '…';
        }
        else {
            return summary;
        }
    }).join(' | ');
    return subject;
}
function eventsEmailBody(introText, events) {
    return introText +
        '\n\n<ul>\n' +
        events.reduce(function (acc, event) {
            return acc + '  <li>' + eventSummary(event) + '</li>\n';
        }, '') +
        '</ul>';
}
function eventsEmailThreadId(events) {
    return md5(events.reduce(function (acc, e, ii) { return ii === 0 ? e.id : acc + '|' + e.id; }, '')) + '@' + Session.getActiveUser().getEmail().replace(/@.*/, '');
}
/**
 * Email a list of events, unless that list is empty.
 *
 * @param {string} introText
 * @param {EventWithCalendarId[]} events
 * @returns {EventWithCalendarId[]}
 */
function sendEmailListingEvents(introText, events) {
    var body = eventsEmailBody(introText, events);
    var subject = eventsEmailSubject(events);
    var threadId = eventsEmailThreadId(events);
    var emailAddress = Session.getActiveUser().getEmail();
    sendEmail(body, {
        'Content-Type': 'text/html; charset=UTF-8',
        'From': emailAddress,
        'In-Reply-To': '<' + threadId + '>',
        'Subject': 'CalendarWarden: ' + subject,
        'To': emailAddress
    });
    return events;
}
/**
 * Remove alerts from events.
 *
 * @param {EventWithCalendarId[]} events
 * @returns {EventWithCalendarId[]}
 */
function removeAlerts(events) {
    return events.map(function (event) {
        if (event.reminders) {
            event.reminders.useDefault = false;
            delete event.reminders;
        }
        else {
            event.reminders = { useDefault: false };
        }
        var patchedEvent = patchCalendarEvent(event);
        return patchedEvent;
    });
}
/**
 * Change state of event.
 *
 * @param {Action} action
 * @param {EventWithCalendarId} event
 * @returns {IReport}
 */
function makeEventBe(action, event) {
    var report;
    switch (action) {
        case ACTION_MAKE_FREE:
            report = makeEventBeFree(event);
            break;
        case ACTION_MAKE_BUSY:
            report = makeEventBeBusy(event);
            break;
        default:
            throw new Error(('Unrecognised action `' + action +
                '`; I know how to `' + ACTION_MAKE_FREE + '` and `' + ACTION_MAKE_BUSY + '`. ' +
                json_out(action) + ' | ' + event));
    }
    return report;
}
function makeEventBeFree(event) {
    var message;
    var dirty = false;
    if (event.transparency === 'transparent') {
        message = 'already <em>free</em>';
    }
    else {
        dirty = true;
        event.transparency = 'transparent';
        message = 'made <em>free</em>';
    }
    if (event.description) {
        if (event.description.match(/\[busy\]/)) {
            dirty = true;
            event.description = event.description.replace(/\n?\[busy\]/, '');
            message = message + ' (<code>[busy]</code> removed from description)';
        }
    }
    var patchedEvent = dirty ? patchCalendarEvent(event) : event;
    var report = { event: patchedEvent, message: message };
    return report;
}
function makeEventBeBusy(event) {
    var message;
    var dirty = false;
    if (event.transparency === 'opaque') {
        message = 'already <em>busy</em>';
    }
    else {
        dirty = true;
        event.transparency = 'opaque';
        message = 'made <em>busy</em>';
    }
    if (event.description && event.description.match(/\[busy\]/)) {
        message += ' (<code>[busy]</code> already present in description)';
    }
    else {
        dirty = true;
        event.description = event.description ? event.description + '\n[busy]' : '[busy]';
    }
    var patchedEvent = dirty ? patchCalendarEvent(event) : event;
    var report = { event: patchedEvent, message: message };
    return report;
}
function is_auto_processible_event(event) {
    return (event.description &&
        event.description.match(/tram home$|tram to|practice$|climbing$/) &&
        !event.description.match(/matt|fallabria/i) && true || false);
}
function is_block_event(event) {
    return event.summary && event.summary.match(/^block$/i) && true || false;
}
/**
 * checkAndFixCalendar
 *
 * @param {string} calendarId
 * @returns {EventWithCalendarId[]}
 */
function checkAndFixCalendar(calendarId) {
    var now = new Date().getTime();
    var threeDays = new Date(1000 * 60 * 60 * 24 * 3).getTime();
    var threeDaysAgo = (new Date(now - threeDays)).getTime();
    var events = upcomingEvents(calendarId);
    removeAlerts(events.filter(is_block_event));
    // TODO:
    // addTravelTime(events.filter(is_appointment));
    // Email list of noncomplying & blocking events in default calendar.
    if (calendarId === getProperty('defaultCalendar')) {
        var eventsAfterAutos = events.filter(function (event) {
            var lastNagged = (new Date(Number(getProperty(nagPropertyKey(event), false)))).getTime();
            var recentlyNagged = lastNagged && (lastNagged > threeDaysAgo) || false;
            var markedFree = (event.transparency === 'transparent');
            var taggedBusy = !(!event.description || !event.description.match(/\[busy\]/));
            return !markedFree && !taggedBusy && !recentlyNagged;
        }).filter(function (event) {
            if (is_auto_processible_event(event)) {
                try {
                    makeEventBe(ACTION_MAKE_FREE, event);
                    return false;
                }
                catch (_) {
                    return true;
                }
            }
            else {
                return true;
            }
        });
        if (eventsAfterAutos.length > 0) {
            var emailMessage = "<p>The following calendar events are blocking your availability, but don't contain " +
                '<code>[busy]</code> in their descriptions. Either make them <em>Free</em> or add ' +
                "<code>[busy]</code> to their descriptions (or I'll keep bugging you about them).";
            return sendEmailListingEvents(emailMessage, eventsAfterAutos)
                .map(function (event) {
                PropertiesService.getScriptProperties().setProperty(nagPropertyKey(event), now.toString());
                return event;
            });
        }
        else {
            return []; // no email to send
        }
    }
    else {
        return [];
    }
}
/**
 * checkCalendars
 *
 * @returns {void}
 */
function checkCalendars() {
    checkAndFixCalendar(getProperty('defaultCalendar'));
}
/**
 * Check calendar, triggered from Calendar change event.
 *
 * @param {GoogleAppsScript.Events.CalendarEventUpdated} data
 * @returns {EventWithCalendarId[]}
 */
function checkCalendarFromChangeEvent(data) {
    // { calendarId=example@gmail.com, authMode=FULL, triggerUid=706555 }
    var calendarId = data.calendarId;
    return checkAndFixCalendar(calendarId);
}
/**
 * Check email account for flagged events.
 *
 * @returns {void}
 */
function checkEmailAccount() {
    [TAG_MAKE_FREE, TAG_MAKE_BUSY].map(function (tag) {
        // Search for events with TAG_MAKE_FREE or TAG_MAKE_BUSY
        var label = GmailApp.getUserLabelByName(tag);
        if (!label) {
            label = GmailApp.createLabel(tag);
        }
        var threads = label.getThreads();
        if (!threads || threads.length === 0) {
            return false;
        }
        var tagDone = (tag === TAG_MAKE_FREE) ? TAG_MADE_FREE : TAG_MADE_BUSY;
        var labelDone = GmailApp.getUserLabelByName(tagDone) || GmailApp.createLabel(tagDone);
        // for all threads
        threads.map(function (thread) {
            thread.getMessages().map(function (message) {
                var re = /data-calendar-id="([^"]+)" +data-event-id="([^"]+)"/g;
                var body = message.getBody();
                // find events
                var matches = re.exec(body);
                while (matches !== null) {
                    var calendarId = matches[1];
                    var eventId = matches[2];
                    var event = getCalendarEvent(calendarId, eventId);
                    // make free or busy and replace with TAG_MADE_FREE|BUSY label
                    if (event) {
                        if (tag === TAG_MAKE_FREE) {
                            makeEventBe(ACTION_MAKE_FREE, event);
                            console.info({
                                'From checkEmailAccount:': "Event " + event.htmlLink + " made free (from email subject:\"" + message.getSubject() + "\")"
                            });
                        }
                        else {
                            makeEventBe(ACTION_MAKE_BUSY, event);
                            console.info({
                                'From checkEmailAccount:': "Event " + event.htmlLink + " made busy (from email subject:\"" + message.getSubject() + "\")"
                            });
                        }
                    }
                    matches = re.exec(body);
                }
            });
            labelDone.addToThread(thread);
            label.removeFromThread(thread);
        });
        return true;
    });
}
function checkEmail() {
    checkEmailAccount();
}
/**
 * Clear event nag timestamp properties.
 * A periodic cleanup function run by a timed trigger.
 *
 * @returns {void}
 */
function clearOldNagTimestamps() {
    var threeDaysAgo = new Date().getTime() - 1000 * 60 * 60 * 24 * 3;
    var scriptProperties = PropertiesService.getScriptProperties();
    var properties = scriptProperties.getProperties();
    Object.keys(properties)
        .filter(function (key) { return key.match(new RegExp('^' + NAG_PROPERTY_PREFIX)); })
        .map(function (key) {
        if (!isNaN(Number(properties[key])) && (Number(properties[key])) < threeDaysAgo) {
            scriptProperties.deleteProperty(key);
        }
    });
}
// ## Utility functions
function missingPropertyErrorMessage(property) {
    return 'No ' + property + ' script property: Set a calendar email address as the value for a `' + property +
        '` key as a "Script Property" (from the script editor on `script.google.com` > "Project properties" ' +
        '> "Script properties")';
}
function json_out(json) {
    return JSON.stringify(json, null, 2);
}
function include(filename) {
    var content = HtmlService.createHtmlOutputFromFile(filename).getContent();
    return content;
}
function eventActionURL(action, event) {
    return getProperty('selfLink') + '?calendarId=' + event.calendarId + '&eventId=' + event.id + '&action=' + action;
}
function eventSummary(event) {
    return eventLink(event) + '<br />' +
        '(' + (event.start ? event.start.date || event.start.dateTime : 'no start date') + ') ' +
        '[' + eventActionLink(ACTION_MAKE_FREE, event) + '] [' + eventActionLink(ACTION_MAKE_BUSY, event) + ']';
}
function eventLink(event) {
    assert(!!event, 'No Event');
    assert(!!event.htmlLink, 'No event.htmlLink', event);
    assert(!!event.summary, 'No event.summary', event);
    return '<a ' +
        'href="' + event.htmlLink + '" ' +
        'data-calendar-id="' + event.calendarId + '" data-event-id="' + event.id + '"' +
        '>' +
        event.summary +
        '</a>';
}
function eventActionLink(action, event) {
    return '<a href="' + eventActionURL(action, event) + '">' +
        'make <em>' + (action === ACTION_MAKE_FREE ? 'Free' : 'Busy') + '</em>' +
        '</a>';
}
function md5(str) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str)
        .reduce(function (outStr, chr) {
        var char = (chr < 0 ? chr + 256 : chr).toString(16);
        return outStr + (char.length === 1 ? '0' : '') + char;
    }, '');
}
function sendEmail(body, headers) {
    var msg = Gmail.newMessage();
    var raw = Object.keys(headers).reduce(function (acc, k, ii) {
        var v = headers[k];
        var e = k + ': ' + v;
        return ii === 0 ? e : acc + '\r\n' + e;
    }, '') +
        '\r\n\r\n' +
        body;
    msg.raw = Utilities.base64EncodeWebSafe(raw);
    var gu = Gmail.Users;
    var gm = gu.Messages;
    gm.send(msg, 'me');
}
function getProperty(propertyKey, errorIfMissing) {
    if (errorIfMissing === void 0) { errorIfMissing = false; }
    var property = PropertiesService.getScriptProperties().getProperty(propertyKey);
    if (errorIfMissing && !property) {
        throw missingPropertyErrorMessage(propertyKey);
    }
    else {
        return property ? property : '';
    }
}
function nagPropertyKey(event) {
    return NAG_PROPERTY_PREFIX + event.id.toString();
}
function withCalendarId(calendarId, event) {
    if (event) {
        return new EventWithCalendarId(calendarId, event);
    }
    else {
        throw new Error('Error: Missing Event!');
    }
}
function withCalendarIds(calendarId, events) {
    return events.map(function (event) { return new EventWithCalendarId(calendarId, event); });
}
function getCalendarEvent(calendarId, eventId) {
    var es = Calendar.Events;
    var gEvent = es.get(calendarId, eventId);
    if (typeof gEvent.id === 'string') {
        return withCalendarId(calendarId, gEvent);
    }
    else {
        throw new Error('Error: Event with no id!' + JSON.stringify(gEvent));
    }
}
function listCalendarEvents(calendarId) {
    var es = Calendar.Events;
    var gEvents = es.list(calendarId, {
        maxResults: 50,
        orderBy: 'startTime',
        singleEvents: true,
        timeMin: (new Date()).toISOString()
    });
    if (gEvents.items) {
        return withCalendarIds(calendarId, gEvents.items);
    }
    else {
        return [];
    }
}
function patchCalendarEvent(event) {
    var es = Calendar.Events;
    return withCalendarId(event.calendarId, es.patch(event, event.calendarId, event.id));
}
// TODO: nice to have a version of this for type assertions
function assert(test) {
    var rest = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        rest[_i - 1] = arguments[_i];
    }
    if (!test) {
        throw new Error('Error: ' + JSON.stringify(rest));
    }
}
var workaroundAssign = function (has) {
    'use strict';
    return assign;
    function assign(target) {
        var source = [];
        for (var _i = 1; _i < arguments.length; _i++) {
            source[_i - 1] = arguments[_i];
        }
        for (var _a = 0, source_1 = source; _a < source_1.length; _a++) {
            var argument = source_1[_a];
            copy(target, argument);
        }
        return target;
    }
    function copy(target, source) {
        for (var key in source) {
            if (has.call(source, key)) {
                target[key] = source[key];
            }
        }
    }
}({}.hasOwnProperty);
