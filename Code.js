'use strict';

// `const` raises redeclaration error
var ACTION_MAKE_FREE = 'makefree';
var ACTION_MAKE_BUSY = 'makebusy';
var NAG_PROPERTY_PREFIX = '_NAG:';
var TAG_MAKE_FREE = 'free-calendar-event';
var TAG_MAKE_BUSY = 'busy-calendar-event';
var TAG_MADE_FREE = 'calendar-event-made-free';
var TAG_MADE_BUSY = 'calendar-event-made-busy';


/**
 * Web App server. Handle GET requests (from links in generated emails).
 */
function doGet(request) {  // Should really be a POST, but that's hard to do from an email
  const params = request.parameter;

  if (!validParams(params)) { return invalidParamsResponse(request); }

  var event = Calendar.Events.get(params.calendarId, params.eventId);
  if (!event) { return eventNotFoundResponse(request); }

  try {
    const report = makeEventBe({ action: params.action, calendarId: params.calendarId }, event);
    return successResponse(report, request);
  } catch(err) {
    return errorResponse(request, event, err);
  }  
}


function validParams(params) {
  return (params.calendarId && params.eventId && params.action);
}

function invalidParamsResponse(request) {
  var template = standardTemplate();
  template.status = 'error';
  template.message = '<p>Missing parameters:</p>\n<ol>\n';
  if (!params.calendarId) { template.message += '<li>Missing calendarId.</li>\n' }
  if (!params.eventId)    { template.message += '<li>Missing eventId.</li>\n' }
  if (!params.action)     { template.message += '<li>Missing action.</li>\n' }
  template.message += '</ol>';
  template.detail = json_out(request);
  return standardResponse(template);
}

function eventNotFoundResponse(request) {
  var template = standardTemplate();
  template.status = 'error';
  template.message = '<p>Event not found.</p>\n';
  template.detail = json_out(request);
  return standardResponse(template);
}

function errorResponse(request, event, err) {
  var template = standardTemplate();
  template.status = 'error';
  template.message = '<p>' + err + '</p>\n';
  template.detail = json_out({ request: request, event: event, error: err });
  return standardResponse(template);
}

function successResponse(report, request) {
  var template = standardTemplate();
  const event = report.event;
  const message = report.message;
  template.status = 'success';
  template.message = '<p><a target="_parent" href="' +  event.htmlLink + '">Event</a> ' +
                     message + '.</p>\n';
  template.detail = json_out(event);
  return standardResponse(template);
}

function standardTemplate() {
  return HtmlService.createTemplateFromFile('Index');
}

function standardResponse(template) {
  const output = template.evaluate();
  output.setTitle('Calendar Warden');
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  output.setFaviconUrl(googleCalendarFavicon());
  return output;
}

function googleCalendarFavicon() {
  return 'https://calendar.google.com/googlecalendar/images/favicon_v2014_' + (new Date).getDate() + '.ico';
}


/**
 * The next 50 upcoming events in a calendar.
 * ([Event])
 */
function upcomingEvents(calendarId) {
  const now = new Date();
  const response = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 50
  });
  if (response.items) {
    return response.items.map(function(event) { event.calendarId = calendarId; return event });
  } else {
    return [];
  }
}

/**
 * Email subject header
 * (String)
 */
function eventsEmailSubject(events) {
  const subject = events.map(function(event) {
    const summary = event.summary && event.summary.match(/^.{0,25}\w{0,15}/)[0] || 'Unnamed event';
    if (event.summary.length > summary.length) {
      return summary + '…';
    } else {
      return summary;
    }
  }).join(' | ');
  return subject;
}

/**
 * Email body text
 * (String)
 */
function eventsEmailBody(intro_text, events) {
  return intro_text +
    '\n\n<ul>\n' +
    events.reduce(function(acc, event) {
      return acc + '  <li>' + eventSummary(event) + '</li>\n';
    }, '') +
    '</ul>';
}

function eventsEmailThreadId(events) {
  return md5(events.reduce(function(acc, e, ii) { return ii === 0 ? e.id : acc + '|' + e.id }, '')) +
    '@' + Session.getActiveUser().getEmail().replace(/@.*/,'');
}

/**
 * Email a list of events, unless that list is empty.
 * (Returns the event list for chaining.)
 */
function sendEmailListingEvents(intro_text, events) {
  if (events && (events.length > 0)) {
    const body = eventsEmailBody(intro_text, events);
    const subject = eventsEmailSubject(events);
    const thread_id = eventsEmailThreadId(events);
    const email_address = Session.getActiveUser().getEmail();
    sendEmail(
      body,
      {
        To: email_address,
        From: email_address,
        Subject: 'CalendarWarden: ' + subject,
        'Content-Type': 'text/html; charset=UTF-8',
        'In-Reply-To': '<' + thread_id + '>',
      }
    )
    return events;
  } else {
    return [];
  }
}

/**
 * Remove alerts from events.
 * (Returns the event list for chaining.)
 */
function removeAlertsFromBlockEvents(events) {
  return events.map(function(event) {
    console.log('Removed reminders from event.', event);
    event.reminders.useDefault = false;
    event.reminders.overrides = [];
    Calendar.Events.patch(event, event.calendarId, event.id)
    return event;
  })
}

/**
 * Change state of event.
 * (Returns the event for chaining.)
 */
function makeEventBe(action, event) {
  var report;
  switch (action.action) {
    case ACTION_MAKE_FREE:
      report = makeEventBeFree(action, event);
      break;
    case ACTION_MAKE_BUSY:
      report = makeEventBeBusy(action, event);
      break;
    default:
      throw ('Unrecognised action `' + action.action +
             '`; I know how to `' + ACTION_MAKE_FREE + '` and `' + ACTION_MAKE_BUSY + '`. ' +
             json_out(action) + ' | ' + event);
  }

  return report;
}

function makeEventBeFree(action, event) {
  var message;

  if (event.transparency == 'transparent') {
    message = 'already <em>free</em>';
  } else {
    event.transparency = 'transparent';
    message = 'made <em>free</em>';
  }
  if (event.description) {
    if (event.description.match(/\[busy\]/)) {
      event.description = event.description.replace(/\n?\[busy\]/,'');
      message = message + ' (<code>[busy]</code> removed from description)';
    }
  }

  event = Calendar.Events.patch(event, action.calendarId, event.id);

  return { event: event, message: message};
}

function makeEventBeBusy(action, event) {
  var message;

  if (event.transparency == 'opaque') {
    message = 'already <em>busy</em>';
  } else {
    event.transparency = 'opaque';
    message = 'made <em>busy</em>';
  }
  if (event.description && event.description.match(/\[busy\]/)) {
    message = message + ' (<code>[busy]</code> already present in description)';
  } else {
    event.description = event.description ? event.description + '\n[busy]' : '[busy]';
  }

  event = Calendar.Events.patch(event, action.calendarId, event.id);

  return { event: event, message: message};
}


function is_auto_processible_event(event) {
  return (event.description &&
          event.description.match(/tram home$|tram to|practice$|climbing$/) &&
          !event.description.match(/matt|fallabria/i))
}

function is_block_event(event) {
  return event.summary && event.summary.match(/^block$/i);
}

/**
 * Check calendar.
 */
function checkCalendar(calendarId) {
  const now = new Date().getTime()
  const three_days_ago = now - 1000 * 60 * 60 * 24 * 3;

  const events = upcomingEvents(calendarId);

  removeAlertsFromBlockEvents(events.filter(is_block_event));

  // Email list of noncomplying & blocking events in default calendar.
  if (calendarId === getProperty('defaultCalendar')) {
    const events_after_autos = events.filter(function (event) {
        const last_nagged = getProperty(nagPropertyKey(event), false);
        const recently_nagged = last_nagged && (last_nagged > three_days_ago);
        const marked_free = (event.transparency == 'transparent');
        const not_tagged_busy = (!event.description || !event.description.match(/\[busy\]/));
        return !marked_free && not_tagged_busy && !recently_nagged;
      }).filter(function(event) {
        if (is_auto_processible_event(event)) {
          try {
            makeEventBe({ action: ACTION_MAKE_FREE, calendarId: calendarId }, event);
            return false;
          } catch(_) {
            return true;
          }
        } else {
          return true;
        }
      });
    const email_message = "<p>The following calendar events are blocking your availability, but don't contain " +
      '<code>[busy]</code> in their descriptions. Either make them <em>Free</em> or add ' +
      "<code>[busy]</code> to their descriptions (or I'll keep bugging you about them).";
    sendEmailListingEvents(email_message, events_after_autos)
      .map(function (event) {  // update last nagged timestamps
        PropertiesService.getScriptProperties().setProperty(nagPropertyKey(event), now);
        return event;
      });
  }
}

/**
 * Check calendars.
 */
function checkCalendars() {
  checkCalendar(getProperty('defaultCalendar'));
}

/**
 * Check calendar, triggered from Calendar change event.
 */
function checkCalendarFromChangeEvent(data) {
  // { calendarId=example@gmail.com, authMode=FULL, triggerUid=706555 }
  const calendarId = data.calendarId;
  checkCalendar(calendarId);
}

/**
 * Check email account for flagged events.
 */
function checkEmailAccount(account) {
  [TAG_MAKE_FREE, TAG_MAKE_BUSY].map(function(tag) {
    // Search for events with TAG_MAKE_FREE or TAG_MAKE_BUSY
    var label = GmailApp.getUserLabelByName(tag);
    if (!label) { label = GmailApp.createLabel(tag); }
    var threads = label.getThreads();
    if (!threads || threads.length == 0) { return false; }
    var tag_done = (tag == TAG_MAKE_FREE) ? TAG_MADE_FREE : TAG_MADE_BUSY;
    var label_done = GmailApp.getUserLabelByName(tag_done) || GmailApp.createLabel(tag_done);
    // for all threads
    threads.map(function(thread) {
      thread.getMessages().map(function(message) {  // assume threads have >0 messages
        var re = /data-calendar-id="([^"]+)" +data-event-id="([^"]+)"/g;
        var body = message.getBody();
        // find events
        var matches = re.exec(body);
        while (matches !== null) {
          var calendar_id = matches[1];
          var event_id = matches[2];
          var event = Calendar.Events.get(calendar_id, event_id);
          // make free or busy and replace with TAG_MADE_FREE|BUSY label
          if (tag == TAG_MAKE_FREE) {
            makeEventBe({ action: ACTION_MAKE_FREE, calendarId: calendar_id }, event);
          } else {
            makeEventBe({ action: ACTION_MAKE_BUSY, calendarId: calendar_id }, event);
          }
          matches = re.exec(body);
        };
      })
      label_done.addToThread(thread);
      label.removeFromThread(thread);
    });
  });
}

function checkEmail() {
  checkEmailAccount(Session.getActiveUser().getEmail());
}


/**
 * Clear event nag timestamp properties.
 * A periodic cleanup function run by a timed trigger.
 */
function clearOldNagTimestamps() {
  const three_days_ago = new Date().getTime() - 1000 * 60 * 60 * 24 * 3;
  const properties = PropertiesService.getScriptProperties().getProperties();
  Object.keys(properties)
    .filter(function(key) { return key.match(new RegExp('^' + NAG_PROPERTY_PREFIX)) })
    .map(function (key) {
      if (!isNaN(Number(properties[key])) &&(Number(properties[key])) < three_days_ago) {
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
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function eventActionURL(action, event) {
  return getProperty('selfLink') + '?calendarId=' + event.calendarId + '&eventId=' + event.id + '&action=' + action
}

function eventSummary(event) {
  return eventLink(event) + '<br />' +
    '(' + (event.start.date || event.start.dateTime) + ') ' +
    '[' + eventActionLink(ACTION_MAKE_FREE, event) + '] [' + eventActionLink(ACTION_MAKE_BUSY, event) + ']';
}

function eventLink(event) {
  assert(event, "No Event");
  assert(event.htmlLink, "No event.htmlLink", event);
  assert(event.summary, "No event.summary", event);
  return '<a ' +
      'href="' +  event.htmlLink + '" ' +
      'data-calendar-id="' + event.calendarId + '" data-event-id="' + event.id + '"' +
    '>' + 
      event.summary +
    '</a>';
}

function eventActionLink(action, event) {
  return '<a href="' + eventActionURL(action, event) + '">' +
      'make <em>' + (action == ACTION_MAKE_FREE ? 'Free' : 'Busy') + '</em>' +
    '</a>';
}

function md5(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce(function(str,chr){
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length==1 ? '0' : '') + chr;
  },'');
}

function sendEmail(body, headers) {
  var msg = Gmail.newMessage();
  const raw =
      Object.keys(headers).reduce(function (acc, k, ii) {
        const v = headers[k]; const e = k + ": " + v; return ii === 0 ? e : acc + '\r\n' + e
      }, '') +
      '\r\n\r\n' +
      body;
  msg.raw = Utilities.base64EncodeWebSafe(raw);
  Gmail.Users.Messages.send(msg, 'me');
}

function getProperty(property_key, error_if_missing) {
  const property = PropertiesService.getScriptProperties().getProperty(property_key);
  if (error_if_missing && !property) {
    throw missingPropertyErrorMessage(property);
  } else {
    return property;
  }
}

function nagPropertyKey(event) {
  return NAG_PROPERTY_PREFIX + event.id.toString();
}

function assert() { if (!arguments[0]) { throw 'Error: ' + JSON.stringify(arguments); } }



function _test() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties();
  
  Logger.log(properties);
}