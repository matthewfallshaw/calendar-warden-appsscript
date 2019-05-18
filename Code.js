'use strict';

var DEFAULT_CALENDAR = PropertiesService.getScriptProperties().getProperty('defaultCalendar');
var SELF_LINK = PropertiesService.getScriptProperties().getProperty('selfLink');
var ACTION_MAKE_FREE = 'makefree';
var ACTION_MAKE_BUSY = 'makebusy';
var NAG_PROPERTY_PREFIX = 'NAG:';


/**
 * Web App server. Handle GET requests (from links in generated emails).
 */
function doGet(request) {  // Should really be a POST, but that's hard to do from an email
  var params = request.parameter;

  if (!validParams(params)) { return invalidParamsResponse(request); }

  var event = Calendar.Events.get(params.calendarId, params.eventId);
  if (!event) { return eventNotFoundResponse(request); }

  try {
    var report = makeEventBe({ action: params.action, calendarId: params.calendarId }, event);
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
  var event = report.event;
  var message = report.message;
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
  var output = template.evaluate();
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
    response.items.calendarId = calendarId;
    return response.items.map(function(event) { event.calendarId = calendarId; return event });
  } else {
    return [];
  }
}

/**
 * Formatted summary of a calendar event.
 * (String)
 */
function eventSummary(event) {
  return '  <li><a href="' +  event.htmlLink + '">' + event.summary + '</a><br />(' +
    (event.start.date || event.start.dateTime) +
     ')' +
     ' [<a href="' + eventActionURL(ACTION_MAKE_FREE, event) + '">make Free</a>]' +
     ' [<a href="' + eventActionURL(ACTION_MAKE_BUSY, event) + '">make Busy</a>]' +
     '</li>\n';
}

/**
 * Email a list of events, unless that list is empty.
 * (Returns the event list for chaining.)
 */
function sendEmailListingEvents(intro_text, events) {
  if (events && (events.length > 0)) {
    const body = intro_text +
      '\n\n<ul>\n' +
      events.reduce(function(acc, event) {
        return acc + eventSummary(event);
      }, '') +
      '</ul>';
    const thread_id =
        md5(events.reduce(function(acc, e, ii) {return ii === 0 ? e.id : acc + '|' + e.id}, '')) +
          '@' + Session.getActiveUser().getEmail().replace(/@.*/,'');
    sendEmail(
      body,
      {
        To: Session.getActiveUser().getEmail(),
        From: Session.getActiveUser().getEmail(),
        Subject: 'Calendar blocking entries',
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
      throw 'Unrecognised action `' + action.action + '`; I know how to `' + ACTION_MAKE_FREE + '` and `' + ACTION_MAKE_BUSY + '`. ' +
            json_out(action) + ' | ' + event;
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


/**
 * Check calendar.
 */
function checkCalendar(calendarId) {
  const now = new Date().getTime()
  const three_days_ago = now - 1000 * 60 * 60 * 24 * 3;
  var scriptProperties = PropertiesService.getScriptProperties()
  var properties = scriptProperties.getProperties();

  const events = upcomingEvents(calendarId);

  removeAlertsFromBlockEvents(
    events.filter(function(event) {
      return event.summary && event.summary.match(/^block$/i);
    })
  );

  // Email list of noncomplying & blocking events in default calendar.
  if (calendarId === DEFAULT_CALENDAR) {
    var events_after_autos = events.filter(function (event) {
        const last_nagged = properties[nagPropertyKey(event)];
        const recently_nagged = last_nagged && (last_nagged > three_days_ago);
        const marked_free = (event.transparency == 'transparent');
        const not_tagged_busy = (!event.description || !event.description.match(/\[busy\]/));
        return !marked_free && not_tagged_busy && !recently_nagged;
      }).filter(function(event) {
        // Auto-processible events
        if (event.description && event.description.match(/tram home$|tram to|practice$|climbing$/) && !event.description.match(/matt|fallabria/i)) {
          try {
            makeEventBe({ action: ACTION_MAKE_FREE, calendarId: calendarId }, event);
            return false;
          } catch(_) {
            return true;
          }
        } else {
          return true;
        }
      })
    var email_message = "<p>The following calendar events are blocking your availability, but don't contain " +
      '<code>[busy]</code> in their descriptions. Either make them <em>Free</em> or add ' +
      "<code>[busy]</code> to their descriptions (or I'll keep bugging you about them)."
    sendEmailListingEvents(email_message, events_after_autos).
    map(function (event) {  // update last nagged timestamps
      scriptProperties.setProperty(nagPropertyKey(event), now);
      return event;
    });
  }
}

/**
 * Check calendars.
 */
function checkCalendars() {
  if (!DEFAULT_CALENDAR) { throw missingPropertyErrorMessage('defaultCalendar') }
  checkCalendar(DEFAULT_CALENDAR);
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
 * Clear event nag timestamp properties.
 * A periodic cleanup function run by a timed trigger.
 */
function clearOldNagTimestamps() {
  const three_days_ago = new Date().getTime() - 1000 * 60 * 60 * 24 * 3;
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties();
  Object.keys(properties)
    .filter(function(key) { return key.match(new RegExp('^' + NAG_PROPERTY_PREFIX)) })
    .map(function (key) {
      if (!isNaN(Number(properties[key])) && (Number(properties[key])) < three_days_ago) { scriptProperties.deleteProperty(key) }
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
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function eventActionURL(action, event) {
  if (!SELF_LINK) { throw missingPropertyErrorMessage('selfLink') }
  return SELF_LINK + '?calendarId=' + event.calendarId + '&eventId=' + event.id + '&action=' + action
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

function nagPropertyKey(event) {
  return NAG_PROPERTY_PREFIX + event.id.toString();
}



function test() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties();
  
  Logger.log(Session.getActiveUser().getEmail());
}