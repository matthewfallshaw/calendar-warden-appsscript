const ACTION_MAKE_FREE: Action = 'makefree';
const ACTION_MAKE_BUSY: Action = 'makebusy';
const NAG_PROPERTY_PREFIX = '_NAG:';
const TAG_MAKE_FREE = 'free-calendar-event';
const TAG_MAKE_BUSY = 'busy-calendar-event';
const TAG_MADE_FREE = 'calendar-event-made-free';
const TAG_MADE_BUSY = 'calendar-event-made-busy';

interface IRequest {
  queryString: string;
  // parameter: {[key: string]: string; },
  parameter: IRequestParameters;
  contextPath: string;
  contentLength: number;
}
interface IRequestParameters {
  calendarId: string;
  eventId: string;
  action: string;
}
interface IResponseTemplate extends GoogleAppsScript.HTML.HtmlTemplate {
  status: string;
  message: string;
  detail: string;
}

interface IReport {
  event: EventWithCalendarId;
  message: string;
}

// TODO: This avoids implicit any warnings when working with PropertiesService, but ugly much?
interface IProperties {
  [key: string]: string;
}

interface IEventWithCalendarId extends GoogleAppsScript.Calendar.Schema.Event {
  calendarId: string;
  id: string;  // @types say this is optional; nuh-ah!
  // TODO: why are the following properties not auto pulled from GAS.C.S.Event?
  description?: string;
  htmlLink?: string;
  reminders?: GoogleAppsScript.Calendar.Schema.EventReminders;
  start?: GoogleAppsScript.Calendar.Schema.EventDateTime;
  summary?: string;
  transparency?: string;
}

class EventWithCalendarId implements IEventWithCalendarId {
  // TODO: is there an elegant way to avoid the duplication with the intnerface?
  calendarId: string;
  id: string;
  description?: string;
  htmlLink?: string;
  reminders?: GoogleAppsScript.Calendar.Schema.EventReminders;
  start?: GoogleAppsScript.Calendar.Schema.EventDateTime;
  summary?: string;
  transparency?: string;

  constructor(calendarId: string, event: GoogleAppsScript.Calendar.Schema.Event) {
    workaroundAssign(this, event);  // typescript happily emits Object.assign, which fails on GAS
    this.calendarId = calendarId;
    // TODO: make an `assert_type` wrapper for this pattern
    if (typeof event.id === 'string') {
      this.id = event.id; // TODO: already assigned by workaroundAssign, but this to cause typescript
                          // to see an assignment
    } else {
      throw new Error('Error: Event has no id. ' + JSON.stringify(event));
    }
  }
}

type Action = 'makefree' | 'makebusy';

/**
 * Web App server. Handle GET requests (from links in generated emails).
 *
 * @param {IRequest} request
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(request: IRequest): GoogleAppsScript.HTML.HtmlOutput {
  // Should really be a POST, but that's hard to do from an email
  const params: IRequestParameters = request.parameter;

  if (!validParams(params)) {return responseInvalidParams(request); }

  const event = getCalendarEvent(params.calendarId, params.eventId);
  if (!event) {return responseEventNotFound(request); }

  try {
    const report = makeEventBe(params.action as Action, event);
    return responseSuccess(report);
  } catch (err) {
    return responseError(request, event, err);
  }
}

function validParams(params: IRequestParameters): boolean  {
  return (params.calendarId && params.eventId && params.action && true || false);
}

function responseInvalidParams(request: IRequest): GoogleAppsScript.HTML.HtmlOutput  {
  let message = '<p>Missing parameters:</p>\n<ol>\n';
  if (!request.parameter.calendarId) {message += '<li>Missing calendarId.</li>\n'; }
  if (!request.parameter.eventId)    {message += '<li>Missing eventId.</li>\n'; }
  if (!request.parameter.action)     {message += '<li>Missing action.</li>\n'; }
  message += '</ol>';
  const template = standardTemplate() as IResponseTemplate;
  template.status = 'error';
  template.message = message;
  template.detail = json_out(request);
  return responseStandard(template);
}

function responseEventNotFound(request: IRequest): GoogleAppsScript.HTML.HtmlOutput  {
  const template = standardTemplate() as IResponseTemplate;
  template.status = 'error';
  template.message = '<p>Event not found.</p>\n';
  template.detail = json_out(request);
  return responseStandard(template);
}

function responseError(request: IRequest,
                       event: EventWithCalendarId,
                       err: Error): GoogleAppsScript.HTML.HtmlOutput {
  const template = standardTemplate() as IResponseTemplate;
  template.status = 'error';
  template.message = '<p>' + err + '</p>\n';
  template.detail = json_out({request, event, error: err});
  return responseStandard(template);
}

function responseSuccess(report: IReport): GoogleAppsScript.HTML.HtmlOutput  {
  const event = report.event;
  const message = '<p><a target="_parent" href="' + event.htmlLink + '">Event</a> ' +
    report.message + '.</p>\n';
  const template = standardTemplate() as IResponseTemplate;
  template.status = 'success';
  template.message = message;
  template.detail = json_out(event);
  return responseStandard(template);
}

function standardTemplate(): GoogleAppsScript.HTML.HtmlTemplate  {
  const template = HtmlService.createTemplateFromFile('Index');
  return template;
}

function responseStandard(template: GoogleAppsScript.HTML.HtmlTemplate): GoogleAppsScript.HTML.HtmlOutput  {
  const output = template.evaluate();
  output.setTitle('Calendar Warden');
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  output.setFaviconUrl(googleCalendarFavicon());
  return output;
}

function googleCalendarFavicon(): string  {
  return 'https://calendar.google.com/googlecalendar/images/favicon_v2014_' + (new Date()).getDate() + '.ico';
}

/**
 * The next 50 upcoming events in a calendar.
 *
 * @param {string} calendarId
 * @returns {EventWithCalendarId[]}
 */
function upcomingEvents(calendarId: string): EventWithCalendarId[]  {
  return listCalendarEvents(calendarId);
}

function eventsEmailSubject(events: EventWithCalendarId[]): string  {
  const subject = events.map((event) => {
    const match = event.summary && event.summary.match(/^.{0,25}\w{0,15}/);
    const summary = match && match[0] || 'Unnamed event';
    if (event.summary && event.summary.length || 0 > summary.length) {
      return summary + '…';
    } else {
      return summary;
    }
  }).join(' | ');
  return subject;
}

function eventsEmailBody(introText: string, events: EventWithCalendarId[]): string  {
  return introText +
    '\n\n<ul>\n' +
    events.reduce((acc, event) => {
      return acc + '  <li>' + eventSummary(event) + '</li>\n';
    }, '') +
    '</ul>';
}

function eventsEmailThreadId(events: EventWithCalendarId[]): string  {
  return md5(events.reduce(
    (acc: string, e: EventWithCalendarId, ii: number): string => ii === 0 ? e.id : acc + '|' + e.id,
    ''),
  ) + '@' + Session.getActiveUser().getEmail().replace(/@.*/, '');
}

/**
 * Email a list of events, unless that list is empty.
 *
 * @param {string} introText
 * @param {EventWithCalendarId[]} events
 * @returns {EventWithCalendarId[]}
 */
function sendEmailListingEvents(introText: string, events: EventWithCalendarId[]): EventWithCalendarId[]  {
  const body = eventsEmailBody(introText, events);
  const subject = eventsEmailSubject(events);
  const threadId = eventsEmailThreadId(events);
  const emailAddress = Session.getActiveUser().getEmail();
  sendEmail(
    body,
    {
      'Content-Type': 'text/html; charset=UTF-8',
      'From': emailAddress,
      'In-Reply-To': '<' + threadId + '>',
      'Subject': 'CalendarWarden: ' + subject,
      'To': emailAddress,
    },
  );
  return events;
}

/**
 * Remove alerts from events.
 *
 * @param {EventWithCalendarId[]} events
 * @returns {EventWithCalendarId[]}
 */
function removeAlerts(events: EventWithCalendarId[]): EventWithCalendarId[]  {
  return events.map((event) => {
    if (event.reminders) {
      event.reminders.useDefault = false;
      delete event.reminders;
    } else {
      event.reminders = {useDefault: false};
    }
    const patchedEvent = patchCalendarEvent(event);
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
function makeEventBe(action: Action, event: EventWithCalendarId): IReport  {
  let report: IReport;
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

function makeEventBeFree(event: EventWithCalendarId): IReport  {
  let message: string;
  let dirty = false;

  if (event.transparency === 'transparent') {
    message = 'already <em>free</em>';
  } else {
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

  const patchedEvent: EventWithCalendarId = dirty ? patchCalendarEvent(event) : event;

  const report: IReport = {event: patchedEvent, message};

  return report;
}

function makeEventBeBusy(event: EventWithCalendarId): IReport  {
  let message: string;
  let dirty = false;

  if (event.transparency === 'opaque') {
    message = 'already <em>busy</em>';
  } else {
    dirty = true;
    event.transparency = 'opaque';
    message = 'made <em>busy</em>';
  }
  if (event.description && event.description.match(/\[busy\]/)) {
    message += ' (<code>[busy]</code> already present in description)';
  } else {
    dirty = true;
    event.description = event.description ? event.description + '\n[busy]' : '[busy]';
  }

  const patchedEvent: EventWithCalendarId = dirty ? patchCalendarEvent(event) : event;
  const report: IReport = {event: patchedEvent, message};

  return report;
}

function is_auto_processible_event(event: EventWithCalendarId): boolean  {
  return (event.description &&
    event.description.match(/tram home$|tram to|practice$|climbing$/) &&
    !event.description.match(/matt|fallabria/i) && true || false);
}

function is_block_event(event: EventWithCalendarId): boolean  {
  return event.summary && event.summary.match(/^block$/i) && true || false;
}

/**
 * checkAndFixCalendar
 *
 * @param {string} calendarId
 * @returns {EventWithCalendarId[]}
 */
function checkAndFixCalendar(calendarId: string): EventWithCalendarId[]  {
  const now: number = new Date().getTime();
  const threeDays: number = new Date(1000 * 60 * 60 * 24 * 3).getTime();
  const threeDaysAgo: number = (new Date(now - threeDays)).getTime();

  const events = upcomingEvents(calendarId);

  removeAlerts(events.filter(is_block_event));

  // TODO:
  // addTravelTime(events.filter(is_appointment));

  // Email list of noncomplying & blocking events in default calendar.
  if (calendarId === getProperty('defaultCalendar')) {
    const eventsAfterAutos = events.filter((event) => {
      const lastNagged: number = (new Date(Number(getProperty(nagPropertyKey(event), false)))).getTime();
      const recentlyNagged: boolean = lastNagged && (lastNagged > threeDaysAgo) || false;
      const markedFree: boolean = (event.transparency === 'transparent');
      const taggedBusy: boolean = !(!event.description || !event.description.match(/\[busy\]/));
      return !markedFree && !taggedBusy && !recentlyNagged;
    }).filter((event) => {
      if (is_auto_processible_event(event)) {
        try {
          makeEventBe(ACTION_MAKE_FREE, event);
          return false;
        } catch (_) {
          return true;
        }
      } else {
        return true;
      }
    });
    if (eventsAfterAutos.length > 0) {
      const emailMessage = "<p>The following calendar events are blocking your availability, but don't contain " +
        '<code>[busy]</code> in their descriptions. Either make them <em>Free</em> or add ' +
        "<code>[busy]</code> to their descriptions (or I'll keep bugging you about them).";
      return sendEmailListingEvents(emailMessage, eventsAfterAutos)
        .map((event) => {  // update last nagged timestamps
          PropertiesService.getScriptProperties().setProperty(nagPropertyKey(event), now.toString());
          return event;
        });
    } else {
      return [];  // no email to send
    }
  } else {
    return [];
  }
}

/**
 * checkCalendars
 *
 * @returns {void}
 */
function checkCalendars(): void {
  checkAndFixCalendar(getProperty('defaultCalendar'));
}

/**
 * Check calendar, triggered from Calendar change event.
 *
 * @param {GoogleAppsScript.Events.CalendarEventUpdated} data
 * @returns {EventWithCalendarId[]}
 */
function checkCalendarFromChangeEvent(data: GoogleAppsScript.Events.CalendarEventUpdated): EventWithCalendarId[]  {
  // { calendarId=example@gmail.com, authMode=FULL, triggerUid=706555 }
  const calendarId = data.calendarId;
  return checkAndFixCalendar(calendarId);
}

/**
 * Check email account for flagged events.
 *
 * @returns {void}
 */
function checkEmailAccount(): void {
  [TAG_MAKE_FREE, TAG_MAKE_BUSY].map((tag) => {
    // Search for events with TAG_MAKE_FREE or TAG_MAKE_BUSY
    let label = GmailApp.getUserLabelByName(tag);
    if (!label) {label = GmailApp.createLabel(tag); }
    const threads = label.getThreads();
    if (!threads || threads.length === 0) {return false; }
    const tagDone = (tag === TAG_MAKE_FREE) ? TAG_MADE_FREE : TAG_MADE_BUSY;
    const labelDone = GmailApp.getUserLabelByName(tagDone) || GmailApp.createLabel(tagDone);
    // for all threads
    threads.map((thread) => {
      thread.getMessages().map((message) => {  // assume threads have >0 messages
        const re = /data-calendar-id="([^"]+)" +data-event-id="([^"]+)"/g;
        const body = message.getBody();
        // find events
        let matches = re.exec(body);
        while (matches !== null) {
          const calendarId = matches[1];
          const eventId = matches[2];
          const event = getCalendarEvent(calendarId, eventId);
          // make free or busy and replace with TAG_MADE_FREE|BUSY label
          if (event) {
            if (tag === TAG_MAKE_FREE) {
              makeEventBe(ACTION_MAKE_FREE, event);
              console.info({
                'From checkEmailAccount:':
                  `Event ${event.htmlLink} made free (from email subject:"${message.getSubject()}")`});
            } else {
              makeEventBe(ACTION_MAKE_BUSY, event);
              console.info({
                'From checkEmailAccount:':
                  `Event ${event.htmlLink} made busy (from email subject:"${message.getSubject()}")`});
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

function checkEmail(): void {
  checkEmailAccount();
}

/**
 * Clear event nag timestamp properties.
 * A periodic cleanup function run by a timed trigger.
 *
 * @returns {void}
 */
function clearOldNagTimestamps(): void {
  const threeDaysAgo: number = new Date().getTime() - 1000 * 60 * 60 * 24 * 3;
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties: IProperties = scriptProperties.getProperties() as IProperties;
  Object.keys(properties)
    .filter((key: string) => key.match(new RegExp('^' + NAG_PROPERTY_PREFIX)))
    .map((key: string) => {
      if (!isNaN(Number(properties[key])) && (Number(properties[key])) < threeDaysAgo) {
        scriptProperties.deleteProperty(key);
      }
    });
}

// ## Utility functions

function missingPropertyErrorMessage(property: string) {
  return 'No ' + property + ' script property: Set a calendar email address as the value for a `' + property +
    '` key as a "Script Property" (from the script editor on `script.google.com` > "Project properties" ' +
    '> "Script properties")';
}

function json_out(json: {}) {
  return JSON.stringify(json, null, 2);
}

function include(filename: string) {  // called from Index.html
  const content = HtmlService.createHtmlOutputFromFile(filename).getContent();
  return content;
}

function eventActionURL(action: string, event: EventWithCalendarId) {
  return getProperty('selfLink') + '?calendarId=' + event.calendarId + '&eventId=' + event.id + '&action=' + action;
}

function eventSummary(event: EventWithCalendarId) {
  return eventLink(event) + '<br />' +
    '(' + (event.start ? event.start.date || event.start.dateTime : 'no start date') + ') ' +
    '[' + eventActionLink(ACTION_MAKE_FREE, event) + '] [' + eventActionLink(ACTION_MAKE_BUSY, event) + ']';
}

function eventLink(event: EventWithCalendarId) {
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

function eventActionLink(action: string, event: EventWithCalendarId): string  {
  return '<a href="' + eventActionURL(action, event) + '">' +
    'make <em>' + (action === ACTION_MAKE_FREE ? 'Free' : 'Busy') + '</em>' +
    '</a>';
}

function md5(str: string): string  {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str)
    .reduce((outStr: string, chr: GoogleAppsScript.Byte) => {
      const char = (chr < 0 ? chr + 256 : chr).toString(16);
      return outStr + (char.length === 1 ? '0' : '') + char;
    }, '');
}

function sendEmail(body: string, headers: {[key: string]: string}): void  {
  const msg = Gmail.newMessage();
  const raw =
    Object.keys(headers).reduce((acc, k, ii) => {
      const v = headers[k]; const e = k + ': ' + v; return ii === 0 ? e : acc + '\r\n' + e;
    }, '') +
    '\r\n\r\n' +
    body;
  msg.raw = Utilities.base64EncodeWebSafe(raw);
  const gu = Gmail.Users as GoogleAppsScript.Gmail.Collection.UsersCollection;
  const gm = gu.Messages as GoogleAppsScript.Gmail.Collection.Users.MessagesCollection;
  gm.send(msg, 'me');
}

function getProperty(propertyKey: string, errorIfMissing = false): string  {
  const property = PropertiesService.getScriptProperties().getProperty(propertyKey);
  if (errorIfMissing && !property) {
    throw missingPropertyErrorMessage(propertyKey);
  } else {
    return property ? property : '';
  }
}

function nagPropertyKey(event: EventWithCalendarId): string  {
  return NAG_PROPERTY_PREFIX + event.id.toString();
}

function withCalendarId(calendarId: string,
                        event: GoogleAppsScript.Calendar.Schema.Event | undefined): EventWithCalendarId  {
  if (event) {
    return new EventWithCalendarId(calendarId, event);
  } else {
    throw new Error('Error: Missing Event!');
  }
}

function withCalendarIds(calendarId: string, events: GoogleAppsScript.Calendar.Schema.Event[]): EventWithCalendarId[]  {
  return events.map((event) => new EventWithCalendarId(calendarId, event));
}

function getCalendarEvent(calendarId: string, eventId: string): EventWithCalendarId  {
  const es = Calendar.Events as GoogleAppsScript.Calendar.Collection.EventsCollection;
  const gEvent = es.get(calendarId, eventId);
  if (typeof gEvent.id === 'string') {
    return withCalendarId(calendarId, gEvent);
  } else {
    throw new Error('Error: Event with no id!' + JSON.stringify(gEvent));
  }
}

function listCalendarEvents(calendarId: string): EventWithCalendarId[]  {
  const es = Calendar.Events as GoogleAppsScript.Calendar.Collection.EventsCollection;
  const gEvents = es.list(calendarId, {
    maxResults: 50,
    orderBy: 'startTime',
    singleEvents: true,
    timeMin: (new Date()).toISOString(),
  });
  if (gEvents.items) {
    return withCalendarIds(calendarId, gEvents.items);
  } else {
    return [];
  }
}

function patchCalendarEvent(event: EventWithCalendarId): EventWithCalendarId  {
  const es = Calendar.Events as GoogleAppsScript.Calendar.Collection.EventsCollection;
  return withCalendarId(event.calendarId, es.patch(event, event.calendarId, event.id as string));
}

// TODO: nice to have a version of this for type assertions
function assert(test: boolean, ...rest: any[]): void | Error  {
  if (!test) {
    throw new Error('Error: ' + JSON.stringify(rest));
  }
}

const workaroundAssign = function(has) {
  'use strict';
  return assign;
  function assign(target: any, ...source: any) {
    for (const argument of source) {
      copy(target, argument);
    }
    return target;
  }
  function copy(target: any, source: any) {
    for (const key in source) {
      if (has.call(source, key)) {
        target[key] = source[key];
      }
    }
  }
}({}.hasOwnProperty);
