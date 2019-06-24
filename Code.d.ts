/// <reference types="google-apps-script" />
declare const ACTION_MAKE_FREE: Action;
declare const ACTION_MAKE_BUSY: Action;
declare const NAG_PROPERTY_PREFIX = "_NAG:";
declare const TAG_MAKE_FREE = "free-calendar-event";
declare const TAG_MAKE_BUSY = "busy-calendar-event";
declare const TAG_MADE_FREE = "calendar-event-made-free";
declare const TAG_MADE_BUSY = "calendar-event-made-busy";
interface IRequest {
    queryString: string;
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
interface IProperties {
    [key: string]: string;
}
interface IEventWithCalendarId extends GoogleAppsScript.Calendar.Schema.Event {
    calendarId: string;
    id: string;
    description?: string;
    htmlLink?: string;
    reminders?: GoogleAppsScript.Calendar.Schema.EventReminders;
    start?: GoogleAppsScript.Calendar.Schema.EventDateTime;
    summary?: string;
    transparency?: string;
}
declare class EventWithCalendarId implements IEventWithCalendarId {
    calendarId: string;
    id: string;
    description?: string;
    htmlLink?: string;
    reminders?: GoogleAppsScript.Calendar.Schema.EventReminders;
    start?: GoogleAppsScript.Calendar.Schema.EventDateTime;
    summary?: string;
    transparency?: string;
    constructor(calendarId: string, event: GoogleAppsScript.Calendar.Schema.Event);
}
declare type Action = 'makefree' | 'makebusy';
/**
 * Web App server. Handle GET requests (from links in generated emails).
 *
 * @param {IRequest} request
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
declare function doGet(request: IRequest): GoogleAppsScript.HTML.HtmlOutput;
declare function validParams(params: IRequestParameters): boolean;
declare function responseInvalidParams(request: IRequest): GoogleAppsScript.HTML.HtmlOutput;
declare function responseEventNotFound(request: IRequest): GoogleAppsScript.HTML.HtmlOutput;
declare function responseError(request: IRequest, event: EventWithCalendarId, err: Error): GoogleAppsScript.HTML.HtmlOutput;
declare function responseSuccess(report: IReport): GoogleAppsScript.HTML.HtmlOutput;
declare function standardTemplate(): GoogleAppsScript.HTML.HtmlTemplate;
declare function responseStandard(template: GoogleAppsScript.HTML.HtmlTemplate): GoogleAppsScript.HTML.HtmlOutput;
declare function googleCalendarFavicon(): string;
/**
 * The next 50 upcoming events in a calendar.
 *
 * @param {string} calendarId
 * @returns {EventWithCalendarId[]}
 */
declare function upcomingEvents(calendarId: string): EventWithCalendarId[];
declare function eventsEmailSubject(events: EventWithCalendarId[]): string;
declare function eventsEmailBody(introText: string, events: EventWithCalendarId[]): string;
declare function eventsEmailThreadId(events: EventWithCalendarId[]): string;
/**
 * Email a list of events, unless that list is empty.
 *
 * @param {string} introText
 * @param {EventWithCalendarId[]} events
 * @returns {EventWithCalendarId[]}
 */
declare function sendEmailListingEvents(introText: string, events: EventWithCalendarId[]): EventWithCalendarId[];
/**
 * Remove alerts from events.
 *
 * @param {EventWithCalendarId[]} events
 * @returns {EventWithCalendarId[]}
 */
declare function removeAlerts(events: EventWithCalendarId[]): EventWithCalendarId[];
/**
 * Change state of event.
 *
 * @param {Action} action
 * @param {EventWithCalendarId} event
 * @returns {IReport}
 */
declare function makeEventBe(action: Action, event: EventWithCalendarId): IReport;
declare function makeEventBeFree(event: EventWithCalendarId): IReport;
declare function makeEventBeBusy(event: EventWithCalendarId): IReport;
declare function is_auto_processible_event(event: EventWithCalendarId): boolean;
declare function is_block_event(event: EventWithCalendarId): boolean;
/**
 * checkAndFixCalendar
 *
 * @param {string} calendarId
 * @returns {EventWithCalendarId[]}
 */
declare function checkAndFixCalendar(calendarId: string): EventWithCalendarId[];
/**
 * checkCalendars
 *
 * @returns {void}
 */
declare function checkCalendars(): void;
/**
 * Check calendar, triggered from Calendar change event.
 *
 * @param {GoogleAppsScript.Events.CalendarEventUpdated} data
 * @returns {EventWithCalendarId[]}
 */
declare function checkCalendarFromChangeEvent(data: GoogleAppsScript.Events.CalendarEventUpdated): EventWithCalendarId[];
/**
 * Check email account for flagged events.
 *
 * @returns {void}
 */
declare function checkEmailAccount(): void;
declare function checkEmail(): void;
/**
 * Clear event nag timestamp properties.
 * A periodic cleanup function run by a timed trigger.
 *
 * @returns {void}
 */
declare function clearOldNagTimestamps(): void;
declare function missingPropertyErrorMessage(property: string): string;
declare function json_out(json: {}): string;
declare function include(filename: string): string;
declare function eventActionURL(action: string, event: EventWithCalendarId): string;
declare function eventSummary(event: EventWithCalendarId): string;
declare function eventLink(event: EventWithCalendarId): string;
declare function eventActionLink(action: string, event: EventWithCalendarId): string;
declare function md5(str: string): string;
declare function sendEmail(body: string, headers: {
    [key: string]: string;
}): void;
declare function getProperty(propertyKey: string, errorIfMissing?: boolean): string;
declare function nagPropertyKey(event: EventWithCalendarId): string;
declare function withCalendarId(calendarId: string, event: GoogleAppsScript.Calendar.Schema.Event | undefined): EventWithCalendarId;
declare function withCalendarIds(calendarId: string, events: GoogleAppsScript.Calendar.Schema.Event[]): EventWithCalendarId[];
declare function getCalendarEvent(calendarId: string, eventId: string): EventWithCalendarId;
declare function listCalendarEvents(calendarId: string): EventWithCalendarId[];
declare function patchCalendarEvent(event: EventWithCalendarId): EventWithCalendarId;
declare function assert(test: boolean, ...rest: any[]): void | Error;
declare const workaroundAssign: (target: any, ...source: any) => any;
//# sourceMappingURL=Code.d.ts.map