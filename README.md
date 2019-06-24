# calendar-warden-appsscript
A Google Apps Script that monitors a shared calendar, making it easier to manage

## Use

If events don't conform to your checks you'll get emails. The emails will contain action links that
allow you to fix the brokennes.

## Config

``` sh
clasp login
clasp create
clasp push
clasp open
```
File > Project properties > Script properties > + Add row >

    defaultCalendar: <default calendar to use>
    selfLink: <https://script.google.com/macros/s/script id from `clasp open`/dev>
