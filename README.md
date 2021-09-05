# Absence Report

I have access to google calendar, that contains the absence of all my colleagues as full day events.
I work with a mostly remote team, and sometimes it's useful to know when will my teammates leave for vacation,
even if they don't tell me many days in advance.

I wrote a simple script, to look through the events every monday and filter the ones I'm interested in, then send me an email report.

## Getting started

Create a new project in [Google Apps Scripts](https://script.google.com/)

- Fill the config values until the `DO NOT TOUCH FROM HERE ON` part.
  - Set the report calendar ID(s)
  - Set the recipients email address, that should get this report
  - Change any other values
- Enable `Show "appsscript.json" manifest file in editor` setting and change the `appscript.json` contents
- Run it once to be able to set the appropriate permissions
- Set up a trigger for your needs. I set up a report 5 days in advance and trigger on every monday morning

## APIs used

- `calendar` for listing events. [Docs](https://developers.google.com/calendar/api/v3/reference)
- `script.send_mail` to send emails. [Docs](https://developers.google.com/apps-script/reference/mail/mail-app)