// Absence report calendars, that follow the format: `FULL NAME - REASON`
const SOURCE_CALENDARS = {
  '[Absence]': 'ADD_CALENDAR_ID_HERE',
};

// Team members whose absence we want to get notified of
const PEOPLE = [
  { name: "Budavári Mátyás", nick: "Mátyás", team: "M" },
  // Add more users below
]

// Who should get these emails
const NOTIFICATION_MAIL_ADDRESSES = [
  "budavariam@gmail.com",
]

// Time period length
const DAYS_TO_SYNC = 5;
const EVENT_END_DATE_SHOULD_BE_INCLUDED = false
// Prefix of mail subject
const EMAIL_SUBJECT = `Absence report`
// Name that should appear as sender
const SENDER_NAME = "Absence Report Bot"

// ----------------------------------------------------------------------------
// DO NOT TOUCH FROM HERE ON
// ----------------------------------------------------------------------------
const NAME_LIST_MAP = PEOPLE.reduce((acc, curr, i) => { acc[curr.name] = i; return acc }, {})
const TEAM_LIST = Object.keys(PEOPLE.reduce((acc, curr) => { acc[curr.team] = 1; return acc }, {})).sort()

/** Print Date object as string in YYYY-MM-DD format */
function dateToYMD(date) {
  const y = date.getFullYear();
  return '' + y + '-' + dateToMD(date);
}

/** Print Date object as string in MM-DD format */
function dateToMD(date) {
  const d = date.getDate();
  const m = date.getMonth() + 1; //Month from 0 to 11
  return '' + (m <= 9 ? '0' + m : m) + '-' + (d <= 9 ? '0' + d : d);
}

/** Creates a new Date object some days later of earlier */
function dayDiff(start, delta) {
  const selectedDate = new Date(start.valueOf());
  selectedDate.setDate(selectedDate.getDate() + delta);
  return selectedDate
}

/** Tells whether a date is between 2 other dates, tested with YYYY-MM-DD format */
function betweenDateStrings(checkDateStr, fromDateStr, toDateStr) {
  const f = new Date(fromDateStr)
  const c = new Date(checkDateStr)
  const t = new Date(toDateStr)
  return (c >= f) && (c <= t)
}

/** In our calendar endDate is nonInclusive, but we might need it */
function getEndDate(endDate) {
  return dateToYMD(dayDiff(endDate, EVENT_END_DATE_SHOULD_BE_INCLUDED ? 0 : -1))
}

/** Return the list of users with their absence records for the next time period */
function getAbsenceEvents(startTime, endTime) {
  const teamAbsence = []
  for (let calendarName in SOURCE_CALENDARS) {
    const calendarId = SOURCE_CALENDARS[calendarName];
    const calendarToCopy = CalendarApp.getCalendarById(calendarId);

    if (!calendarToCopy) {
      console.log("Calendar not found: '%s'.", calendarId);
      continue;
    }

    // Find events
    const events = Calendar.Events.list(calendarId, {
      timeMin: startTime.toISOString(),
      timeMax: endTime.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
    });

    // If nothing find, move to next calendar
    if (!(events.items && events.items.length > 0)) {
      continue;
    }
    events.items.forEach((event) => {
      // The event name format is: 'LastName Firstname - Reason'
      const name = event.summary.split(" - ")[0]
      if (name in NAME_LIST_MAP) {
        // Only get those users who we listed in NAME_LIST_MAP
        const person = PEOPLE[NAME_LIST_MAP[name]]
        const record = {
          name: person.name,
          nick: person.nick,
          team: person.team,
          start: event.start.date,
          end: getEndDate(event.end.date),
          absences: [],
        }
        for (let i = 0; i < DAYS_TO_SYNC; i++) {
          const curr = dateToYMD(dayDiff(startTime, i))
          const abscent = betweenDateStrings(curr, record.start, record.end)
          console.log("Debug:", curr, record)
          record.absences[i] = abscent
          if (abscent) {
            console.log(`${record.nick} is on vacation on ${curr}`)
          }
        }
        // TODO: should have only one record for users with multiple absences in a single time period
        teamAbsence.push(record)
      }
    });
  }
  console.log("Calculated Report:", teamAbsence)
  return teamAbsence;
}

/** get the HTML content of the email message about the absence report */
function formatEmail(startTime, absences) {
  let description = []
  if (absences.length === 0) {
    description = [
      "The team is fully operating in this time period!"
    ]
  } else {
    const header = ["Name"]
    for (let i = 0; i < DAYS_TO_SYNC; i++) {
      const curr = dateToMD(dayDiff(startTime, i))
      header.push(curr)
    }

    htmlTable = [
      `<table border="1">`,
      `<thead>`,
      `  <tr>${header.map(item => `<td>${item}</td>`).join("")}</tr>`,
      `</thead>`,
      `<tbody>`,
      absences.map(line => [
        "<tr>",
        `<td>${line.nick}</td>`,
        line.absences.map((isAbscent) => `<td style="background-color:${isAbscent ? "yellow" : "green"}">${isAbscent ? "." : ""}</td>`).join(""),
        "</tr>"
      ].join("")).join(""),
      `</tbody>`,
      `</table>`,
    ].join("")

    description = [
      "Here is the team absence report for the next time period.",
      "Yellow is vacation, Green is back on track.",
      htmlTable,
    ]
  }

  return [
    "Hi!",
    description.map(line => `<p>${line}</p>`).join(""),
    `Take care,<br/>${SENDER_NAME}`,
  ].join("")
}

/** Send email notifications */
function sendEmailNotifications(startTime, endTime, absences) {
  const message = formatEmail(startTime, absences)
  console.log("Message", message)
  const subject = EMAIL_SUBJECT + ` (${dateToYMD(startTime)} - ${getEndDate(endTime)})`
  console.log(`You can send ${MailApp.getRemainingDailyQuota()} mails today.`)
  NOTIFICATION_MAIL_ADDRESSES.forEach(emailAddress => {
    MailApp.sendEmail({
      to: emailAddress,
      name: SENDER_NAME,
      subject: subject,
      htmlBody: message,
    });
  })
}

/** App entry point */
function SendAbsenceReport() {
  const startTime = new Date();
  startTime.setHours(0, 0, 0, 0);
  const endTime = dayDiff(startTime, DAYS_TO_SYNC)
  const absences = getAbsenceEvents(startTime, endTime)
  sendEmailNotifications(startTime, endTime, absences)
}