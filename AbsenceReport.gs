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
const EMAIL_SUBJECT = `Absence Report`
// Name that should appear as sender
const SENDER_NAME = "Absence Report Bot"
const EXPORT_TO_DRIVE_PATH = null;
// const EXPORT_TO_DRIVE_PATH = "absences.json"
const SEND_MAIL = true

// ----------------------------------------------------------------------------
// DO NOT TOUCH FROM HERE ON
// ----------------------------------------------------------------------------
const DEBUG_DAY_CNT = 0; // to see how things would be n days ago
const NAME_LIST_MAP = PEOPLE.reduce((acc, curr, i) => { acc[curr.name] = i; return acc }, {})
const TEAM_LIST = Object.keys(PEOPLE.reduce((acc, curr) => { acc[curr.team] = 1; return acc }, {})).sort()
let debugNames = []
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

  return EVENT_END_DATE_SHOULD_BE_INCLUDED
    ? (c >= f) && (c <= t) // should include end day
    : (c >= f) && (c < t) // should not include end day
}

/** In our calendar endDate is nonInclusive, but we might need it */
function getEndDate(endDate) {
  return dateToYMD(dayDiff(endDate, EVENT_END_DATE_SHOULD_BE_INCLUDED ? 0 : -1))
}

function getEvents(calendarId, startTime, endTime) {
  let eventList = []
  let pageToken = null
  while (pageToken !== undefined) {
    const events = Calendar.Events.list(calendarId, {
        timeMin: startTime.toISOString(),
        timeMax: endTime.toISOString(),
        singleEvents: true,
        maxResults: 2000,
        orderBy: 'startTime',
        pageToken: pageToken, 
      });
    console.log(`Got ${events.items.length} events.`)
    eventList = eventList.concat(events.items)
    if ("nextPageToken" in events) {
      pageToken = events.nextPageToken
    } else {
      break;
    }
  }
  return {items: eventList}
}

/** Return the list of users with their absence records for the next time period */
function getAbsenceEvents(startTime, endTime, sync_day_cnt) {
  const teamAbsence = {}
  for (let calendarName in SOURCE_CALENDARS) {
    const calendarId = SOURCE_CALENDARS[calendarName];
    const calendarToCopy = CalendarApp.getCalendarById(calendarId);

    if (!calendarToCopy) {
      console.log("Calendar not found: '%s'.", calendarId);
      continue;
    }

    // Find events
    const events = getEvents(calendarId, startTime, endTime)

    if (EXPORT_TO_DRIVE_PATH) {
      const evts = JSON.stringify(events)
      console.log(`File saved to ${EXPORT_TO_DRIVE_PATH}`)
      DriveApp.createFile(EXPORT_TO_DRIVE_PATH,evts);
    }
    // If nothing is found, move to next calendar
    if (!(events.items && events.items.length > 0)) {
      continue;
    }
    events.items.forEach((event) => {
      // The event name format is: 'LastName Firstname - Reason'
      const name = event.summary.split(" - ")[0]
      debugNames.push(name)
      if (name in NAME_LIST_MAP) {
        // Only get those users who we listed in NAME_LIST_MAP
        const person = PEOPLE[NAME_LIST_MAP[name]]
        const record = teamAbsence[person.name]
          ? teamAbsence[person.name]
          : {
            ...person,
            absences: [],
          }
        for (let i = 0; i < sync_day_cnt; i++) {
          const curr = dateToYMD(dayDiff(startTime, i))
          const isAbsent = betweenDateStrings(curr, event.start.date, event.end.date)
          console.log("Debug:", curr, record)
          record.absences[i] = record.absences[i]
            ? record.absences[i]
            : isAbsent
          if (isAbsent) {
            console.log(`${record.nick || record.name} is on vacation on ${curr}`)
          }
        }
        teamAbsence[person.name] = record
      }
    });
  }
  console.log("Calculated Report:", teamAbsence)
  return Object.values(teamAbsence);
}

/** Get the HTML content of the email message about the absence report */
function formatEmail(startTime, absences) {
  let description = []
  if (absences.length === 0) {
    description = [
      "Everyone is on board in this time period!"
    ]
  } else {
    const header = ["Name"]
    for (let i = 0; i < DAYS_TO_SYNC; i++) {
      const curr = dateToMD(dayDiff(startTime, i))
      header.push(curr)
    }

    const htmlTableBlock = [undefined].concat(TEAM_LIST).map((currentTeamName) => {
      const memberList = absences.filter(member => member.team === currentTeamName)
      if (memberList.length === 0) {
        return ""
      } 
      return [
        `<p>`,
        `<h3>${currentTeamName || "Team"}</h3>`,
        `<table border="1">`,
        `<thead>`,
        `  <tr>${header.map(item => `<td>${item}</td>`).join("")}</tr>`,
        `</thead>`,
        `<tbody>`,
        memberList.map(line => [
          "<tr>",
          `<td>${line.nick || line.name}</td>`,
          line.absences.map((isAbsent) => `<td style="background-color:${isAbsent ? "yellow" : "green"}">${isAbsent ? "." : ""}</td>`).join(""),
          "</tr>"
        ].join("")).join(""),
        `</tbody>`,
        `</table>`,
        `</p>`
      ].join("")
    }).join("")

    description = [
      "Here is the team absence report for the next time period.",
      "Yellow is vacation, Green is available.",
      htmlTableBlock,
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
  if (!SEND_MAIL) {
    console.log("Skip mail sending as per config")
    return
  }
  const message = formatEmail(startTime, absences)
  console.log("Message", message)
  const subject = `${EMAIL_SUBJECT} (${dateToYMD(startTime)} - ${getEndDate(endTime)})`
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

function printReport() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sync_day_cnt = 20
  const startTime = dayDiff(today, 10)
  const endTime = dayDiff(today, sync_day_cnt)
  const absences = getAbsenceEvents(startTime, endTime, sync_day_cnt)
  console.log(absences)
  const header = ["Name"]
  for (let i = 0; i < sync_day_cnt; i++) {
    const curr = dateToMD(dayDiff(startTime, i))
    header.push(curr)
  }
  const data = [].concat(TEAM_LIST).map((currentTeamName) => {
      const memberList = absences.filter(member => member.team === currentTeamName)
      if (memberList.length === 0) {
        return ""
      } 
      return [
        `## Team ${currentTeamName || "Team"}`,
        `${header.map(item => `${item} | `).join("")}`,
        `${header.map(item => `---- | `).join("")}`,
        ...memberList.map(line => [
          `${line.nick || line.name} |`,
          line.absences.map((isAbsent) => `${isAbsent ? "-" : ""}|`).join(""),
        ].join(""))
      ]
    }).filter(e => e)
    console.log(data)
    console.log(debugNames)
}

/** App entry point */
function SendAbsenceReport() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const startTime = dayDiff(today, -DEBUG_DAY_CNT)
  const endTime = dayDiff(startTime, DAYS_TO_SYNC)
  const absences = getAbsenceEvents(startTime, endTime, DAYS_TO_SYNC)
  sendEmailNotifications(startTime, endTime, absences)
}
