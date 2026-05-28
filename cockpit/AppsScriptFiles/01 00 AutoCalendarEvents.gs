/**
 * =================================================================
 * Main function to be triggered by the button in the sheet.
 * Processes all calendar requests from the "Meetings" sheet.
 * =================================================================
 */
function processCalendarRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Meetings");
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: "Meetings" sheet not found.');
    return;
  }

  // Adjust range to exclude the final "instruction" row
  const dataRange = sheet.getRange(3, 1, sheet.getLastRow() - 3, sheet.getLastColumn());
  const data = dataRange.getDisplayValues();
  const ui = SpreadsheetApp.getUi();
  let summary = { created: 0, updated: 0, deleted: 0, errors: 0 };

  const COL = {
    NAME: 0, START_DATE: 1, START_TIME: 2, END_DATE: 3, END_TIME: 4,
    ORD_1ST: 5, ORD_2ND: 6, ORD_3RD: 7, ORD_4TH: 8,
    MO: 9, TU: 10, WE: 11, TH: 12, FR: 13,
    EVERY: 14, UNTIL: 15,
    INVITE_START: 16,
    ADDL_INVITEES: 23,
    MEETING_ID: 24,
    UPDATE: 25,
    DELETE: 26
  };

  const recipientMap = getRecipientMap();
  if (!recipientMap) return;

  // Loop backwards to prevent errors when deleting rows.
  for (let i = data.length - 1; i >= 0; i--) {
    const rowData = data[i];
    const rowNum = i + 3;

    if (!rowData[COL.NAME]) {
      // When looping backwards, we don't need to stop on an empty row,
      // just continue to the next one above it.
      continue;
    }

    const meetingId = rowData[COL.MEETING_ID];
    const isUpdate = rowData[COL.UPDATE] === 'TRUE';
    const isDelete = rowData[COL.DELETE] === 'TRUE';

    try {
      if (isDelete) {
        if (deleteEvent(rowData, rowNum, sheet, ui, COL)) {
          summary.deleted++;
        }
      } else if (isUpdate && meetingId) {
        if (updateEvent(rowData, rowNum, sheet, recipientMap, COL)) {
          summary.updated++;
        }
      } else if (!meetingId) {
        if (createEvent(rowData, rowNum, sheet, recipientMap, COL)) {
          summary.created++;
        }
      }
    } catch (e) {
      Logger.log(`FATAL ERROR processing row ${rowNum}: ${e.message}\n${e.stack}`);
      sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).setBackground('#f4cccc');
      summary.errors++;
    }
  }

  SpreadsheetApp.flush();
  ui.alert('Processing Complete',
           `Summary:\n- ${summary.created} meetings created.\n- ${summary.updated} meetings updated.\n- ${summary.deleted} rows deleted.\n- ${summary.errors} rows had errors (marked in red).`,
           ui.ButtonSet.OK);
}


/**
 * =================================================================
 * Helper Functions
 * =================================================================
 */

function getRecipientMap() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const namedRange = ss.getRangeByName("CalendarRecipients");
    if (!namedRange) {
        SpreadsheetApp.getUi().alert('Error: Named range "CalendarRecipients" not found. Please create it in "Lookuptables!G2:H7".');
        return null;
    }
    const values = namedRange.getValues();
    const map = {};
    for (const row of values) {
        if (row[0] && row[1]) {
            map[row[0].toString().trim()] = row[1].toString().trim();
        }
    }
    return map;
}

function parseDateTime(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;

  const dateParts = dateStr.toString().split('-');
  const timeParts = timeStr.toString().split(':');

  if (dateParts.length !== 3 || timeParts.length !== 2) {
    Logger.log(`Invalid date/time format. Date: ${dateStr}, Time: ${timeStr}`);
    return null;
  }

  const year = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10) - 1;
  const day = parseInt(dateParts[2], 10);

  const hour = parseInt(timeParts[0], 10);
  const minute = parseInt(timeParts[1], 10);

  const dateObj = new Date(year, month, day, hour, minute);

  if (dateObj.getFullYear() !== year || dateObj.getMonth() !== month || dateObj.getDate() !== day) {
      Logger.log(`Invalid date components created from string: ${dateStr}`);
      return null;
  }

  return dateObj;
}

function getAttendees(rowData, recipientMap, COL) {
    let emails = [];
    const headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings").getRange(2, COL.INVITE_START + 1, 1, 6).getValues()[0];

    for (let i = 0; i < headers.length; i++) {
        const isChecked = rowData[COL.INVITE_START + i] === 'TRUE';
        if (isChecked && recipientMap[headers[i]]) {
            emails.push(recipientMap[headers[i]]);
        }
    }

    if (rowData[COL.ADDL_INVITEES]) {
        const additionalEmails = rowData[COL.ADDL_INVITEES].toString().split(',').map(email => email.trim()).filter(email => email);
        emails = emails.concat(additionalEmails);
    }

    return [...new Set(emails)];
}

function deleteEvent(rowData, rowNum, sheet, ui, COL) {
    const confirmation = ui.alert(
        'Confirm Deletion',
        `Are you sure you want to delete all future occurrences of "${rowData[0]}"? This action cannot be undone.`,
        ui.ButtonSet.YES_NO
    );

    if (confirmation !== ui.Button.YES) {
        sheet.getRange(rowNum, COL.DELETE + 1).setValue(false);
        return false;
    }

    const meetingId = rowData[COL.MEETING_ID];
    const recurrenceType = rowData[COL.EVERY];
    const calendar = CalendarApp.getCalendarById('c_d0b6c68ff4c233b0fb74c62c24951eaaefe8322b5a31af7290116e9c70d26228@group.calendar.google.com');

    try {
        if (!meetingId) {
            Logger.log(`No Meeting ID in row ${rowNum}. Assuming event never existed.`);
            // Fall through to delete the row.
        } else if (recurrenceType === 'Month') {
            const events = calendar.getEvents(new Date(2020, 0, 1), new Date(2099, 11, 31), {
                privateExtendedProperty: `customSeriesId=${meetingId}`
            });
            if (events.length > 0) {
                for (const event of events) {
                    event.deleteEvent();
                }
                Logger.log(`Successfully deleted ${events.length} monthly events for row ${rowNum}.`);
            } else {
                Logger.log(`Monthly events for row ${rowNum} were already gone.`);
            }
        } else { // Handles "Week", "Other Week", and single non-recurring events
            const event = calendar.getEventById(meetingId);
            if (event.getEventSeries()) {
                event.getEventSeries().deleteEventSeries();
                Logger.log(`Successfully deleted event series for row ${rowNum}.`);
            } else {
                event.deleteEvent();
                Logger.log(`Successfully deleted single event for row ${rowNum}.`);
            }
        }
    } catch (e) {
        const errorMessage = e.toString();
        // Check if the error is the specific one we want to ignore.
        if (errorMessage.includes('does not exist') || errorMessage.includes('already been deleted')) {
            Logger.log(`Event for row ${rowNum} was already deleted from calendar (error was caught and correctly handled).`);
        } else {
            // This is a different, unexpected error. We should stop and report it.
            Logger.log(`An unexpected error occurred while deleting event for row ${rowNum}: ${errorMessage}`);
            ui.alert(`An unexpected error occurred while deleting the event for "${rowData[0]}". Please check the logs.`);
            sheet.getRange(rowNum, COL.DELETE + 1).setValue(false);
            return false;
        }
    }

    // If we get here, the desired state (event is gone) is achieved. Delete the sheet row.
    Logger.log(`Proceeding to delete sheet row ${rowNum}.`);
    sheet.deleteRow(rowNum);
    return true;
}

function createEvent(rowData, rowNum, sheet, recipientMap, COL) {
    const title = rowData[COL.NAME];
    const startDate = parseDateTime(rowData[COL.START_DATE], rowData[COL.START_TIME]);

    const endDateStr = rowData[COL.END_DATE] ? rowData[COL.END_DATE] : rowData[COL.START_DATE];
    const endDate = parseDateTime(endDateStr, rowData[COL.END_TIME]);

    if (!title || !startDate || !endDate) {
        throw new Error(`Invalid input for row ${rowNum}. Check meeting name, dates, and times.`);
    }
    if (endDate < startDate) {
        throw new Error(`End date cannot be before the start date for row ${rowNum}.`);
    }

    const recurrenceType = rowData[COL.EVERY];
    const untilDateRaw = rowData[COL.UNTIL];
    const untilDate = untilDateRaw ? new Date(untilDateRaw) : null;
    const guests = getAttendees(rowData, recipientMap, COL);
    const options = { guests: guests.join(','), sendInvites: false };
    const calendar = CalendarApp.getCalendarById('c_d0b6c68ff4c233b0fb74c62c24951eaaefe8322b5a31af7290116e9c70d26228@group.calendar.google.com');
    let newId = '';

    if (recurrenceType === 'Week' || recurrenceType === 'Other Week') {
        const recurrence = CalendarApp.newRecurrence();
        const weekdays = [
            rowData[COL.MO] === 'TRUE' ? CalendarApp.Weekday.MONDAY : null,
            rowData[COL.TU] === 'TRUE' ? CalendarApp.Weekday.TUESDAY : null,
            rowData[COL.WE] === 'TRUE' ? CalendarApp.Weekday.WEDNESDAY : null,
            rowData[COL.TH] === 'TRUE' ? CalendarApp.Weekday.THURSDAY : null,
            rowData[COL.FR] === 'TRUE' ? CalendarApp.Weekday.FRIDAY : null,
        ].filter(Boolean);

        if (weekdays.length === 0) throw new Error(`No weekdays selected for weekly recurrence in row ${rowNum}.`);

        const weeklyRule = recurrence.addWeeklyRule().onlyOnWeekdays(weekdays);

        if (recurrenceType === 'Other Week') {
            weeklyRule.interval(2);
        }
        if (untilDate) {
            weeklyRule.until(untilDate);
        }

        const series = calendar.createEventSeries(title, startDate, endDate, recurrence, options);
        newId = series.getId();
        Logger.log(`Created event series for "${title}" with ID: ${newId}`);

    } else if (recurrenceType === 'Month') {
        const customSeriesId = Utilities.getUuid();
        options.privateExtendedProperty = { customSeriesId: customSeriesId };

        const weekdays = [];
        if (rowData[COL.MO] === 'TRUE') weekdays.push(1);
        if (rowData[COL.TU] === 'TRUE') weekdays.push(2);
        if (rowData[COL.WE] === 'TRUE') weekdays.push(3);
        if (rowData[COL.TH] === 'TRUE') weekdays.push(4);
        if (rowData[COL.FR] === 'TRUE') weekdays.push(5);

        if (weekdays.length !== 1) throw new Error(`For "Month" recurrence, exactly one weekday must be selected in row ${rowNum}.`);
        const weekday = weekdays[0];

        const ordinals = [
            rowData[COL.ORD_1ST] === 'TRUE' ? 1 : null,
            rowData[COL.ORD_2ND] === 'TRUE' ? 2 : null,
            rowData[COL.ORD_3RD] === 'TRUE' ? 3 : null,
            rowData[COL.ORD_4TH] === 'TRUE' ? 4 : null,
        ].filter(Boolean);

        if (ordinals.length === 0) throw new Error(`No ordinals (1st, 2nd, etc.) selected for monthly recurrence in row ${rowNum}.`);

        let loopUntilDate = untilDate ? new Date(untilDate) : new Date(new Date().getFullYear(), 11, 31);
        let currentMonthDate = new Date(startDate.getTime());
        const eventDuration = endDate.getTime() - startDate.getTime();
        let createdCount = 0;

        while (currentMonthDate <= loopUntilDate) {
            for (const n of ordinals) {
                const eventDate = getNthWeekdayOfMonth(currentMonthDate.getFullYear(), currentMonthDate.getMonth(), weekday, n);
                if (eventDate && eventDate >= startDate && eventDate <= loopUntilDate) {
                    const eventStart = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), startDate.getHours(), startDate.getMinutes());
                    const eventEnd = new Date(eventStart.getTime() + eventDuration);

                    calendar.createEvent(title, eventStart, eventEnd, options);
                    createdCount++;
                }
            }
            currentMonthDate.setMonth(currentMonthDate.getMonth() + 1);
        }

        newId = customSeriesId;
        Logger.log(`Created ${createdCount} individual events for "${title}" with custom series ID: ${newId}`);

    } else {
        const event = calendar.createEvent(title, startDate, endDate, options);
        newId = event.getId();
        Logger.log(`Created single event for "${title}" with ID: ${newId}`);
    }

    if (newId) {
        sheet.getRange(rowNum, COL.MEETING_ID + 1).setValue(newId);
        sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).setBackground(null);
        return true;
    }
    return false;
}

function updateEvent(rowData, rowNum, sheet, recipientMap, COL) {
    Logger.log(`Starting update for row ${rowNum}...`);
    const meetingId = rowData[COL.MEETING_ID];
    const recurrenceType = rowData[COL.EVERY];
    const calendar = CalendarApp.getCalendarById('c_d0b6c68ff4c233b0fb74c62c24951eaaefe8322b5a31af7290116e9c70d26228@group.calendar.google.com');

    try {
        if (recurrenceType === 'Month') {
            const events = calendar.getEvents(new Date(2020,0,1), new Date(2099, 11, 31), { privateExtendedProperty: `customSeriesId=${meetingId}`});
             if (events.length === 0) {
               Logger.log(`For update, no events found with customSeriesId: ${meetingId} for row ${rowNum}. Proceeding to create.`);
            }
            for (const event of events) {
                event.deleteEvent();
            }
        } else {
            const series = calendar.getEventSeriesById(meetingId);
            if (series) {
               series.deleteEventSeries();
            } else {
               Logger.log(`For update, event series with ID ${meetingId} from row ${rowNum} not found. Proceeding to create.`);
            }
        }
    } catch (e) {
        Logger.log(`Could not find/delete old event for update on row ${rowNum}. It may have been deleted. Error: ${e.message}`);
    }

    if (createEvent(rowData, rowNum, sheet, recipientMap, COL)) {
        sheet.getRange(rowNum, COL.UPDATE + 1).setValue(false);
        Logger.log(`Update successful for row ${rowNum}.`);
        return true;
    } else {
        Logger.log(`Update failed for row ${rowNum} during the creation phase.`);
        return false;
    }
}

function getNthWeekdayOfMonth(year, month, weekday, n) {
    const d = new Date(year, month, 1);

    while (d.getDay() !== weekday) {
        d.setDate(d.getDate() + 1);
    }

    d.setDate(d.getDate() + (n - 1) * 7);

    if (d.getMonth() !== month) {
        return null;
    }
    return d;
}
