function createTimeDrivenTriggers() {

  // 1. Reset
  deleteAllTriggers();

  // --- PUEReport (Kept at 09:00 window) ---
  // If you want this at 9:30, add .nearMinute(30) here too.
  ScriptApp.newTrigger('PUEReport')
      .timeBased().onMonthDay(1).atHour(9).create();

  // --- WeatherData (Kept at 09:00 window) ---
  // If you want this at 9:30, add .nearMinute(30) here too.
  ScriptApp.newTrigger('UpdateWeatherData')
      .timeBased().onMonthDay(1).atHour(9).create();

  // --- DCO Tickets (Kept at 09:00 window on the second - n8n runs on the first) ---
    ScriptApp.newTrigger('UpdateTA')
      .timeBased().onMonthDay(2).atHour(9).create();

  // --- FTRProd (Thursday 08:45) ---
  ScriptApp.newTrigger('FTRProd')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.THURSDAY)
      .atHour(8)
      .nearMinute(45)   // <--- Added this line
      .create();

  // --- FTRTesting (Monday 08:45) ---
  ScriptApp.newTrigger('FTRTesting_Monday')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(8)
      .nearMinute(45)   // <--- Added this line
      .create();

  // --- FTRTesting (Tuesday 08:45) ---
  ScriptApp.newTrigger('FTRTesting_Tuesday')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .atHour(8)
      .nearMinute(45)   // <--- Added this line
      .create();

  // --- FTRTesting (Wednesday 08:45) ---
  ScriptApp.newTrigger('FTRTesting_Wednesday')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .atHour(8)
      .nearMinute(45)   // <--- Added this line
      .create();

  // --- FTRTesting (Friday 08:45) ---
  ScriptApp.newTrigger('FTRTesting_Friday')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .atHour(8)
      .nearMinute(45)   // <--- Added this line
      .create();
}

/**
 * Helper: Deletes ALL triggers.
 */
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ==========================================
// WRAPPER FUNCTIONS
// These exist solely to give your triggers a nice name in the list.
// ==========================================

function FTRTesting_Monday() {
  FTRTesting(); // Calls your main logic
}

function FTRTesting_Tuesday() {
  FTRTesting();
}

function FTRTesting_Wednesday() {
  FTRTesting();
}

function FTRTesting_Friday() {
  FTRTesting();
}
