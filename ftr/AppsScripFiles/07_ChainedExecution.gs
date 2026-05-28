/**
 * @file 07_ChainedExecution.gs
 * @description The "Relay Race" engine. Manages state and self-triggering to bypass 6-minute limits.
 * Updated to split Data Archive from Visuals with a mandatory settling pause to prevent timeouts.
 * Fixed: Step Index incrementing during forced pauses to prevent loops.
 */

const WORKFLOW_STEPS = [
  'RELINK_CHARTS',      // 0
  'CLONE_SHEETS',       // 1
  'IMPORT_AND_PREPARE', // 2
  'UPDATE_CABLING_DATA',// 3
  'UPDATE_WAVE_DATA',   // 4
  'INTEGRITY_CHECK',    // 5
  'CALCULATE_DELTAS',   // 6
  'DATA_ARCHIVE_STEP',  // 7: Updates week cell and moves table data
  'CHART_UPDATE_STEP',  // 8: Updates titles/axes (Heavy API step)
  'SLIDE_UPDATE_STEP',  // 9: Refreshes charts in slides
  'FINISH_AND_NOTIFY'   // 10: Sends Chat card
];

/**
 * Main entry point for the self-triggering chain.
 */
function processChainedWorkflow() {
  const props = PropertiesService.getScriptProperties();
  const startTime = new Date().getTime();

  let currentStepIndex = parseInt(props.getProperty('STEP_INDEX') || '0');
  let configString = props.getProperty('WF_CONFIG');

  if (!configString) {
    Logger.log("⛔ Error: WF_CONFIG not found. Workflow cannot resume.");
    return;
  }

  let config = JSON.parse(configString);
  const ss = SpreadsheetApp.openById(config.sheetId);

  try {
    while (currentStepIndex < WORKFLOW_STEPS.length) {
      let now = new Date().getTime();
      let elapsedSeconds = (now - startTime) / 1000;
      const stepName = WORKFLOW_STEPS[currentStepIndex];

      // SAFETY CHECK: Baton pass if we are approaching the 5-minute mark
      if (elapsedSeconds > 240) {
        createResumptionTrigger_(ss);
        return;
      }

      workflowLog(`>>> Starting Step ${currentStepIndex}: ${stepName}`, "CHAIN_ENGINE", startTime, ss);

      try {
        // Execute the specific logic for this step
        executeStep_(stepName, config, startTime, ss, currentStepIndex);

        // Force commit to trigger recalculations
        SpreadsheetApp.flush();
      } catch (e) {
        // Catch intentional relay signals
        if (e.message === "PAUSE_REQUESTED") {
          let catchElapsed = (new Date().getTime() - startTime) / 1000;
          workflowLog(`Relay requested for recalculation at ${catchElapsed.toFixed(1)}s.`, "RELAY_PAUSE", startTime, ss);
          createResumptionTrigger_(ss);
          return;
        }
        throw e;
      }

      // Successfully completed the step, move to next
      currentStepIndex++;
      props.setProperty('STEP_INDEX', currentStepIndex.toString());

      // Mandatory 2-second settle time between standard steps
      Utilities.sleep(2000);
    }

    cleanupWorkflow_(ss);
  } catch (e) {
    const errorMsg = `⛔ CRITICAL ERROR in ${WORKFLOW_STEPS[currentStepIndex] || 'Engine'}: ${e.message}`;
    workflowLog(errorMsg, "CHAIN_ERROR", startTime, ss);
    sendErrorNotificationToChat(errorMsg, WORKFLOW_STEPS[currentStepIndex], config);
    stopAllResumptionTriggers_();
  }
}

/**
 * Executes individual steps and maps them to your logic files.
 * @private
 */
function executeStep_(stepName, config, startTime, ss, currentStepIndex) {
  const props = PropertiesService.getScriptProperties();

  switch (stepName) {
    case 'RELINK_CHARTS':
      if (config.context === "Testing") {
        relinkChartsChained(config.slidesId, config.sheetId, FTR_SHEET_ID_PROD, startTime);
      }
      break;

    case 'CLONE_SHEETS':
      if (config.isWeeklyRun) cloneDataSheets(ss);
      break;

    case 'IMPORT_AND_PREPARE':
      processAndLoadCSVs(config, startTime);
      break;

    case 'UPDATE_CABLING_DATA':
      const count = parseInt(props.getProperty('UNIQUE_RACK_COUNT') || '0');
      modifyCablingData(ss, count + 1);
      break;

    case 'UPDATE_WAVE_DATA':
      modifyWaveAssignment(ss);
      break;

    case 'INTEGRITY_CHECK':
      performIntegrityCheck(ss, config);
      break;

    case 'CALCULATE_DELTAS':
      if (config.isWeeklyRun) { rackDelta(ss); pduDelta(ss); cablingDelta(ss); }
      break;

    case 'DATA_ARCHIVE_STEP':
      // 1. Update week cell (Phase 1)
      updateWeekAndCheck_(ss);
      // 2. Perform table roll (Phase 2)
      rollWeeksAndPullCurrent(ss);

      // Commit changes immediately so the spreadsheet is ready for math settling
      SpreadsheetApp.flush();

      // FIXED: Increment the index property BEFORE throwing the pause error.
      // This ensures that when the script resumes, it starts at Step 8 (CHART_UPDATE_STEP).
      props.setProperty('STEP_INDEX', (currentStepIndex + 1).toString());

      workflowLog("Data Archive complete. Forcing relay pause for spreadsheet math settling.", "WEEKLY_ARCHIVE", startTime, ss);
      throw new Error("PAUSE_REQUESTED");

    case 'CHART_UPDATE_STEP':
      updateCharts(ss);
      break;

    case 'SLIDE_UPDATE_STEP':
      updateSlides(config.slidesId);
      break;

    case 'FINISH_AND_NOTIFY':
      sendCardMessageToChat(config.sheetId, config.slidesId, config.context);
      break;
  }
}

/**
 * Schedules a resumption of the workflow.
 */
function createResumptionTrigger_(ss) {
  stopAllResumptionTriggers_();

  ScriptApp.newTrigger('processChainedWorkflow')
    .timeBased()
    .after(60 * 1000)
    .create();

  workflowLog("Resumption trigger created (1 minute delay).", "RELAY_PAUSE", null, ss);
}

/**
 * Helper to clean up existing triggers to prevent duplicates.
 */
function stopAllResumptionTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processChainedWorkflow') ScriptApp.deleteTrigger(t);
  });
}

/**
 * Finalizes the workflow and wipes all progress properties.
 */
function cleanupWorkflow_(ss) {
  const props = PropertiesService.getScriptProperties();
  stopAllResumptionTriggers_();

  const keysToWipe = [
    'STEP_INDEX',
    'WF_CONFIG',
    'RELINK_CHART_INDEX',
    'IMPORT_STAGE',
    'PDU_IMPORT_ROW_INDEX',
    'UNIQUE_RACK_COUNT'
  ];

  keysToWipe.forEach(p => props.deleteProperty(p));
  workflowLog("✅ Workflow Finished Successfully.", "FINISHED", null, ss);
}
