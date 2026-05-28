/**
 * @file 17_Config.gs
 * @description Global configuration and environment settings.
 * Merged with centralized folder IDs and today's recovery logic.
 */

// --- NOTIFICATION SETTINGS ---
/** @type {boolean} Set to false to silence all Google Chat notifications */
const ENABLE_CHAT_NOTIFICATIONS = false; //Sandbox

// --- FOLDER IDs ---
const DAVE_FOLDER_ID = '1KQyIZQoQRDGXrUHPleIFCcG5u_ykLIIn'; //Remains unchanged in Sandbox
// const DAVE_FOLDER_ID_OVERRIDE = '1GnOgMWzTa1qHP_JK5YJ41TF3fNs0tpws'; // Manual override placeholder

// Centralized Report & Backup Location
const CENTRAL_DESTINATION_ID = '1JuoKItuYGMXu325zJUEzSBxz00Ub3rQs'; //Sandbox

const FTRIMPORT_FOLDER_ID_PROD = CENTRAL_DESTINATION_ID;
const DAVEREPORTS_FOLDER_ID_PROD = CENTRAL_DESTINATION_ID;
const SOURCE_FOLDER_ID_PROD = '1tiBI_2h9nHjZwxAUp1vweuLudCgSZ1hM'; //Sandbox
const TARGET_FOLDER_ID_PROD = CENTRAL_DESTINATION_ID;

const FTRIMPORT_FOLDER_ID_TEST = CENTRAL_DESTINATION_ID;
const DAVEREPORTS_FOLDER_ID_TEST = CENTRAL_DESTINATION_ID;
const SOURCE_FOLDER_ID_TEST = '10KxCCs-yVfseMVQC-gwTatYU_kZPYWOT'; //Sandbox
const TARGET_FOLDER_ID_TEST = CENTRAL_DESTINATION_ID;
const BACKUP_FOLDER_ID_TEST = CENTRAL_DESTINATION_ID;

// --- FILE IDs and NAMES ---
const FTR_SHEET_ID_PROD = '16ofaDyDX8JWFvJhGdxNTHhUA8nlZ0R1SMTEYZli0vg4'; //Sandbox
const FTR_SLIDES_ID_PROD = '1deqrvAECl42t4Cp5PYC3QHvewDfMG4aWG0Tq9Ubg1-g'; //Sandbox
const FTR_SHEET_NAME_PROD = 'FTRMK4'; //Sandbox

const FTR_SHEET_NAME_TEST = 'Copy of Fix the Rack';
const FTR_SLIDES_NAME_TEST = 'Copy of Fix the Rack Slides';

// --- WEBHOOKS and COLORS ---
const GOOGLE_CHAT_WEBHOOK_LINK = "https://chat.googleapis.com/v1/spaces/AAAA6bERKyU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Izk2fFQjvA8Zy-PufTiw7JWoquIjLmESRr_yf1pWs9w";
const GREEN_BACKGROUND = "#b6d7a8";
const RED_BACKGROUND = "#ea9999";

// --- SHEET-SPECIFIC SETTINGS ---
const LOOKUP_SHEET_NAME = "LookupTables";
const FULL_CW_CELL = "K33";
const ROW_CELL = "H31";
const CW_CELL = "K22";
const ACTUAL_CW_CELL = "J22";

/**
 * Creates a custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('FTR Automation')
    .addItem('► Run Full PRODUCTION Update', 'runProdUpdateManually')
    .addItem('► Run Full TESTING Update', 'runTestUpdateManually')
    .addSeparator()
    .addItem('Initialize New Test Environment', 'initializeTestingAndRelink')
    .addToUi();
}

/**
 * Manual trigger for Production with Recovery Logic.
 */
function runProdUpdateManually() {
  const isResuming = handleWorkflowRecovery();

  if (isResuming) {
    // If resuming, we call the engine directly to pick up the existing state
    processChainedWorkflow();
  } else {
    // If starting clean, we run the normal router logic
    runDailyUpdate(true);
  }
}

/**
 * Manual trigger for Testing with Recovery Logic.
 */
function runTestUpdateManually() {
  const isResuming = handleWorkflowRecovery();

  if (isResuming) {
    // If resuming, we call the engine directly to pick up the existing state
    processChainedWorkflow();
  } else {
    // If starting clean, we run the normal router logic
    runDailyUpdate(false, true);
  }
}
