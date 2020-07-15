// ABOUT
// Script for processing acceleration data for 1D project
// Created by Elvis on 20/06/20

let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let ui = SpreadsheetApp.getUi();


/*
 * TRIGGER onOpen
 * Executes code upon opening the file.
 *
 * Adds new menu items
 */

function onOpen() {
  ui.createMenu('1D Menu')
    .addItem('Set Variables Acc', 'setVariablesAcc')
    .addItem('Process Acc Data', 'processAccData')
    .addToUi();
}

/*
 * Function: getVarByPrompt
 * Returns user's input from a generated prompt
 *
 * @param {String} instruction
 * @param {String} text
 *
 * @return {String} result
 */

function getVarByPrompt(instruction, text) {
  let prompt = ui.prompt(instruction, text, ui.ButtonSet.OK_CANCEL);
  let button = prompt.getSelectedButton();
  let result = prompt.getResponseText();
  if (button == ui.Button.OK) {
    return result;
  }
}

/*
 * Function: createSheet
 * Inserts a new spreadsheet with given name at the given index position
 *
 * @param {String} name
 * @param {Int} index
 */

function createSheet(name, index) {
  activeSpreadsheet.insertSheet(name, index);
}

/* 
 * Function: cellA1ToIndex
 * Convert A1 to Index
 * Yanked from https://codereview.stackexchange.com/questions/90112/a1notation-conversion-to-row-column-index
 *
 * @param {String} cellA1
 *
 * @result {Object} row, col
 */

function cellA1ToIndex(cellA1) {
  let match = cellA1.match(/^\$?([A-Z]+)\$?(\d+)$/);

  if (!match) {
    throw new Error( "Invalid cell reference" );
  }

  return {
    row: rowA1ToIndex(match[2]),
    col: colA1ToIndex(match[1])
  };
}

/* 
 * Function: colA1ToIndex
 * Returns column value from A1 notation
 * Yanked from https://codereview.stackexchange.com/questions/90112/a1notation-conversion-to-row-column-index
 *
 * @param {String} cellA1
 *
 * @result {Int} sum
 */

function colA1ToIndex(colA1) {
  let i, l, chr,
      sum = 0,
      A = "A".charCodeAt(0),
      radix = "Z".charCodeAt(0) - A + 1;

  if (typeof colA1 !== 'string' || !/^[A-Z]+$/.test(colA1)) {
    throw new Error("Expected column label");
  }

  for (i = 0, l = colA1.length ; i < l ; i++) {
    chr = colA1.charCodeAt(i);
    sum = sum * radix + chr - A + 1;
  }

  return sum;
}

/* 
 * Function: rowA1ToIndex
 * Returns row value from A1 notation
 * Yanked from https://codereview.stackexchange.com/questions/90112/a1notation-conversion-to-row-column-index
 *
 * @param {String} cellA1
 *
 * @result {Int} index
 */

function rowA1ToIndex(rowA1) {
  let index = parseInt(rowA1, 10);
  if (isNaN(index)) {
    throw new Error("Expected row number");
  }
  return index;
}

/* 
 * Function: columnToLetter
 * Converts Column Index to Corresponding Letter
 * 
 * @param {Integer} column index
 *
 * @return {String} column letter
 */

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/*
 * Function: copyColumn
 * Copies columns given the source and destination
 *
 * @param {Object} srcSheet
 * @param {Object} dstSheet
 * @param {Integer} srcColumnIndex
 * @param {Integer} dstColumnIndex
 */

function copyColumn(srcSheet, dstSheet, srcColumnIndex, dstColumnIndex) {
  let range = srcSheet.getRange(1, srcColumnIndex, srcSheet.getLastRow());
  let dataToCopy = range.getValues();
  
  // Using the getRange(rowIndex, colIndex, numRows) which returns the range starting from coords set by rowIndex and colIndex
  dstSheet.getRange(1, dstColumnIndex, srcSheet.getLastRow()).setValues(dataToCopy);
}


/* 
 * Function: setVariablesAcc
 * Sets acceleration variables by storing them in document properties
 */
function setVariablesAcc() {
  PropertiesService.getDocumentProperties().setProperty("dataSheet", getVarByPrompt("Enter name of sheet with logged data", " ").toString());
  PropertiesService.getDocumentProperties().setProperty("unixTime", getVarByPrompt("Enter column id to extract", "Log timestamp in UNIX time").toString());
  PropertiesService.getDocumentProperties().setProperty("logSample", getVarByPrompt("Enter column id to extract", "Log sample count").toString());
  PropertiesService.getDocumentProperties().setProperty("accX", getVarByPrompt("Enter column id to extract", "X-axis acceleration").toString());
  PropertiesService.getDocumentProperties().setProperty("accY", getVarByPrompt("Enter column id to extract", "Y-axis acceleration").toString());
  PropertiesService.getDocumentProperties().setProperty("accZ", getVarByPrompt("Enter column id to extract", "Z-axis acceleration").toString());
  PropertiesService.getDocumentProperties().setProperty("shiftSampleFrom", getVarByPrompt("Enter row number", "For the start cell in order to sample for the baseline shift").toString());
  PropertiesService.getDocumentProperties().setProperty("shiftSampleTo", getVarByPrompt("Enter row number", "For the end cell in order to sample for the baseline shift").toString());
  
  ui.alert("Variables set");
}


/* 
 * Function: calculateAccColumns
 * Calculates values required for acceleration insight generation. Does things like numerical integration etc.
 *
 * @param {Object} dstSheet
 * @param {Int} srcSheetRowCount
 * @param {Object} columnNumbers
 * @param {Object} columnLetters
 */
function calculateAccColumns(dstSheet, srcSheetRowCount, columnNumbers, columnLetters) {
  let colLetterUnixTime = columnNumbers.colLetterUnixTime;
  let colLetterLogSample = columnNumbers.colLetterLogSample;
  let colLetterAccY = columnNumbers.colLetterAccY;
  
  let colNumElapsedTime = columnNumbers.colNumElapsedTime;
  let colLetterElapsedTime = columnLetters.colLetterElapsedTime;
  
  let colNumGI = columnNumbers.colNumGravityCompensatedInverted;
  let colLetterGI = columnLetters.colLetterGravityCompensatedInverted;
  
  let colNumShiftValue = columnNumbers.colNumShiftValue;
  let colLetterShiftValue = columnLetters.colLetterShiftValue;
  
  let colNumShift = columnNumbers.colNumShift;
  let colLetterShift = columnLetters.colLetterShift;
  
  let colNumDeltaT = columnNumbers.colNumDeltaT;
  let colLetterDeltaT = columnLetters.colLetterDeltaT;
  
  let colNumIV = columnNumbers.colNumIntegralVelocity;
  let colLetterIV = columnLetters.colLetterIntegralVelocity;
  
  let colNumVT = columnNumbers.colNumVelocityAtTime;
  let colLetterVT = columnLetters.colLetterVelocityAtTime;
  
  let colNumID = columnNumbers.colNumIntegralDisplacement;
  let colLetterID = columnLetters.colLetterIntegralDisplacement;
  
  let colNumDT = columnNumbers.colNumDisplacementAtTime;
  let colLetterDT = columnLetters.colLetterDisplacementAtTime;
  
  
  
  // COLUMN Calculate Elapsed Time
  dstSheet.getRange(1, colNumElapsedTime).setValue("loggingElapsedTime");
  dstSheet.getRange(2, colNumElapsedTime).setValue(0);
  
  for (let i = 3; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumElapsedTime).setFormula("=" + colLetterUnixTime + i + "-" + colLetterUnixTime + (i-1) + "+" + colLetterElapsedTime + (i-1));
  }

  // COLUMN Compensate for g and invert it
  dstSheet.getRange(1, colNumGI).setValue("accY-compensated-inverted");
  
  for (let i = 2; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumGI).setFormula("-(" + colLetterAccY + i + "*9.81)");
  }
  
  // COLUMN Calculate Shift Value
  let shiftSampleFrom = PropertiesService.getDocumentProperties().getProperty("shiftSampleFrom");
  let shiftSampleTo = PropertiesService.getDocumentProperties().getProperty("shiftSampleTo");
  dstSheet.getRange(1, colNumShiftValue).setValue("shiftValue");
  dstSheet.getRange(2, colNumShiftValue).setFormula("SUM(" + colLetterGI + shiftSampleFrom + ":" + colLetterGI + shiftSampleTo + ")/COUNT(" + colLetterGI + shiftSampleFrom + ":" + colLetterGI + shiftSampleTo + ")");
  
  // COLUMN Compensate for Shift
  dstSheet.getRange(1, colNumShift).setValue("accY-compensated-inverted-shifted");
  
  for (let i = 2; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumShift).setFormula(colLetterGI + i + "-" + "$" + colLetterShiftValue + "$" + 2);
  }
  
  // COLUMN DeltaT
  dstSheet.getRange(1, colNumDeltaT).setValue("deltaT");
  dstSheet.getRange(2, colNumDeltaT).setValue(0);
  
  for (let i = 3; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumDeltaT).setFormula(colLetterElapsedTime + i + "-" + colLetterElapsedTime + (i-1));
  }
  
  // COLUMN int_v
  dstSheet.getRange(1, colNumIV).setValue("int_v");
  
  for (let i = 3; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumIV).setFormula("0.5*(" + colLetterShift + i + "+" + colLetterShift + (i-1) + ")*" + colLetterDeltaT + i);
  }
  
  // COLUMN v(t)
  dstSheet.getRange(1, colNumVT).setValue("v(t)");
  dstSheet.getRange(2, colNumVT).setValue(0);
  
  for (let i = 3; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumVT).setFormula("0.5*(" + colLetterShift + i + "+" + colLetterShift + (i-1) + ")*" + colLetterDeltaT + i + "+" + colLetterVT + (i-1));
  }
  
  // COLUMN int_d
  dstSheet.getRange(1, colNumID).setValue("int_d");
  
  for (let i = 3; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumID).setFormula("0.5*(" + colLetterVT + i + "+" + colLetterVT + (i-1) + ")*" + colLetterDeltaT + i);
  }
  
  // COLUMN d(t)
  dstSheet.getRange(1, colNumDT).setValue("d(t)");
  dstSheet.getRange(2, colNumDT).setValue(0);
  
  for (let i = 3; i <= srcSheetRowCount; i++) {
    dstSheet.getRange(i, colNumDT).setFormula("0.5*(" + colLetterVT + i + "+" + colLetterVT + (i-1) + ")*" + colLetterDeltaT + i + "+" + colLetterDT + (i-1));
  }
}


/*
 * Function: createChart
 * Creates a chart
 * 
 * @param {type} variable
 */

function createChart(options, dstSheet) {
  let chart = dstSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dstSheet.getRange(options.rangeX))
    .addRange(dstSheet.getRange(options.rangeY))
    .setPosition(options.position.anchorRowPos, options.position.anchorColPos, options.position.offsetX, options.position.offsetY)
    .setOption('title', options.title)
    .setOption('legend', {position: 'bottom', textStyle: {color: 'black', fontSize: 12}})
    .setOption('useFirstColumnAsDomain', true)
    .setNumHeaders(1)
    .build();
  
  dstSheet.insertChart(chart);
}


/* 
 * Function: createAccInsightGraphs 
 * Creates A/t, V/t, and S/t graphs as part of processAccData
 *
 * @param {Object} dstSheet
 * @param {Int} srcSheetRowCount
 * @param {Object} columnNumbers
 * @param {Object} columnLetters
 */

function createAccInsightGraphs(dstSheet, srcSheetRowCount, columnNumbers, columnLetters) {
  
  // Chart Options
  let options = {
    title: "",
    position: {
      anchorRowPos: 4,
      anchorColPos: 20,
      offsetX: 0,
      offsetY: 0,
    },
    rangeX: "",
    rangeY: "",
  }
  
  // Create d/t graph
  options.rangeX = columnLetters.colLetterElapsedTime + 1 + ":" + columnLetters.colLetterElapsedTime + srcSheetRowCount;
  options.rangeY = columnLetters.colLetterDisplacementAtTime + 1 + ":" + columnLetters.colLetterDisplacementAtTime + srcSheetRowCount;
  options.title = "d/t";
  createChart(options, dstSheet);
  
  // Create v/t graph
  options.rangeX = columnLetters.colLetterElapsedTime + 1 + ":" + columnLetters.colLetterElapsedTime + srcSheetRowCount;
  options.rangeY = columnLetters.colLetterVelocityAtTime + 1 + ":" + columnLetters.colLetterVelocityAtTime + srcSheetRowCount;
  options.title = "v/t";
  options.position.anchorRowPos += 20;
  createChart(options, dstSheet);
  
  // Create a/t graph
  options.rangeX = columnLetters.colLetterElapsedTime + 1 + ":" + columnLetters.colLetterElapsedTime + srcSheetRowCount;
  options.rangeY = columnLetters.colLetterShift + 1 + ":" + columnLetters.colLetterShift + srcSheetRowCount;
  options.title = "a/t";
  options.position.anchorRowPos += 20;
  createChart(options, dstSheet);
}

/* 
 * Function: calculateAccInsights 
 * Calculates insights such as max velocity attained, the displacement etc.
 *
 * @param {Object} dstSheet
 * @param {Int} srcSheetRowCount
 * @param {Object} columnNumbers
 * @param {Object} columnLetters
 */

function calculateAccInsights(srcSheet, srcSheetRowCount, columnNumbers, columnLetters) {

  let colLDT = columnLetters.colLetterDisplacementAtTime;
  let colLVT = columnLetters.colLetterVelocityAtTime;
  let colLShift = columnLetters.colLetterShift;
  
  // Total Distance Travelled
  srcSheet.getRange(2, columnNumbers.colNumInsightAnchor).setValue("Total Distance Travelled");
  srcSheet.getRange(2, columnNumbers.colNumInsightAnchor+1).setFormula("MAX(" + colLDT + 2 + ":" + colLDT + srcSheetRowCount + ")");
  // Max Velocity Attained
  srcSheet.getRange(3, columnNumbers.colNumInsightAnchor).setValue("Max Velocity Attained");
  srcSheet.getRange(3, columnNumbers.colNumInsightAnchor+1).setFormula("MAX(" + colLVT + 2 + ":" + colLVT + srcSheetRowCount + ")");
  // Max Acceleration Attained
  srcSheet.getRange(4, columnNumbers.colNumInsightAnchor).setValue("Max Acceleration Attained");
  srcSheet.getRange(4, columnNumbers.colNumInsightAnchor+1).setFormula("MAX(" + colLShift + 2 + ":" + colLShift + srcSheetRowCount + ")");
  // Mean Velocity
  srcSheet.getRange(5, columnNumbers.colNumInsightAnchor).setValue("Mean Velocity");
  srcSheet.getRange(5, columnNumbers.colNumInsightAnchor+1).setFormula("AVERAGE(" + colLVT + 2 + ":" + colLVT + srcSheetRowCount + ")");
  // Mean Acceleration
  srcSheet.getRange(6, columnNumbers.colNumInsightAnchor).setValue("Mean Acceleration");
  srcSheet.getRange(6, columnNumbers.colNumInsightAnchor+1).setFormula("AVERAGE(" + colLShift + 2 + ":" + colLShift + srcSheetRowCount + ")");
}



/* 
 * Function: processAccData 
 * Function made up of sub-functions that are used to calculate and produce insights and charts based on acceleration data
 */
function processAccData() {
  
  if (PropertiesService.getDocumentProperties().getProperty("unixTime") === null) {
    throw new Error("Variables not set! Go to script-1 > Set Acc Variables");
    return;
  }
  
  // INITIAL STEP
  // Create new sheet
  let srcSheetName = PropertiesService.getDocumentProperties().getProperty("dataSheet");
  let dstSheetName = "acc-processed";
  let sheetCount = activeSpreadsheet.getSheets().length;
 
  createSheet(dstSheetName, sheetCount);
  
  // VARIABLES
  // --- Sheet
  let srcSheet = activeSpreadsheet.getSheetByName(srcSheetName);
  let dstSheet = activeSpreadsheet.getSheetByName(dstSheetName);
  let srcSheetRowCount = srcSheet.getLastRow();

  // --- Columns
  let colLetterUnixTime = columnToLetter(1);
  let colLetterLogSample = columnToLetter(2);
  let colLetterAccX = columnToLetter(3);
  let colLetterAccY = columnToLetter(4);
  let colLetterAccZ = columnToLetter(5);

  let columnNumbers = {
    colLetterUnixTime: colLetterUnixTime, 
    colLetterLogSample: colLetterLogSample,
    colLetterAccX: colLetterAccX,
    colLetterAccY: colLetterAccY,
    colLetterAccZ: colLetterAccZ,
    colNumElapsedTime: 6,
    colNumShiftValue: 7,
    colNumGravityCompensatedInverted: 8,
    colNumShift: 9,
    colNumDeltaT: 11,
    colNumIntegralVelocity: 12,
    colNumVelocityAtTime: 13,
    colNumIntegralDisplacement: 14,
    colNumDisplacementAtTime: 15,
    colNumInsightAnchor: 17
  }
  
  // --- Column Letters
  let colLetterElapsedTime = columnToLetter(columnNumbers.colNumElapsedTime);
  let colLetterShiftValue = columnToLetter(columnNumbers.colNumShiftValue);
  let colLetterGravityCompensatedInverted = columnToLetter(columnNumbers.colNumGravityCompensatedInverted);
  let colLetterShift = columnToLetter(columnNumbers.colNumShift);
  let colLetterDeltaT = columnToLetter(columnNumbers.colNumDeltaT);
  let colLetterIntegralVelocity = columnToLetter(columnNumbers.colNumIntegralVelocity);
  let colLetterVelocityAtTime = columnToLetter(columnNumbers.colNumVelocityAtTime);
  let colLetterIntegralDisplacement = columnToLetter(columnNumbers.colNumIntegralDisplacement);
  let colLetterDisplacementAtTime = columnToLetter(columnNumbers.colNumDisplacementAtTime);
  

  let columnLetters = {
    colLetterElapsedTime: colLetterElapsedTime,
    colLetterShiftValue: colLetterShiftValue,
    colLetterGravityCompensatedInverted: colLetterGravityCompensatedInverted,
    colLetterShift: colLetterShift,
    colLetterDeltaT: colLetterDeltaT,
    colLetterIntegralVelocity: colLetterIntegralVelocity,
    colLetterVelocityAtTime: colLetterVelocityAtTime,
    colLetterIntegralDisplacement: colLetterIntegralDisplacement,
    colLetterDisplacementAtTime: colLetterDisplacementAtTime,
  }

  // STEPS
  
  // Copy required data in columns to new sheet
  copyColumn(srcSheet, dstSheet, cellA1ToIndex(PropertiesService.getDocumentProperties().getProperty("unixTime") + "1").col, 1);
  copyColumn(srcSheet, dstSheet, cellA1ToIndex(PropertiesService.getDocumentProperties().getProperty("logSample") + "1").col, 2);
  copyColumn(srcSheet, dstSheet, cellA1ToIndex(PropertiesService.getDocumentProperties().getProperty("accX") + "1").col, 3);
  copyColumn(srcSheet, dstSheet, cellA1ToIndex(PropertiesService.getDocumentProperties().getProperty("accY") + "1").col, 4);
  copyColumn(srcSheet, dstSheet, cellA1ToIndex(PropertiesService.getDocumentProperties().getProperty("accZ") + "1").col, 5);
  
  // Calculate results required for analysis and put them into separate columns
  calculateAccColumns(dstSheet, srcSheetRowCount, columnNumbers, columnLetters);
  
  // Calculate insights such as max velocity achieved from the previous
  calculateAccInsights(dstSheet, srcSheetRowCount, columnNumbers, columnLetters);
  
  // Create Graphs
  createAccInsightGraphs(dstSheet, srcSheetRowCount, columnNumbers, columnLetters)

}
