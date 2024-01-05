/**
 * Runs when the spreadsheet is opened.
 */
function onOpen () {
  const ui = SpreadsheetApp.getUi()

  ui.createMenu('Sheets to Doc')
    .addItem('Write selected row to Doc', 'main')
    .addSeparator()
    .addItem('Info', 'info')
    .addToUi()
}

/**
 * Displays an alert with the creator's information.
 *
 * @returns {void}
 */
function info () {
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.
    .alert('Created by: Ruben Jonkers. For information contact me at sheetstodoc@rubenjonkers.nl')
}

/**
 * Main function that orchestrates the process of converting a Google Sheet to a Google Doc.
 * It performs various checks, retrieves data from the sheet, prompts the user for input,
 * creates necessary folders, writes the data to a document, and displays the document URL.
 * @returns {void}
 */
function main () {
  const checkResult = checkSelection()
  if (!checkResult) {
    // Stop execution if selection check fails
    return
  }

  const sheetEntriesResult = getSheetEntries();
  if (sheetEntriesResult.status === 'failure') {
    SpreadsheetApp.getUi().alert('Error: ' + sheetEntriesResult.message)
    return
  }

  const headersResult = getColumnHeaders();
  if (headersResult.status === 'failure') {
    SpreadsheetApp.getUi().alert('Error: ' + headersResult.message)
    return
  }

  const filenameResult = getFilenameFromUser();
  if (filenameResult.status === 'failure') {
    SpreadsheetApp.getUi().alert('Filename input error: ' + filenameResult.message)
    return
  }

  const foldersResult = createFolders();
  if (foldersResult.status === 'failure') {
    SpreadsheetApp.getUi().alert('Error creating folder: ' + foldersResult.message)
    return
  }

  const docResult = writeToDoc(filenameResult.filename, foldersResult, sheetEntriesResult.data, headersResult.data)
  if (docResult.status === 'failure') {
    SpreadsheetApp.getUi().alert('Error writing to document: ' + docResult.message)
    return
  }

  // Display the final message including the document URL
  showUrlDialog(docResult.url)
}

/**
 * Checks the current selection in the active spreadsheet.
 * The selection must be a single row.
 * If the selection is valid, returns true.
 * If the selection is invalid, shows an error message and returns false.
 * @returns {boolean} True if the selection is valid, false otherwise.
 */
function checkSelection () {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const range = sheet.getActiveRange()
  const numRows = range.getNumRows()
  const numColumns = range.getNumColumns()

  if (numRows === 1 && numColumns > 0) {
    // Valid selection
    return true
  } else {
    // Invalid selection, show an error message
    const errorMessage = numRows === 0 || numColumns === 0
      ? 'Please select a row.'
      : 'Please select only one row at a time.'
    SpreadsheetApp.getUi().alert(errorMessage)
    return false
  }
}

/**
 * Retrieves the entries from the active sheet.
 * @returns {Object} An object containing the status and data of the sheet entries.
 * - status: 'success' if the entries are retrieved successfully, 'failure' otherwise.
 * - data: An array of rows representing the sheet entries.
 * @throws {Error} If there is an error retrieving the sheet entries.
 */
function getSheetEntries () {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    const range = sheet.getActiveRange()
    const startRow = range.getRow()

    // Retrieve full rows for the selected range
    const sheetEntries = sheet.getRange(startRow, 1, 1, sheet.getLastColumn()).getValues()

    return { status: 'success', data: sheetEntries }
  } catch (error) {
    console.error('Error in getSheetEntries:', error)
    return { status: 'failure', message: error.message }
  }
}

/**
 * Retrieves the column headers from the active sheet in the spreadsheet.
 * @returns {Object} An object containing the status and data of the column headers.
 *                   - If successful, the status is 'success' and the data is an array of headers.
 *                   - If no headers are found, the status is 'failure' and the message is 'No headers found in the sheet.'
 *                   - If an error occurs, the status is 'failure' and the message is the error message.
 */
function getColumnHeaders () {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()

    if (headers && headers.length > 0) {
      // Check if headers are non-empty
      return { status: 'success', data: headers }
    } else {
      // No headers found
      return { status: 'failure', message: 'No headers found in the sheet.' }
    }
  } catch (error) {
    console.error('Error in getColumnHeaders:', error)
    return { status: 'failure', message: error.message }
  }
}

/**
 * Prompts the user to provide a filename and returns the result.
 * @returns {Object} An object containing the status and filename/message.
 */
function getFilenameFromUser () {
  const ui = SpreadsheetApp.getUi()
  const result = ui.prompt('Please provide a filename', 'filename', ui.ButtonSet.OK_CANCEL)

  // Process the user's response.
  const button = result.getSelectedButton()
  const text = result.getResponseText()

  if (button === ui.Button.OK && text) {
    // User clicked "OK" and entered a filename.
    ui.alert('The file will be saved as: ' + text)
    return { status: 'success', filename: text }
  } else {
    // User clicked "Cancel", "Close", or did not enter a filename.
    const message = button === ui.Button.CANCEL
      ? 'Action cancelled, no file will be saved.'
      : button === ui.Button.CLOSE
        ? 'You closed the dialog, no file will be saved.'
        : 'No filename provided.'
    ui.alert(message)
    return { status: 'failure', message }
  }
}

/**
 * Creates a folder named 'Sheet to Doc' in the parent directory of the active spreadsheet.
 * If the folder already exists, it retrieves the existing folder.
 * @returns {Object} An object containing the status and the folder.
 * - status: 'success' if the folder is created or retrieved successfully, 'failure' otherwise.
 * - folder: The created or retrieved folder.
 */
function createFolders () {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
    const id = sheet.getId()
    const folderId = DriveApp.getFileById(id).getParents().next().getId()
    const folder = DriveApp.getFolderById(folderId)

    let targetFolder

    // Check if 'Sheet to Doc' folder exists, create if it doesn't
    if (!folder.getFoldersByName('Sheet to Doc').hasNext()) {
      targetFolder = folder.createFolder('Sheet to Doc')
    } else {
      targetFolder = folder.getFoldersByName('Sheet to Doc').next()
    }

    return { status: 'success', folder: targetFolder }
  } catch (error) {
    console.error('Error in createFolders:', error)
    return { status: 'failure', message: error.message }
  }
}

/**
 * Creates a new document and writes the provided row data and headers to it.
 * @param {string} fileName - The name of the document to be created.
 * @param {object} foldersResult - The result of the folder selection.
 * @param {Array} row - The row data to be written to the document.
 * @param {Array} header - The headers to be written to the document.
 * @returns {object} - An object containing the status, message, and URL of the created document.
 */
function writeToDoc (fileName, foldersResult, row, header) {
  try {
    const doc = DocumentApp.create(fileName)
    DriveApp.getFileById(doc.getId()).moveTo(foldersResult.folder)
    const body = doc.getBody()

    const rowData = row[0] // Assuming single row selection
    const headers = header[0]

    // Write headers and row data to the document
    for (let i = 0; i < headers.length; i++) {
      // Ensure header exists and append as 'Heading 2'
      if (headers[i] !== undefined && headers[i] !== null) {
        const headerParagraph = body.appendParagraph(String(headers[i]))
        headerParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING2)
      } else {
        throw new Error('Header at index ' + i + ' is undefined or null.')
      }

      // Check if rowData[i] exists and append
      if (rowData[i] !== undefined && rowData[i] !== null) {
        let answerText = rowData[i]
        // Convert non-string data to string
        if (typeof answerText !== 'string') {
          answerText = answerText instanceof Date
            ? answerText.toISOString() // Convert Date to ISO string
            : JSON.stringify(answerText) // Convert other non-string types to JSON string
        }
        body.appendParagraph(answerText)
      } else {
        throw new Error('Row data at index ' + i + ' is undefined or null.')
      }
    }
    // Save and close the document
    doc.saveAndClose()

    // Get the URL of the created document
    const docUrl = doc.getUrl()

    return { status: 'success', message: 'Document created successfully. URL: ' + docUrl, url: docUrl }
  } catch (error) {
    console.error('Error in writeToDoc:', error)
    return { status: 'failure', message: error.message }
  }
}

/**
 * Displays a modal dialog with a URL link to open a document.
 *
 * @param {string} url - The URL of the document to be opened.
 * @returns {void}
 */
function showUrlDialog (url) {
  const htmlOutput = HtmlService
    .createHtmlOutput('<p><b>Open the file by clicking the link below</b></p> <p><a href="' + url + '" target="_blank">Open Document</a></p>')
    .setWidth(500)
    .setHeight(200)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'The file has successfully been saved')
}
