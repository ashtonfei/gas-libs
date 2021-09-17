class App {
  constructor({ id, sheetName, sourceHeaderName }) {
    if (id) {
      try {
        this.ss = SpreadsheetApp.openById(id)
      } catch (err) {
        throw new Error(err.message)
      }
    } else {
      this.ss = SpreadsheetApp.getActive()
    }
    this.sheetName = sheetName
    this.sourceHeaderName = sourceHeaderName
    this.ws = this.ss.getSheetByName(this.sheetName) || this.ss.insertSheet(this.sheetName)
  }

  static createTrigger(functionName) {
    ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger))
    return ScriptApp.newTrigger(functionName)
      .forSpreadsheet(SpreadsheetApp.getActive().getId())
      .onFormSubmit()
      .create()
  }

  handleNamedValues(source, namedValues, currentHeaders) {
    const headers = currentHeaders.map(v => v.toString().trim()).filter(v => v !== "")
    if (!headers.includes(this.sourceHeaderName)) {
      headers.unshift(this.sourceHeaderName)
    }
    const values = headers.map(v => {
      if (v === this.sourceHeaderName) {
        return source
      } else {
        return null
      }
    })
    const keys = Object.keys(namedValues)
    keys.sort()
    keys.forEach((key) => {
      if (key) {
        const value = namedValues[key].length === 1 ? namedValues[key][0] : namedValues[key].join(", ")
        if (headers.includes(key)) {
          values[headers.indexOf(key)] = value
        } else {
          headers.push(key)
          values.push(value)
        }
      }
    })
    console.log({headers, values, source})
    return { headers, values }
  }

  concatenate(e) {
    const { range, namedValues } = e
    const source = range.getSheet().getName()
    const currentHeaders = this.ws.getDataRange().getValues()[0]
    const { headers, values } = this.handleNamedValues(source, namedValues, currentHeaders)
    this.ws.getRange(1, 1, 1, headers.length).setValues([headers])
    this.ws.appendRow(values)
  }
}

/**
 * Concatenate form response to a target sheet
 * @param {ScriptApp.EventType.ON_FORM_SUBMIT} event An ON_FORM_SUBMIT event
 * @param {string} sheetName The name of the target sheet
 * @param {string} sourceHeaderName The header name of the form source column
 * @param {string=} spreadsheetId The id of the target spreadsheet, default value is null
 * @return {object} The form resonpnse object
 */
function concatenate(event, sheetName, sourceHeaderName, spreadsheetId = null) {
  const app = new App({ id: spreadsheetId, sheetName, sourceHeaderName })
  return app.concatenate(event)
}

/**
 * Create a onFormSubmit trigger with a function name
 * @param {string=} functionName The name of function to be triggered, default value is "onFormSubmit_"
 * @return {ScriptApp.Trigger} ScriptApp Trigger object
 */
function createTrigger(functionName = "onFormSubmit_") {
  return App.createTrigger(functionName)
}

if (typeof module === "object") {
  module.exports = {
    concatenate,
    createTrigger,
  }
}