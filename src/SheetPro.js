/**
* MIT License
*
* Copyright (c) 2021 Ashton Fei
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in all
* copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
* SOFTWARE.
*/

/**
 * Get selected ranges
 * 
 * @returns {SpreadsheetApp.Range[]} An array of selected ranges
 */
function getSelectedRanges() {
    return SpreadsheetApp.getActive().getSelection().getActiveRangeList().getRanges()
}

/**
 * Get selected rows
 * 
 * @returns {SpreadsheetApp.Range[]} An array of selected rows
 */
function getSelectedRows() {
    const rowNotationRegex = /[\d]+:[\d]+/
    const ranges = getSelectedRanges()
    return ranges.filter(range => rowNotationRegex.test(range.getA1Notation()))
}

/**
 * Get selected columns
 * 
 * @returns {SpreadsheetApp.Range[]} An array of selected columns
 */
function getSelectedColumns() {
    const rowNotationRegex = /[A-Z]+:[A-Z]+/
    const ranges = getSelectedRanges()
    return ranges.filter(range => rowNotationRegex.test(range.getA1Notation()))
}


/**
 * Create a new UI object
 * 
 * @exampe <caption>Create a UI objet for app SheetPro</caption>
 * const ui = new UI("SheetPro")
 * 
 * @param {string} [appName = SheetPro] The name of your application
 * @returns {UI} The UI object
 */
function UI(appName = "SheetPro") {
    this.appName = appName
    this.ui = SpreadsheetApp.getUi()
    this.ALERT_TYPE = {
      ALERT: "Alert",
      WARNING: "Warning",
      MESSAGE: "Message",
      INFO: "Info",
      ERROR: "Error",
      SUCCESS: "Success",
    }
}

/**
 * Show an alert message
 * 
 * @param {string} message The alert message 
 * @param {title} [title=this.appName] The alert title
 * @return {void} 
 */
UI.prototype.alert = function (message, title) {
    title = title || `${this.appName}`
    this.ui.alert(title, message, this.ui.ButtonSet.OK)
}

/**
 * Show an alert message with a message type in the alert title
 * 
 * @param {string} message The alert message
 * @param {UI.ALERT_TYPE} type The type of alert 
 */
UI.prototype.alertWithType = function(message, type=this.ALERT_TYPE.ALERT){
  this.alert(message, `${this.appName} [${type}]`)
}