(function (host, expose) {
   var module = { exports: {} };
   var exports = module.exports;
   /****** code begin *********/
/**
* MIT License
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
 * Get the first paragraph in the document with a keyword
 * @example <caption>Get paragrahp with keyword 'Google'</caption>
 * getParagraphByKeyword(doc, "Google")
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the paragraph
 * @returns {(DocumentApp.Paragraph | void)} The DocumentApp.Paragraph object or undefined when keyword not found
 */
function getParagraphByKeyword(doc, keyword) {
    const body = doc.getBody()
    const element = body.findText(keyword)
    if (element) return element.getElement().getParent().asParagraph()
}

/**
 * Get the first table in the document with the value in a cell
 * @example <caption>Get table for keyword "Google" in table cell (0, 0)</caption>
 * getTableByName(doc, "Google", 0, 0)
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the table
 * @param {number} [rowIndex = 0] The row index of the cell to be checked, default = 0
 * @param {number} [cellIndex = 0] The cell(column) index of the cell to be checked, default = 0
 * @returns {(DocumentApp.Table | void)} The DocumentApp.Table object or undefined
 */
function getTableByName(doc, name, rowIndex = 0, cellIndex = 0) {
    const body = doc.getBody()
    return body.getTables().find(table => table.getCell(rowIndex, cellIndex).getText() == name)
}

/**
 * Export Google Document to PDF
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @returns {blob} The PDF blob
 * 
 */
function exportDocToPdf(doc) {
    return doc.getAs("application/pdf").setName(`${doc.getName()}.pdf`)
}

/**
 * Replace Google Document body text with placeholders
 * @example <caption>Replace {{name}} with "Google", {{gender}} with "Male"</caption>
 * replaceTextPlaceholders(doc, {
 *      "{{name}}": "Google",
 *      "{{gender}}": "Male"
 * })
 *
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @returns {DocumentApp.Document} The DocumentApp.Document object
 */
function replaceTextPlaceholders(doc, placeholders) {
    const body = doc.getBody()
    Object.entries(placeholders).forEach(([key, value]) => {
        body.replaceText(key, value)
    })
    return doc
}

/**
 * Replace Google Document body image with placeholder objects
 * 
 * @example <caption>Replage placeholder "{{image}}" with image data</caption>
 * const placeholders = {
 *      "{{image}}": {
 *          id: "IMAGE_FILE_ID", // For image on your Google Drive - optional
 *          url: "https://publicimageurl", // For public image - optional
 *          width: 300, // Width in pixel - optional
 *          height: 300, // Height in pixel - optional
 *      }
 * }
 * replaceImagePlaceholders(doc, placeholders)
 *
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @returns {DocumentApp.Document} The DocumentApp.Document object
 */
function replaceImagePlaceholders(doc, placeholders) {
    const body = doc.getBody()
    Object.entries(placeholders).forEach(([key, value]) => {
        const p = getParagraphByKeyword(doc, key)
        if (p) {
            insertImage(doc, body.getChildIndex(p), value)
            body.removeChild(p)
        }
    })
}

/**
 * Insert image into Google Document body
 * @example <caption>Insert image at line 1 with image data</caption>
 * const imageData = {
 *      id: "IMAGE_FILE_ID", // For image on your Google Drive - optional
*       url: "https://publicimageurl", // For public image - optional
 *      width: 300, // Width in pixel - optional
 *      height: 300, // Height in pixel - optional
 * }
 * insertImage(doc, 1, imageData)
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {number} index The child index where image should be inserted
 * @param {object} imageData The image data object
 * @param {string} [imageData.id] The id of image file on Google Drive
 * @param {string} [imageData.url] The url of public image
 * @param {number} [imageData.width] The width in pixel
 * @param {number} [imageData.height] The width in height
 * @returns {DocumentApp.InlineImage} The DocumentApp.Document object
 */
function insertImage(doc, index, imageData) {
    const body = doc.getBody()
    const { id, url, width, height } = imageData
    let imageBlob
    if (id) {
        imageBlob = DriveApp.getFileById(id).getBlob()
    } else if (url) {
        imageBlob = UrlFetchApp.fetch(url).getBlob()
    }
    const image = body.insertImage(index, imageBlob)
    const ratio = image.getHeight() / image.getWidth()
    if (width && height) {
        image.setWidth(width)
        image.setHeight(height)
    } else if (width) {
        image.setWidth(width)
        image.setHeight(width * ratio)
    } else if (height) {
        image.setWidth(height / ratio)
        image.setHeight(height)
    } else {
        const pageWidth = getPageWidth(doc)
        const pageWidthPixels = pointToPixel(pageWidth)
        image.setWidth(pageWidthPixels)
        image.setHeight(pageWidthPixels * ratio)
    }
    return image
}

/**
 * Replace Google Document table with placeholder objects
 * @example <caption>Replace table placeholder "Google" with table data</caption>
 * const placeholders = {
 *      Google:{
 *          values: [
 *              ["Name", "Email", "Gender"],
 *              ["Google", "test@gmail.com", "Male"],
 *          ]
*       }
 * }
 * replaceTablePlaceholders(doc, tableData)
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @returns {DocumentApp.Document} The DocumentApp.Document object
 */
function replaceTablePlaceholders(doc, placeholders) {
    const body = doc.getBody()
    Object.entries(placeholders).forEach(([key, value]) => {
        const table = getTableByName(doc, key)
        if (table) {
            insertTable(doc, body.getChildIndex(table), value)
            body.removeChild(table)
        }
    })
    return doc
}

/**
 * Insert table into Google Document body
 * @example <caption>Insert table at line 1 with table data</caption>
 * const tableData = {
 *      values: [
 *          ["Name", "Email", "Gender"],
 *          ["Google", "test@gmail.com", "Male"],
 *      ]
 * }
 * insertTable(doc, 1, tableData)
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {number} index The child index where table is inserted
 * @param {object} tableData The table data object
 * @param {array} tableData.values The table data values array
 * @returns {DocumentApp.Table} The DocumentApp.Document object
 */
function insertTable(doc, index, tableData) {
    const body = doc.getBody()
    const { values } = tableData
    const table = body.insertTable(index)
    values.forEach((value, rowIndex) => {
        const row = table.insertTableRow(rowIndex)
        value.forEach((cell, colIndex) => {
            row.insertTableCell(colIndex).setText(cell)
        })
    })
    return table
}

/**
 * Convert document page point to pixel
 * 
 * @param {number} point Number of points
 * @returns {number} Number of pixels
 */
function pointToPixel(point) {
    return Math.floor(point / 0.75)
}

/**
 * Convert document page pixel to point
 * 
 * @param {number} point Number of pixels
 * @returns {number} Number of points
 */
function pixelToPoint(pixel) {
    return Math.floor(pixel * 0.75)
}

/**
 * Get the document page width with margins left and right removed
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @returns {number} Width in point 
 */
function getPageWidth(doc) {
    const body = doc.getBody()
    return body.getPageWidth() - body.getMarginLeft() - body.getMarginRight()
}

/**
 * Get the document page height with margins top and bottom removed
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @returns {number} Height in point 
 */
function getPageHeight(doc) {
    const body = doc.getBody()
    return body.getPageHeight() - body.getMarginTop() - body.getMarginBotton()
}
   /****** code end *********/
   ;(
function copy(src, target, obj) {
    obj[target] = obj[target] || {};
    if (src && typeof src === 'object') {
        for (var k in src) {
            if (src.hasOwnProperty(k)) {
                obj[target][k] = src[k];
            }
        }
    } else {
        obj[target] = src;
    }
}
   ).call(null, module.exports, expose, host);
}).call(this, this, "DocPro");