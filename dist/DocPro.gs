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
 * Find the first paragraph in the document with a keyword
 * example:
 * getParagraphByKeyword(doc, "{{keyword}}")
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the paragraph
 * @return {DocumentApp.Paragraph} The DocumentApp.Paragraph object or undefined
 */
function getParagraphByKeyword(doc, keyword) {
    const body = doc.getBody()
    const element = body.findText(keyword)
    if (element) return element.getElement().getParent().asParagraph()
}

/**
 * Find the first table in the document with the value in a cell
 * example:
 * getTableByName(doc, "{{table}}", 0, 0)
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the paragraph
 * @param {number?} rowIndex The row index of the cell to be checked, default = 0
 * @param {number?} cellIndex The cell(column) index of the cell to be checked, default = 0
 * @return {DocumentApp.Table} The DocumentApp.Paragraph object or undefined
 */
function getTableByName(doc, name, rowIndex = 0, cellIndex = 0) {
    const body = doc.getBody()
    return body.getTables().find(table => table.getCell(rowIndex, cellIndex).getText() == name)
}

/**
 * Export Google Document to PDF
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @return {blob} The PDF blob
 */
function exportDocToPdf(doc) {
    return doc.getAs("application/pdf").setName(`${doc.getName()}.pdf`)
}

/**
 * Replace Google Document body text with placeholder objects
 * 
 * example of placeholders, the object key is the text placeholder in the document
 * const placeholders = {
 *  "{{name}}": "Ashton Fei",
 *  "{{email}}": "ashton.fei@test.com",
 * }
 *
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @return {DocumentApp.Document} The DocumentApp.Document object
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
 * example of placeholders, the object key is the text placeholder in the document
 * const placeholders = {
 *  "{{imageOne}}": {id: "1OoDd2ywDks3BMyd-3KvGIT8wL8_RmVLy", width: 200, height: 200}, // The id of the image file on your Google Drive
 *  "{{imageTow}}": {url: "https://images.unsplash.com/photo-1628191013085-990d39ec25b8?ixid=MnwxMjA3fDF8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=2340&q=80", width: 200, height: 200}, // The URL of any public photo
 * }
 *
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @return {DocumentApp.Document} The DocumentApp.Document object
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
 * example:
 * insertImage(doc, 1, {
 *    id: "", // for image on your google drive,
 *    url: "https://", for public images,
 *    width: 400, // optional
 *    height: 400, // optional 
 * })
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {number} index The child index where image should be inserted
 * @param {object} imageData The image data object
 * @return {DocumentApp.InlineImage} The DocumentApp.Document object
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
    if (width) image.setWidth(width)
    if (height) image.setHeight(height)
    if (!(width && height)) {
        const pageWidth = getPageWidth(doc)
        const pageWidthPixels = pointToPixel(pageWidth)
        const ratio = image.getHeight() / image.getWidth()
        image.setWidth(pageWidthPixels)
        image.setHeight(pageWidthPixels * ratio)
    }
    return image
}

/**
 * Replace Google Document table with placeholder objects
 * example:
 * replaceTablePlaceholders(
 *  doc,
 *  {
 *    "{{tableOne}}": {
 *      values: [
 *        ["Name", "Email", "Gender"],
 *        ["Ashton Fei", "ashton.fei@test.com", "Male"],
 *      ]
 *    }
 *  }
 * )
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @return {DocumentApp.Document} The DocumentApp.Document object
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
 * example:
 * insertTable(doc, 1, {
 *        values: [
 *            ["Name", "Email", "Gender"],
 *            ["Ashton Fei", "afei@test.com", "Male"],
 *        ] 
 * })
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {number} index The child index where table is inserted
 * @param {object} tableData The table data object
 * @param {array} tableData.values The table data values object
 * @return {DocumentApp.Table} The DocumentApp.Document object
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
 * @return {number} Number of pixels
 */
function pointToPixel(point) {
    return Math.floor(point / 0.75)
}

/**
 * Convert document page pixel to point
 * 
 * @param {number} point Number of pixels
 * @return {number} Number of points
 */
function pixelToPoint(pixel) {
    return MailApp.floor(pixel * 0.75)
}

/**
 * Get the document page width with margins removed
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @return {number} Width in point 
 */
function getPageWidth(doc) {
    const body = doc.getBody()
    return body.getPageWidth() - body.getMarginLeft() - body.getMarginRight()
}

/**
 * Get the document page height with margins removed
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @return {number} Height in point 
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