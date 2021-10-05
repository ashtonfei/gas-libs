(function (host, expose) {
   var module = { exports: {} };
   var exports = module.exports;
   /****** code begin *********/
/**
MIT License

Copyright (c) 2021 Ashton Fei

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */


/**
 * Get the first image by the title
 * 
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {string} The title of the image in the slide
 * @returns {SlidesApp.Image|void} The first SlidesApp.Image or undefined
 */
function getImageByTitle(presentation, title) {
  const slides = presentation.getSlides()
  slides.forEach(slide => {
    slide.getImages().forEach(image => {
      if (image.getTitle() === title) return image
    })
  })
}

/**
 * Get all images by the title
 * 
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {string} The title of the image in the slide
 * @returns {SlidesApp.Image[]} An array of SlidesApp.Image
 */
function getImagesByTitle(presentation, title) {
  const images = []
  const slides = presentation.getSlides()
  slides.forEach(slide => {
    slide.getImages().forEach(image => {
      if (image.getTitle() === title) images.push(image)
    })
  })
  return images
}

/**
 * Get the first table by header text
 * 
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {string} header The text of the header
 * @param {number} [row = 1] The row number of the header
 * @param {number} [column = 1] The column number of the header
 * @returns {SlidesApp.Table|void} The first SlidesApp.Table or undefined
 */
function getTableByHeader(presentation, header, row = 0, column = 0) {
  presentation.getSlides().forEach(slide => {
    slide.getTables().forEach(table => {
      const cellValue = table.getCell(row, column).getText().asString().replace(/\n+/, "")
      if (cellValue === header) return table
    })
  })
}

/**
 * Get all tables by header text
 * 
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {string} header The text of the header
 * @param {number} [row = 1] The row number of the header
 * @param {number} [column = 1] The column number of the header
 * @returns {SlidesApp.Table[]} An array of SlidesApp.Table
 */
function getTablesByHeader(presentation, header, row = 0, column = 0) {
  const tables = []
  presentation.getSlides().forEach(slide => {
    slide.getTables().forEach(table => {
      const cellValue = table.getCell(row, column).getText().asString().replace(/\n+/, "")
      if (cellValue === header) {
        tables.push(table)
      }
    })
  })
  return tables
}

/**
 * Update an image with image data
 * 
 * @param {SlidesApp.Image} image The SlidesApp.Image object
 * @param {object} data The image data object
 * @param {string} [data.id] The id of image on Google Drive
 * @param {string} [data.url] The url of public image
 * @param {string} [data.link] The url to be linked to the image
 * @param {boolean} [data.crop] Crop the image
 * @returns {SlidesApp.Image} The SlidesApp.Image object after updated
 */
function updateImage(image, data) {
  if (data.id) {
    image.replace(DriveApp.getFileById(data.id, data.crop))
  } else if (data.url) {
    image.replace(data.url, data.crop)
  }
  if (data.link) {
    image.setLinkUrl(data.link)
  }
  return image
}

/**
 * Update a table with table data
 * 
 * @param {SlidesApp.Table} table The SlidesApp.Table object
 * @param {object[][]} data The table data object
 * @param {string|number} [data[][].value] The value of the table cell
 * @param {string} [data[][].color] The font color of the table cell
 * @param {string} [data[][].bgColor] The background color of the table cell
 * @param {string} [data[][].link] The url to be linked to the cell
 *  
 */
function updateTable(table, data) {
  const page = table.getParentPage()
  const left = table.getLeft()
  const top = table.getTop()
  const newTable = page.insertTable(data.length, data[0].length)
    .setTop(top)
    .setLeft(left)
  data.forEach((rowValues, rowIndex) => {
    rowValues.forEach(({ value, bgColor, color, link }, columnIndex) => {
      const cell = newTable.getCell(rowIndex, columnIndex)
      cell.getText().setText(value)
      if (link) cell.getText().getTextStyle().setLinkUrl(link)
      if (color) cell.getText().getTextStyle().setForegroundColor(color)
      if (bgColor) cell.getFill().setSolidFill(bgColor)
    })
  })
  table.remove()
  return newTable
}

/**
 * Replace text placeholders for the all slides in the presentation
 * @example <caption>Replace {{name}} with "Google" in all slides</caption>
 * const presentation = SlidesApp.getActivePresentation()
 * const placeholders = {
 *  "{{name}}": "Google",
 * }
 * replaceTextPlaceholders(presentation, placeholders)
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {object} placeholders The text placeholders object
 * @return {SlidesApp.Presentation} The SlidesApp.Presentation object after udpate
 */
function replaceTextPlaceholders(presentation, placeholders) {
  const slides = presentation.getSlides()
  slides.forEach(slide => {
    Object.entries(placeholders).forEach(([key, value]) => {
      console.log({
        key,
        value
      })
      slide.replaceAllText(key, value)
    })
  })
  return presentation
}

/**
 * Replace image placeholders for the all slides in the presentation
 * @example <caption>Replace image with title {{logo}} with image from url https://fakeimageurl</caption>
 * const presentation = SlidesApp.getActivePresentation()
 * const placeholders = {
 *  "{{logo}}": {
 *      url: "https://fakeimageurl",
 *      id: null,
 *      crop: true,
 *   }
 * }
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {object} placeholders The image placeholders object
 * @return {SlidesApp.Presentation} The SlidesApp.Presentation object after udpate
 */
function replaceImagePlaceholders(presentation, placeholders) {
  Object.entries(placeholders).forEach(([key, data]) => {
    const images = getImagesByTitle(presentation, key)
    images.forEach(image => updateImage(image, data))
  })
  return presentation
}

/**
 * Update table placeholders for all slides in the presentation
 * 
 * @param {SlidesApp.Presentation} presentation The SlidesApp.Presentation object
 * @param {object} placeholders The table placeholders object
 * @return {SlidesApp.Presentation} The SlidesApp.Presentation object after udpate
 */
function replaceTablePlaceholders(presentation, placeholders) {
  Object.entries(placeholders).forEach(([key, data]) => {
    const tables = getTablesByHeader(presentation, key)
    tables.forEach(table => updateTable(table, data))
  })
}



/**Export functions for SlidePro library */
if (typeof module === "object") {
  module.exports = {
    getImageByTitle,
    getImagesByTitle,
    getTableByHeader,
    getTablesByHeader,
    updateImage,
    updateTable,
    replaceTextPlaceholders,
    replaceImagePlaceholders,
    replaceTablePlaceholders,
  }
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
}).call(this, this, "SlidePro");