## Functions

<dl>
<dt><a href="#getParagraphByKeyword">getParagraphByKeyword(doc, keyword)</a> ⇒ <code>DocumentApp.Paragraph</code> | <code>void</code></dt>
<dd><p>Get the first paragraph in the document with a keyword</p>
</dd>
<dt><a href="#getTableByName">getTableByName(doc, keyword, [rowIndex], [cellIndex])</a> ⇒ <code>DocumentApp.Table</code> | <code>void</code></dt>
<dd><p>Get the first table in the document with the value in a cell</p>
</dd>
<dt><a href="#exportDocToPdf">exportDocToPdf(doc)</a> ⇒ <code>blob</code></dt>
<dd><p>Export Google Document to PDF</p>
</dd>
<dt><a href="#replaceTextPlaceholders">replaceTextPlaceholders(doc, placeholders)</a> ⇒ <code>DocumentApp.Document</code></dt>
<dd><p>Replace Google Document body text with placeholders</p>
</dd>
<dt><a href="#replaceImagePlaceholders">replaceImagePlaceholders(doc, placeholders)</a> ⇒ <code>DocumentApp.Document</code></dt>
<dd><p>Replace Google Document body image with placeholder objects</p>
</dd>
<dt><a href="#insertImage">insertImage(doc, index, imageData)</a> ⇒ <code>DocumentApp.InlineImage</code></dt>
<dd><p>Insert image into Google Document body</p>
</dd>
<dt><a href="#replaceTablePlaceholders">replaceTablePlaceholders(doc, placeholders)</a> ⇒ <code>DocumentApp.Document</code></dt>
<dd><p>Replace Google Document table with placeholder objects</p>
</dd>
<dt><a href="#insertTable">insertTable(doc, index, tableData)</a> ⇒ <code>DocumentApp.Table</code></dt>
<dd><p>Insert table into Google Document body</p>
</dd>
<dt><a href="#pointToPixel">pointToPixel(point)</a> ⇒ <code>number</code></dt>
<dd><p>Convert document page point to pixel</p>
</dd>
<dt><a href="#pixelToPoint">pixelToPoint(point)</a> ⇒ <code>number</code></dt>
<dd><p>Convert document page pixel to point</p>
</dd>
<dt><a href="#getPageWidth">getPageWidth(doc)</a> ⇒ <code>number</code></dt>
<dd><p>Get the document page width with margins left and right removed</p>
</dd>
<dt><a href="#getPageHeight">getPageHeight(doc)</a> ⇒ <code>number</code></dt>
<dd><p>Get the document page height with margins top and bottom removed</p>
</dd>
</dl>

<a name="getParagraphByKeyword"></a>

## getParagraphByKeyword(doc, keyword) ⇒ <code>DocumentApp.Paragraph</code> \| <code>void</code>
Get the first paragraph in the document with a keyword

**Kind**: global function  
**Returns**: <code>DocumentApp.Paragraph</code> \| <code>void</code> - The DocumentApp.Paragraph object or undefined when keyword not found  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| keyword | <code>string</code> | The keyword in the paragraph |

**Example** *(Get paragrahp with keyword &#x27;Google&#x27;)*  
```js
getParagraphByKeyword(doc, "Google")
```
<a name="getTableByName"></a>

## getTableByName(doc, keyword, [rowIndex], [cellIndex]) ⇒ <code>DocumentApp.Table</code> \| <code>void</code>
Get the first table in the document with the value in a cell

**Kind**: global function  
**Returns**: <code>DocumentApp.Table</code> \| <code>void</code> - The DocumentApp.Table object or undefined  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| doc | <code>DocumentApp.Document</code> |  | The DocumentApp.Document object |
| keyword | <code>string</code> |  | The keyword in the table |
| [rowIndex] | <code>number</code> | <code>0</code> | The row index of the cell to be checked, default = 0 |
| [cellIndex] | <code>number</code> | <code>0</code> | The cell(column) index of the cell to be checked, default = 0 |

**Example** *(Get table for keyword &quot;Google&quot; in table cell (0, 0))*  
```js
getTableByName(doc, "Google", 0, 0)
```
<a name="exportDocToPdf"></a>

## exportDocToPdf(doc) ⇒ <code>blob</code>
Export Google Document to PDF

**Kind**: global function  
**Returns**: <code>blob</code> - The PDF blob  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |

<a name="replaceTextPlaceholders"></a>

## replaceTextPlaceholders(doc, placeholders) ⇒ <code>DocumentApp.Document</code>
Replace Google Document body text with placeholders

**Kind**: global function  
**Returns**: <code>DocumentApp.Document</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| placeholders | <code>object</code> | The placeholder object |

**Example** *(Replace {{name}} with &quot;Google&quot;, {{gender}} with &quot;Male&quot;)*  
```js
replaceTextPlaceholders(doc, {
     "{{name}}": "Google",
     "{{gender}}": "Male"
})
```
<a name="replaceImagePlaceholders"></a>

## replaceImagePlaceholders(doc, placeholders) ⇒ <code>DocumentApp.Document</code>
Replace Google Document body image with placeholder objects

**Kind**: global function  
**Returns**: <code>DocumentApp.Document</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| placeholders | <code>object</code> | The placeholder object |

**Example** *(Replage placeholder &quot;{{image}}&quot; with image data)*  
```js
const placeholders = {
     "{{image}}": {
         id: "IMAGE_FILE_ID", // For image on your Google Drive - optional
         url: "https://publicimageurl", // For public image - optional
         width: 300, // Width in pixel - optional
         height: 300, // Height in pixel - optional
     }
}
replaceImagePlaceholders(doc, placeholders)
```
<a name="insertImage"></a>

## insertImage(doc, index, imageData) ⇒ <code>DocumentApp.InlineImage</code>
Insert image into Google Document body

**Kind**: global function  
**Returns**: <code>DocumentApp.InlineImage</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| index | <code>number</code> | The child index where image should be inserted |
| imageData | <code>object</code> | The image data object |
| [imageData.id] | <code>string</code> | The id of image file on Google Drive |
| [imageData.url] | <code>string</code> | The url of public image |
| [imageData.width] | <code>number</code> | The width in pixel |
| [imageData.height] | <code>number</code> | The width in height |

**Example** *(Insert image at line 1 with image data)*  
```js
const imageData = {
     id: "IMAGE_FILE_ID", // For image on your Google Drive - optional
      url: "https://publicimageurl", // For public image - optional
     width: 300, // Width in pixel - optional
     height: 300, // Height in pixel - optional
}
insertImage(doc, 1, imageData)
```
<a name="replaceTablePlaceholders"></a>

## replaceTablePlaceholders(doc, placeholders) ⇒ <code>DocumentApp.Document</code>
Replace Google Document table with placeholder objects

**Kind**: global function  
**Returns**: <code>DocumentApp.Document</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| placeholders | <code>object</code> | The placeholder object |

**Example** *(Replace table placeholder &quot;Google&quot; with table data)*  
```js
const placeholders = {
     Google:{
         values: [
             ["Name", "Email", "Gender"],
             ["Google", "test@gmail.com", "Male"],
         ]
      }
}
replaceTablePlaceholders(doc, tableData)
```
<a name="insertTable"></a>

## insertTable(doc, index, tableData) ⇒ <code>DocumentApp.Table</code>
Insert table into Google Document body

**Kind**: global function  
**Returns**: <code>DocumentApp.Table</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| index | <code>number</code> | The child index where table is inserted |
| tableData | <code>object</code> | The table data object |
| tableData.values | <code>array</code> | The table data values array |

**Example** *(Insert table at line 1 with table data)*  
```js
const tableData = {
     values: [
         ["Name", "Email", "Gender"],
         ["Google", "test@gmail.com", "Male"],
     ]
}
insertTable(doc, 1, tableData)
```
<a name="pointToPixel"></a>

## pointToPixel(point) ⇒ <code>number</code>
Convert document page point to pixel

**Kind**: global function  
**Returns**: <code>number</code> - Number of pixels  

| Param | Type | Description |
| --- | --- | --- |
| point | <code>number</code> | Number of points |

<a name="pixelToPoint"></a>

## pixelToPoint(point) ⇒ <code>number</code>
Convert document page pixel to point

**Kind**: global function  
**Returns**: <code>number</code> - Number of points  

| Param | Type | Description |
| --- | --- | --- |
| point | <code>number</code> | Number of pixels |

<a name="getPageWidth"></a>

## getPageWidth(doc) ⇒ <code>number</code>
Get the document page width with margins left and right removed

**Kind**: global function  
**Returns**: <code>number</code> - Width in point  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |

<a name="getPageHeight"></a>

## getPageHeight(doc) ⇒ <code>number</code>
Get the document page height with margins top and bottom removed

**Kind**: global function  
**Returns**: <code>number</code> - Height in point  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |

