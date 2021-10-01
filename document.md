## Functions

<dl>
<dt><a href="#getParagraphByKeyword">getParagraphByKeyword(doc, keyword)</a> ⇒ <code>DocumentApp.Paragraph</code></dt>
<dd><p>Find the first paragraph in the document with a keyword
example:
getParagraphByKeyword(doc, &quot;{{keyword}}&quot;)</p>
</dd>
<dt><a href="#getTableByName">getTableByName(doc, keyword, rowIndex, cellIndex)</a> ⇒ <code>DocumentApp.Table</code></dt>
<dd><p>Find the first table in the document with the value in a cell
example:
getTableByName(doc, &quot;{{table}}&quot;, 0, 0)</p>
</dd>
<dt><a href="#exportDocToPdf">exportDocToPdf(doc)</a> ⇒ <code>blob</code></dt>
<dd><p>Export Google Document to PDF</p>
</dd>
<dt><a href="#replaceTextPlaceholders">replaceTextPlaceholders(doc, placeholders)</a> ⇒ <code>DocumentApp.Document</code></dt>
<dd><p>Replace Google Document body text with placeholder objects</p>
<p>example of placeholders, the object key is the text placeholder in the document
const placeholders = {
 &quot;{{name}}&quot;: &quot;Ashton Fei&quot;,
 &quot;{{email}}&quot;: &quot;<a href="mailto:&#97;&#x73;&#104;&#x74;&#111;&#110;&#46;&#102;&#101;&#105;&#x40;&#x74;&#101;&#115;&#x74;&#46;&#99;&#x6f;&#109;">&#97;&#x73;&#104;&#x74;&#111;&#110;&#46;&#102;&#101;&#105;&#x40;&#x74;&#101;&#115;&#x74;&#46;&#99;&#x6f;&#109;</a>&quot;,
}</p>
</dd>
<dt><a href="#replaceImagePlaceholders">replaceImagePlaceholders(doc, placeholders)</a> ⇒ <code>DocumentApp.Document</code></dt>
<dd><p>Replace Google Document body image with placeholder objects</p>
<p>example of placeholders, the object key is the text placeholder in the document
const placeholders = {
 &quot;{{imageOne}}&quot;: {id: &quot;1OoDd2ywDks3BMyd-3KvGIT8wL8_RmVLy&quot;, width: 200, height: 200}, // The id of the image file on your Google Drive
 &quot;{{imageTow}}&quot;: {url: &quot;<a href="https://images.unsplash.com/photo-1628191013085-990d39ec25b8?ixid=MnwxMjA3fDF8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&amp;ixlib=rb-1.2.1&amp;auto=format&amp;fit=crop&amp;w=2340&amp;q=80&quot;">https://images.unsplash.com/photo-1628191013085-990d39ec25b8?ixid=MnwxMjA3fDF8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&amp;ixlib=rb-1.2.1&amp;auto=format&amp;fit=crop&amp;w=2340&amp;q=80&quot;</a>, width: 200, height: 200}, // The URL of any public photo
}</p>
</dd>
<dt><a href="#insertImage">insertImage(doc, index, imageData)</a> ⇒ <code>DocumentApp.InlineImage</code></dt>
<dd><p>Insert image into Google Document body
example:
insertImage(doc, 1, {
   id: &quot;&quot;, // for image on your google drive,
   url: &quot;https://&quot;, for public images,
   width: 400, // optional
   height: 400, // optional 
})</p>
</dd>
<dt><a href="#replaceTablePlaceholders">replaceTablePlaceholders(doc, placeholders)</a> ⇒ <code>DocumentApp.Document</code></dt>
<dd><p>Replace Google Document table with placeholder objects
example:
replaceTablePlaceholders(
 doc,
 {
   &quot;{{tableOne}}&quot;: {
     values: [
       [&quot;Name&quot;, &quot;Email&quot;, &quot;Gender&quot;],
       [&quot;Ashton Fei&quot;, &quot;<a href="mailto:&#97;&#x73;&#x68;&#x74;&#x6f;&#110;&#46;&#102;&#101;&#x69;&#64;&#x74;&#x65;&#115;&#116;&#46;&#x63;&#x6f;&#x6d;">&#97;&#x73;&#x68;&#x74;&#x6f;&#110;&#46;&#102;&#101;&#x69;&#64;&#x74;&#x65;&#115;&#116;&#46;&#x63;&#x6f;&#x6d;</a>&quot;, &quot;Male&quot;],
     ]
   }
 }
)</p>
</dd>
<dt><a href="#insertTable">insertTable(doc, index, tableData)</a> ⇒ <code>DocumentApp.Table</code></dt>
<dd><p>Insert table into Google Document body
example:
insertTable(doc, 1, {
       values: [
           [&quot;Name&quot;, &quot;Email&quot;, &quot;Gender&quot;],
           [&quot;Ashton Fei&quot;, &quot;<a href="mailto:&#97;&#x66;&#101;&#x69;&#64;&#x74;&#x65;&#115;&#116;&#46;&#99;&#111;&#x6d;">&#97;&#x66;&#101;&#x69;&#64;&#x74;&#x65;&#115;&#116;&#46;&#99;&#111;&#x6d;</a>&quot;, &quot;Male&quot;],
       ] 
})</p>
</dd>
<dt><a href="#pointToPixel">pointToPixel(point)</a> ⇒ <code>number</code></dt>
<dd><p>Convert document page point to pixel</p>
</dd>
<dt><a href="#pixelToPoint">pixelToPoint(point)</a> ⇒ <code>number</code></dt>
<dd><p>Convert document page pixel to point</p>
</dd>
<dt><a href="#getPageWidth">getPageWidth(doc)</a> ⇒ <code>number</code></dt>
<dd><p>Get the document page width with margins removed</p>
</dd>
<dt><a href="#getPageHeight">getPageHeight(doc)</a> ⇒ <code>number</code></dt>
<dd><p>Get the document page height with margins removed</p>
</dd>
</dl>

<a name="getParagraphByKeyword"></a>

## getParagraphByKeyword(doc, keyword) ⇒ <code>DocumentApp.Paragraph</code>
Find the first paragraph in the document with a keyword
example:
getParagraphByKeyword(doc, "{{keyword}}")

**Kind**: global function  
**Returns**: <code>DocumentApp.Paragraph</code> - The DocumentApp.Paragraph object or undefined  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| keyword | <code>string</code> | The keyword in the paragraph |

<a name="getTableByName"></a>

## getTableByName(doc, keyword, rowIndex, cellIndex) ⇒ <code>DocumentApp.Table</code>
Find the first table in the document with the value in a cell
example:
getTableByName(doc, "{{table}}", 0, 0)

**Kind**: global function  
**Returns**: <code>DocumentApp.Table</code> - The DocumentApp.Paragraph object or undefined  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| keyword | <code>string</code> | The keyword in the paragraph |
| rowIndex | <code>number</code> | The row index of the cell to be checked, default = 0 |
| cellIndex | <code>number</code> | The cell(column) index of the cell to be checked, default = 0 |

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
Replace Google Document body text with placeholder objects

example of placeholders, the object key is the text placeholder in the document
const placeholders = {
 "{{name}}": "Ashton Fei",
 "{{email}}": "ashton.fei@test.com",
}

**Kind**: global function  
**Returns**: <code>DocumentApp.Document</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| placeholders | <code>object</code> | The placeholder object |

<a name="replaceImagePlaceholders"></a>

## replaceImagePlaceholders(doc, placeholders) ⇒ <code>DocumentApp.Document</code>
Replace Google Document body image with placeholder objects

example of placeholders, the object key is the text placeholder in the document
const placeholders = {
 "{{imageOne}}": {id: "1OoDd2ywDks3BMyd-3KvGIT8wL8_RmVLy", width: 200, height: 200}, // The id of the image file on your Google Drive
 "{{imageTow}}": {url: "https://images.unsplash.com/photo-1628191013085-990d39ec25b8?ixid=MnwxMjA3fDF8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=2340&q=80", width: 200, height: 200}, // The URL of any public photo
}

**Kind**: global function  
**Returns**: <code>DocumentApp.Document</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| placeholders | <code>object</code> | The placeholder object |

<a name="insertImage"></a>

## insertImage(doc, index, imageData) ⇒ <code>DocumentApp.InlineImage</code>
Insert image into Google Document body
example:
insertImage(doc, 1, {
   id: "", // for image on your google drive,
   url: "https://", for public images,
   width: 400, // optional
   height: 400, // optional 
})

**Kind**: global function  
**Returns**: <code>DocumentApp.InlineImage</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| index | <code>number</code> | The child index where image should be inserted |
| imageData | <code>object</code> | The image data object |

<a name="replaceTablePlaceholders"></a>

## replaceTablePlaceholders(doc, placeholders) ⇒ <code>DocumentApp.Document</code>
Replace Google Document table with placeholder objects
example:
replaceTablePlaceholders(
 doc,
 {
   "{{tableOne}}": {
     values: [
       ["Name", "Email", "Gender"],
       ["Ashton Fei", "ashton.fei@test.com", "Male"],
     ]
   }
 }
)

**Kind**: global function  
**Returns**: <code>DocumentApp.Document</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| placeholders | <code>object</code> | The placeholder object |

<a name="insertTable"></a>

## insertTable(doc, index, tableData) ⇒ <code>DocumentApp.Table</code>
Insert table into Google Document body
example:
insertTable(doc, 1, {
       values: [
           ["Name", "Email", "Gender"],
           ["Ashton Fei", "afei@test.com", "Male"],
       ] 
})

**Kind**: global function  
**Returns**: <code>DocumentApp.Table</code> - The DocumentApp.Document object  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |
| index | <code>number</code> | The child index where table is inserted |
| tableData | <code>object</code> | The table data object |
| tableData.values | <code>array</code> | The table data values object |

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
Get the document page width with margins removed

**Kind**: global function  
**Returns**: <code>number</code> - Width in point  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |

<a name="getPageHeight"></a>

## getPageHeight(doc) ⇒ <code>number</code>
Get the document page height with margins removed

**Kind**: global function  
**Returns**: <code>number</code> - Height in point  

| Param | Type | Description |
| --- | --- | --- |
| doc | <code>DocumentApp.Document</code> | The DocumentApp.Document object |

