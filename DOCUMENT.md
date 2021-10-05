## Functions

<dl>
<dt><a href="#getImageByTitle">getImageByTitle(presentation, The)</a> ⇒ <code>SlidesApp.Image</code> | <code>void</code></dt>
<dd><p>Get the first image by the title</p>
</dd>
<dt><a href="#getImagesByTitle">getImagesByTitle(presentation, The)</a> ⇒ <code>Array.&lt;SlidesApp.Image&gt;</code></dt>
<dd><p>Get all images by the title</p>
</dd>
<dt><a href="#getTableByHeader">getTableByHeader(presentation, header, [row], [column])</a> ⇒ <code>SlidesApp.Table</code> | <code>void</code></dt>
<dd><p>Get the first table by header text</p>
</dd>
<dt><a href="#getTablesByHeader">getTablesByHeader(presentation, header, [row], [column])</a> ⇒ <code>Array.&lt;SlidesApp.Table&gt;</code></dt>
<dd><p>Get all tables by header text</p>
</dd>
<dt><a href="#updateImage">updateImage(image, data)</a> ⇒ <code>SlidesApp.Image</code></dt>
<dd><p>Update an image with image data</p>
</dd>
<dt><a href="#updateTable">updateTable(table, data)</a></dt>
<dd><p>Update a table with table data</p>
</dd>
<dt><a href="#replaceTextPlaceholders">replaceTextPlaceholders(presentation, placeholders)</a> ⇒ <code>SlidesApp.Presentation</code></dt>
<dd><p>Replace text placeholders for the all slides in the presentation</p>
</dd>
<dt><a href="#replaceImagePlaceholders">replaceImagePlaceholders(presentation, placeholders)</a> ⇒ <code>SlidesApp.Presentation</code></dt>
<dd><p>Replace image placeholders for the all slides in the presentation</p>
</dd>
<dt><a href="#replaceTablePlaceholders">replaceTablePlaceholders(presentation, placeholders)</a> ⇒ <code>SlidesApp.Presentation</code></dt>
<dd><p>Update table placeholders for all slides in the presentation</p>
</dd>
</dl>

<a name="getImageByTitle"></a>

## getImageByTitle(presentation, The) ⇒ <code>SlidesApp.Image</code> \| <code>void</code>
Get the first image by the title

**Kind**: global function  
**Returns**: <code>SlidesApp.Image</code> \| <code>void</code> - The first SlidesApp.Image or undefined  

| Param | Type | Description |
| --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> | The SlidesApp.Presentation object |
| The | <code>string</code> | title of the image in the slide |

<a name="getImagesByTitle"></a>

## getImagesByTitle(presentation, The) ⇒ <code>Array.&lt;SlidesApp.Image&gt;</code>
Get all images by the title

**Kind**: global function  
**Returns**: <code>Array.&lt;SlidesApp.Image&gt;</code> - An array of SlidesApp.Image  

| Param | Type | Description |
| --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> | The SlidesApp.Presentation object |
| The | <code>string</code> | title of the image in the slide |

<a name="getTableByHeader"></a>

## getTableByHeader(presentation, header, [row], [column]) ⇒ <code>SlidesApp.Table</code> \| <code>void</code>
Get the first table by header text

**Kind**: global function  
**Returns**: <code>SlidesApp.Table</code> \| <code>void</code> - The first SlidesApp.Table or undefined  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> |  | The SlidesApp.Presentation object |
| header | <code>string</code> |  | The text of the header |
| [row] | <code>number</code> | <code>1</code> | The row number of the header |
| [column] | <code>number</code> | <code>1</code> | The column number of the header |

<a name="getTablesByHeader"></a>

## getTablesByHeader(presentation, header, [row], [column]) ⇒ <code>Array.&lt;SlidesApp.Table&gt;</code>
Get all tables by header text

**Kind**: global function  
**Returns**: <code>Array.&lt;SlidesApp.Table&gt;</code> - An array of SlidesApp.Table  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> |  | The SlidesApp.Presentation object |
| header | <code>string</code> |  | The text of the header |
| [row] | <code>number</code> | <code>1</code> | The row number of the header |
| [column] | <code>number</code> | <code>1</code> | The column number of the header |

<a name="updateImage"></a>

## updateImage(image, data) ⇒ <code>SlidesApp.Image</code>
Update an image with image data

**Kind**: global function  
**Returns**: <code>SlidesApp.Image</code> - The SlidesApp.Image object after updated  

| Param | Type | Description |
| --- | --- | --- |
| image | <code>SlidesApp.Image</code> | The SlidesApp.Image object |
| data | <code>object</code> | The image data object |
| [data.id] | <code>string</code> | The id of image on Google Drive |
| [data.url] | <code>string</code> | The url of public image |
| [data.link] | <code>string</code> | The url to be linked to the image |
| [data.crop] | <code>boolean</code> | Crop the image |

<a name="updateTable"></a>

## updateTable(table, data)
Update a table with table data

**Kind**: global function  

| Param | Type | Description |
| --- | --- | --- |
| table | <code>SlidesApp.Table</code> | The SlidesApp.Table object |
| data | <code>Array.&lt;Array.&lt;object&gt;&gt;</code> | The table data object |
| [data[][].value] | <code>string</code> \| <code>number</code> | The value of the table cell |
| [data[][].color] | <code>string</code> | The font color of the table cell |
| [data[][].bgColor] | <code>string</code> | The background color of the table cell |
| [data[][].link] | <code>string</code> | The url to be linked to the cell |

<a name="replaceTextPlaceholders"></a>

## replaceTextPlaceholders(presentation, placeholders) ⇒ <code>SlidesApp.Presentation</code>
Replace text placeholders for the all slides in the presentation

**Kind**: global function  
**Returns**: <code>SlidesApp.Presentation</code> - The SlidesApp.Presentation object after udpate  

| Param | Type | Description |
| --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> | The SlidesApp.Presentation object |
| placeholders | <code>object</code> | The text placeholders object |

**Example** *(Replace {{name}} with &quot;Google&quot; in all slides)*  
```js
const presentation = SlidesApp.getActivePresentation()
const placeholders = {
 "{{name}}": "Google",
}
replaceTextPlaceholders(presentation, placeholders)
```
<a name="replaceImagePlaceholders"></a>

## replaceImagePlaceholders(presentation, placeholders) ⇒ <code>SlidesApp.Presentation</code>
Replace image placeholders for the all slides in the presentation

**Kind**: global function  
**Returns**: <code>SlidesApp.Presentation</code> - The SlidesApp.Presentation object after udpate  

| Param | Type | Description |
| --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> | The SlidesApp.Presentation object |
| placeholders | <code>object</code> | The image placeholders object |

**Example** *(Replace image with title {{logo}} with image from url https://fakeimageurl)*  
```js
const presentation = SlidesApp.getActivePresentation()
const placeholders = {
 "{{logo}}": {
     url: "https://fakeimageurl",
     id: null,
     crop: true,
  }
}
```
<a name="replaceTablePlaceholders"></a>

## replaceTablePlaceholders(presentation, placeholders) ⇒ <code>SlidesApp.Presentation</code>
Update table placeholders for all slides in the presentation

**Kind**: global function  
**Returns**: <code>SlidesApp.Presentation</code> - The SlidesApp.Presentation object after udpate  

| Param | Type | Description |
| --- | --- | --- |
| presentation | <code>SlidesApp.Presentation</code> | The SlidesApp.Presentation object |
| placeholders | <code>object</code> | The table placeholders object |

