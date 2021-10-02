## Functions

<dl>
<dt><a href="#getSelectedRanges">getSelectedRanges()</a> ⇒ <code>Array.&lt;SpreadsheetApp.Range&gt;</code></dt>
<dd><p>Get selected ranges</p>
</dd>
<dt><a href="#getSelectedRows">getSelectedRows()</a> ⇒ <code>Array.&lt;SpreadsheetApp.Range&gt;</code></dt>
<dd><p>Get selected rows</p>
</dd>
<dt><a href="#getSelectedColumns">getSelectedColumns()</a> ⇒ <code>Array.&lt;SpreadsheetApp.Range&gt;</code></dt>
<dd><p>Get selected columns</p>
</dd>
<dt><a href="#UI">UI([appName])</a> ⇒ <code><a href="#UI">UI</a></code></dt>
<dd><p>Create a new UI object</p>
</dd>
</dl>

<a name="getSelectedRanges"></a>

## getSelectedRanges() ⇒ <code>Array.&lt;SpreadsheetApp.Range&gt;</code>
Get selected ranges

**Kind**: global function  
**Returns**: <code>Array.&lt;SpreadsheetApp.Range&gt;</code> - An array of selected ranges  
<a name="getSelectedRows"></a>

## getSelectedRows() ⇒ <code>Array.&lt;SpreadsheetApp.Range&gt;</code>
Get selected rows

**Kind**: global function  
**Returns**: <code>Array.&lt;SpreadsheetApp.Range&gt;</code> - An array of selected rows  
<a name="getSelectedColumns"></a>

## getSelectedColumns() ⇒ <code>Array.&lt;SpreadsheetApp.Range&gt;</code>
Get selected columns

**Kind**: global function  
**Returns**: <code>Array.&lt;SpreadsheetApp.Range&gt;</code> - An array of selected columns  
<a name="UI"></a>

## UI([appName]) ⇒ [<code>UI</code>](#UI)
Create a new UI object

**Kind**: global function  
**Returns**: [<code>UI</code>](#UI) - The UI object  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| [appName] | <code>string</code> | <code>&quot;SheetPro&quot;</code> | The name of your application |


* [UI([appName])](#UI) ⇒ [<code>UI</code>](#UI)
    * [.alert(message, [title])](#UI+alert) ⇒ <code>void</code>
    * [.info(message, [title])](#UI+info) ⇒ <code>void</code>
    * [.error(message, [title])](#UI+error) ⇒ <code>void</code>
    * [.success(message, [title])](#UI+success) ⇒ <code>void</code>

<a name="UI+alert"></a>

### uI.alert(message, [title]) ⇒ <code>void</code>
**Kind**: instance method of [<code>UI</code>](#UI)  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| message | <code>string</code> |  | The alert message |
| [title] | <code>title</code> | <code>this.appName</code> | The alert title |

<a name="UI+info"></a>

### uI.info(message, [title]) ⇒ <code>void</code>
**Kind**: instance method of [<code>UI</code>](#UI)  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| message | <code>string</code> |  | The info message |
| [title] | <code>title</code> | <code>this.appName</code> | The info title |

<a name="UI+error"></a>

### uI.error(message, [title]) ⇒ <code>void</code>
**Kind**: instance method of [<code>UI</code>](#UI)  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| message | <code>string</code> |  | The error message |
| [title] | <code>title</code> | <code>this.appName</code> | The error title |

<a name="UI+success"></a>

### uI.success(message, [title]) ⇒ <code>void</code>
**Kind**: instance method of [<code>UI</code>](#UI)  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| message | <code>string</code> |  | The success message |
| [title] | <code>title</code> | <code>this.appName</code> | The success title |

