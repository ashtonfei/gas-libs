## Members

<dl>
<dt><a href="#ALERT_TYPE">ALERT_TYPE</a></dt>
<dd></dd>
</dl>

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

<a name="ALERT_TYPE"></a>

## ALERT\_TYPE
**Kind**: global variable  
**Properties**

| Name | Type | Default | Description |
| --- | --- | --- | --- |
| ALERT_TYPE.SUCCESS | <code>string</code> | <code>&quot;Success&quot;</code> | Success type |

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
**Exampe**: <caption>Create a UI objet for app SheetPro</caption>
const ui = new UI("SheetPro")  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| [appName] | <code>string</code> | <code>&quot;SheetPro&quot;</code> | The name of your application |


* [UI([appName])](#UI) ⇒ [<code>UI</code>](#UI)
    * [.alert(message, [title])](#UI+alert) ⇒ <code>void</code>
    * [.alertWithType(message, type)](#UI+alertWithType)

<a name="UI+alert"></a>

### uI.alert(message, [title]) ⇒ <code>void</code>
Show an alert message

**Kind**: instance method of [<code>UI</code>](#UI)  

| Param | Type | Default | Description |
| --- | --- | --- | --- |
| message | <code>string</code> |  | The alert message |
| [title] | <code>title</code> | <code>this.appName</code> | The alert title |

<a name="UI+alertWithType"></a>

### uI.alertWithType(message, type)
Show an alert message with a message type in the alert title

**Kind**: instance method of [<code>UI</code>](#UI)  

| Param | Type | Description |
| --- | --- | --- |
| message | <code>string</code> | The alert message |
| type | <code>UI.ALERT\_TYPE</code> | The type of alert |

