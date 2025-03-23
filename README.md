# Upsertable
*We have Linked tables at home.*

## Why?
Because when I started writing this script I was under the impression that [Google Docs](https://en.wikipedia.org/wiki/Google_Docs) documents that contain lots of [Linked tables](https://support.google.com/docs/answer/7009814?hl=en) are difficult to maintain because of having to go through the whole document to update each Linked table manually, one by one, with no way of easily updating all outdated tables at once. Well, after writing this script, I found out there actually is a way to do precisely that (Tools > Linked objects > Update All). But the script has been written, so it shall be given an alternative reason to exist. Basically, it *is* Linked tables but slower but more customizable but uglier and it can filter rows by executing arbitrary unsanitized user code. Also because this is an Apps Script script, you could choose to execute it periodically or on opening a Google Docs document, instead of having to click some ugly god-ordained sequence of buttons like some unthinking mechanical chimp or macaque, even.

By the way, if you have a Google Docs document that contains lots of tables that may be modified in the future and you aren't using Linked tables, you should.

Also, I don't like Google. I only use Google Docs because of work.

## What?
This is an [Apps Script](https://developers.google.com/apps-script) script meant to be imported as a [library](https://developers.google.com/apps-script/guides/libraries) by the [Container-bound Script](https://developers.google.com/apps-script/guides/bound?hl=en) script of (bounded to) a Google Docs document.

Using [custom syntax](#usage) (don't even click on the link, it just redirects to a later section which you are going to get to eventually, just keep reading, trust me) to reference a [Google Sheets](https://en.wikipedia.org/wiki/Google_Sheets) sheet from within a Google Docs document, this script can then be executed on the Google Docs document to automatically insert or update (upsert) all Google Sheets sheets referenced in this way throughout the entire Google Docs document (including all nested tabs). This would be awesome had Google not already implemented it. But hey, this way you can also execute arbitrary unsanitized user code while you are at it.

If a reference to a Google Sheets sheet is found as a standalone paragraph in the Google Docs document, it is replaced by its corresponding up-to-date Google Spreadhseets sheet, converted into a Google Docs document table. Also, the reference itself is added at the top of the inserted table, so that future executions of the script can still find it and update the table as a result: when a reference to a Google Sheets sheet is found as part of a pre-existing table, the table is deleted and re-inserted (effectively updated) using the current up-to-date Google Spreadhseets sheet data.

## How?
### Setup
- Create the library as a [standalone Apps Script project](https://developers.google.com/apps-script/guides/projects#create-standalone), using the code in this repo as the Code.gs file of the library.
- Get the project ID (e.g. something like `1eQtGxTvJ34kcqlP9ZBzbqiNdswr0T4FsoB_3nWVSGvaZctpT7BKneA3N`) from the script's project's "Project Settings" tab.
- Go to your desired Google Docs document and [create a new Container-bound Script script](https://developers.google.com/apps-script/guides/projects#create-from-docs-sheets-slides) for it.
- Import the library using its project ID.
- Create a function that calls the library's `upsert` function, passing the current Google Docs document's [Document class](https://developers.google.com/apps-script/reference/document/document) instance as a parameter (`Upsertable.upsert(DocumentApp.getActiveDocument());`).
- [Trigger](https://developers.google.com/apps-script/guides/triggers) the execution of the Container-bound Script script in whichever way suits your needs.

For example, if you'd like to follow some ugly god-ordained sequence of buttons like some unthinking mechanical chimp or macaque, even, you could use the following snippet as your Container-bound Script script:

```javascript
function onOpen(e) {
  DocumentApp.getUi()
    .createMenu("Tables") // Creates a new "Tables" menu on the Google Docs document toolbar
    .addItem("Upsert", "main") // Adds a new "Upsert" menu item, which calls main when clicked
    .addToUi();
}

function main() {
  Upsertable.upsert(DocumentApp.getActiveDocument()); // Calls the library's upsert function
}
```

### Usage
Google Spreadsheets sheets can be referenced throughout the Google Docs document by using the following custom syntax:

`@<spreadsheet UID><query string>`

Where `<spreadsheet UID>` is the 44 characters long unique identifier of a Google Sheets sheet you have access to, and `<query string>` is optional and has a very similar structure to that of an [URL's query string](https://en.wikipedia.org/wiki/Query_string) (including a leading `?`), except parameters are separated by `%` instead of `&` (worry not, the reason for this subversion of implicit expectations will be revealed to the devoted documentation reader). All parameters are optional and can be provided in any order.

#### Query string parameters
##### `sheet`
The `sheet` parameter defines which sheet of the Google Sheets spreadsheet must be upserted as a table in the Google Docs document.
Its value must be a valid numeric unique identifier of a sheet in the Google Sheets spreadsheet (e.g. something like `451951589`).

If `sheet` is not provided, the first sheet of the Google Sheets spreadsheet will be used.

##### `range`
The `range` parameter defines the range of rows and columns of the Google Sheets sheet that must make up the table in the Google Docs document (before applying [filters](#filters)).
Its value must be a valid range written in [A1 notation](https://developers.google.com/workspace/sheets/api/guides/concepts#expandable-1), except the sheet name must be omitted (e.g. `A:D` o `A2:C10`). Additionally, the range must represent a rectangular and continuous region of the Google Sheets sheet.

If `range` is not provided, the smallest possible range that includes all non-empty cells of the Google Sheets sheet will be used.

##### `filters`
The `filters` parameter is by far the most fun of all parameters, as it's the only one that allows the execution of arbitrary unsanitized user code. Its value consists of a series of column:filter pairs. Only rows that pass all filters (and are part of the defined [range](#range)) will be shown in the Google Docs document, except for the first row (which is assumed to contain column names and is always shown).

The column:filter pairs must be separated by `$` (the devoted domentation reader may suggest something like `|` would be a better separator and would be correct, the reason for this subversion of strawmanned explicit expectations will soon be revealed to the most devoted of documentation readers). There can be at most one filter per column.

In each column:filter pair, column and filter must be separated by `:`. The column can be specified using either its label (e.g. `A`, `D`, `BC`) or name (in the latter case, it is assumed that all column names are unique and can be found in the first row of the Google Sheets sheet).

Each filter must be a valid single-line JavaScript expression that is a unary function of `x` and evaluates to `true` or `false` (or any other value that can be automatically coerced into a boolean). In these expressions, the `x` variable represents the table cell value being tested.

As an additional restriction, neither `%` or `$` can be used as part of column names or filters. Only for the true devoted: here lies the reason for `&` and `|` not being used as separators - so that the `&&` and `||` operators may be used inside filters without breaking parameter parsing (or, rather, so that I can keep avoiding writing an actual parser).

**Warning:** When using string literals inside filters, beware the usage of `“` and `‘` for they are invalid tokens and false icons. For some reason that surely is beyond mortal understanding, when typing `"` or `'` Google Docs will instead use `“` or `‘` respectively, so that ultimately it is necessary to copy-paste the valid character from a third-party text editor that actually respects you as a human being.

##### `fontFamily`
The `fontFamily` parameter... I mean come on, do I really need to say it?

If `fontFamily` is not provided, `Roboto Slab` will be used. Because that's the one we use at work.

##### `fontSize`
`9` is the default value. Oh, is that too small for you? Congratulations, you just outed yourself as being 30+ years old.

## Example
Here's an example [Google Sheets sheet](https://docs.google.com/spreadsheets/d/10kkei0NCiVISweVuV2N6i1iL2CIMXbGaBHSRMkbrdhc?gid=1608842880) and [Google Docs document](https://docs.google.com/document/d/1jm-UC0v4VvNG8Bjn3hKcb9eT5Oeg8XAglVvlH9fFe18) that has already been processed using this script. Before, the document consisted of this single line of text: `@10kkei0NCiVISweVuV2N6i1iL2CIMXbGaBHSRMkbrdhc?sheet=1608842880%range=A:B%filters=Class Level: parseInt(x[0]) >= 3$Major: x !== "Art"`

You do the math!
