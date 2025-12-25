// Initialize menu buttons
const onOpen = () => {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	ss?.addMenu("Export JSON", [
		{ name: "Export JSON for this sheet", functionName: "export_active_sheet" },
		{ name: "Export JSON for all sheets", functionName: "export_all_sheets" },
	]);
}

// Header functions

const normalize_header = (header) => header.toLowerCase().replace(/\W+/g, "_");

const normalize_headers = (headers) => {
	const h = headers.map(key => normalize_header(key));
	return h.toSpliced(h.findLastIndex((c) => c.length > 0) + 1);
}

// Data functions

const array_of_arrays = (n) => Array.from(Array(n), () => Array());

const transpose_array = (data) => {
	if (data.length === 0 || data[0].length === 0) return [];

	const arr = array_of_arrays(data[0].length);

	for (const i = 0; i < data.length; ++i)
		for (const j = 0; j < data[i].length; ++j)
			arr[j][i] = data[i][j];

	return arr;
}

const is_cell_not_empty = (cell) => typeof(cell) !== "string" || cell !== "";

const is_object_empty = (o) => {
	if (!o) return false;
	for (let prop in o) if (o.hasOwnProperty(prop)) return false;
	return true;
}

const get_objects = (data, headers) => {
	const o = data.map((row, i) => Object.fromEntries(row.map((cell, j) => [headers[j] || j, cell]).filter(([,cell]) => is_cell_not_empty(cell))));
	return o.toSpliced(o.findLastIndex((r) => !is_object_empty(r)) + 1);
}

const get_rows_data = (sheet) => {
	const frozen_rows = sheet.getFrozenRows() || 1,
		headers = normalize_headers(sheet.getRange(1, 1, frozen_rows, sheet.getMaxColumns() || 1000)?.getValues()?.[0] ?? []),
		data = sheet.getRange(frozen_rows + 1, 1, sheet.getMaxRows(), sheet.getMaxColumns())?.getValues() ?? [[]];

	return get_objects(data, headers);
}

// Output functions

const format_JSON = (o) => JSON.stringify(o, null, 4);

const display_text = (text) => {
	const output = HtmlService.createHtmlOutput(
		`<style>body { display: grid; margin: 0; padding: 0; max-width: 900px; max-height: 600px; overflow: auto; } pre { white-space: pre-wrap; word-wrap: anywhere; }</style>
		<pre><code id="code" tabindex="0">${text}</code></pre>
		<script>document.querySelector("#code")?.addEventListener("click", (e) => { window.getSelection().selectAllChildren(e.currentTarget); navigator?.clipboard?.writeText?.(e.currentTarget.textContent).then(() => {}) });</script>`);
	output.setWidth(900);
	output.setHeight(600);

	SpreadsheetApp.getUi().showModalDialog(output, "Exported JSON");
}

// Main functions

const export_all_sheets = (e) => {
	const ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheets = ss.getSheets(),
		json = format_JSON(Object.fromEntries(sheets.map((sheet) => [sheet.getName(), get_rows_data(sheet)])));

	display_text(json);
}

const export_active_sheet = (e) => {
	const ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = ss.getActiveSheet(),
		json = format_JSON(get_rows_data(sheet));

	display_text(json);
}

