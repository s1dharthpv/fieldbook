// Copyright (c) 2026, Sidharth PV and contributors
// For license information, please see license.txt

frappe.pages["fieldbook"].on_page_load = function (wrapper) {
	const API = "fieldbook.fieldbook.exporter.";
	const escape = frappe.utils.escape_html;

	const page = frappe.ui.make_app_page({
		parent: wrapper,
		title: __("Fieldbook"),
		single_column: true,
	});

	const $body = $(`
		<div style="max-width: 560px; margin: var(--margin-2xl) auto;">
			<p class="text-muted">
				${__("Download an Excel spec sheet for any DocType.")}
			</p>
			<div class="doctype-picker"></div>
			<div class="standard-fields"></div>
			<div class="customisation-only"></div>
			<div class="more-doctypes"></div>
			<div class="schema-preview" style="margin-top: var(--margin-lg);"></div>
		</div>
	`).appendTo(page.main);
	const $preview = $body.find(".schema-preview");

	// ---- turning a failed response into a sentence a person can act on ----------------------

	const stripHtml = (html) => $("<div>").html(html).text().trim();

	function serverMessages(data) {
		try {
			return JSON.parse((data || {})._server_messages || "[]")
				.map((raw) => {
					try {
						return JSON.parse(raw).message;
					} catch (e) {
						return raw;
					}
				})
				.map(stripHtml)
				.filter(Boolean);
		} catch (e) {
			return [];
		}
	}

	function errorText(data, status) {
		const messages = serverMessages(data);
		if (messages.length) return messages.join(" ");
		if (status === 403)
			return __("You do not have permission to export schemas. Ask a System Manager.");
		if (status === 404) return __("That DocType could not be found.");
		return __("The export failed. Check the Error Log for details.");
	}

	// ---- inputs -----------------------------------------------------------------------------

	const picker = frappe.ui.form.make_control({
		parent: $body.find(".doctype-picker"),
		df: {
			fieldname: "doctype",
			fieldtype: "Link",
			options: "DocType",
			label: __("DocType"),
			reqd: 1,
			change: () => refreshPreview(),
		},
		render_input: true,
	});

	const standard = frappe.ui.form.make_control({
		parent: $body.find(".standard-fields"),
		df: {
			fieldname: "include_standard",
			fieldtype: "Check",
			label: __("Include standard fields"),
			description: __("Adds owner, creation, modified, docstatus and idx."),
			change: () => refreshPreview(),
		},
		render_input: true,
	});

	// what is in the box right now; the control's own value can lag behind typed text
	const customised = frappe.ui.form.make_control({
		parent: $body.find(".customisation-only"),
		df: {
			fieldname: "customisations_only",
			fieldtype: "Check",
			label: __("Customisations only"),
			description: __("Only fields added or changed by Custom Fields and Property Setters."),
			change: () => refreshPreview(),
		},
		render_input: true,
	});

	const more = frappe.ui.form.make_control({
		parent: $body.find(".more-doctypes"),
		df: {
			fieldname: "more_doctypes",
			fieldtype: "Small Text",
			label: __("Also export these DocTypes"),
			description: __("One per line, up to 50. Downloads a zip."),
			change: () => refreshPreview(),
		},
		render_input: true,
	});

	more.$input.css({ height: "4.5rem", minHeight: "4.5rem" }); // optional box: a few lines are plenty

	const chosen = () => (picker.$input.val() || "").trim();
	const standardFlag = () => (standard.get_value() ? 1 : 0);
	const customisedFlag = () => (customised.get_value() ? 1 : 0);
	const extraNames = () =>
		(more.get_value() || "")
			.split(/[\n,]/)
			.map((name) => name.trim())
			.filter(Boolean);

	// ---- preview ----------------------------------------------------------------------------

	function renderPreview(info) {
		const traits = [info.kind];
		if (info.custom) traits.push(__("custom"));
		if (info.submittable) traits.push(__("submittable"));
		const heading = `<div><strong>${escape(info.doctype)}</strong>
			<span class="text-muted">${escape(traits.join(", "))} &middot; ${escape(
			info.module || ""
		)}</span></div>`;

		if (info.empty) {
			const reason = info.customisations_only
				? __(
						"This DocType has no customisations, so the workbook will contain only the header row."
				  )
				: __(
						"This DocType has no fields to export, so the workbook will contain only the header row."
				  );
			return `${heading}<p class="text-muted" style="margin: var(--margin-sm) 0 0;">${reason}</p>`;
		}

		const required = [__("{0} required", [info.required])];
		if (info.conditional) required.push(__("{0} conditional", [info.conditional]));
		let line = __("{0} fields", [info.fields]) + ` (${required.join(", ")})`;
		if (info.custom_fields) line += `, ${__("{0} custom", [info.custom_fields])}`;
		line += ` &middot; ${__("{0} child tables", [info.tables.length])} &middot; ${__(
			"{0} rows in the sheet",
			[info.rows]
		)}`;

		return `${heading}<div style="margin-top: var(--margin-xs);">${line}</div>`;
	}

	function extrasNote() {
		const count = extraNames().length;
		return count
			? `<div class="text-muted" style="margin-top: var(--margin-sm);">${__(
					"+{0} more DocType(s) in a zip",
					[count]
			  )}</div>`
			: "";
	}

	let previewToken = 0; // a slow answer for an earlier choice must not overwrite a newer one
	let previewKey = null; // what the card currently describes

	function refreshPreview() {
		const doctype = chosen();
		const key = `${doctype}|${standardFlag()}|${customisedFlag()}|${extraNames().join("+")}`;
		if (key === previewKey) return;
		previewKey = key;
		const mine = ++previewToken;
		if (!doctype) {
			$preview.empty();
			return;
		}
		$preview.html(`<div class="text-muted">${__("Checking...")}</div>`);
		frappe.call({
			method: API + "get_preview",
			args: {
				doctype,
				include_standard: standardFlag(),
				customisations_only: customisedFlag(),
			},
			silent: true,
			callback: (r) => {
				if (mine === previewToken) {
					$preview.html(
						`<div class="frappe-card" style="padding: var(--padding-md);">${renderPreview(
							r.message
						)}</div>${extrasNote()}`
					);
				}
			},
			error: (xhr) => {
				if (mine === previewToken) {
					$preview.html(
						`<div class="text-danger">${escape(
							errorText(xhr.responseJSON, xhr.status)
						)}</div>`
					);
				}
			},
		});
	}

	// ---- download ---------------------------------------------------------------------------

	async function download() {
		const names = [chosen(), ...extraNames()].filter(Boolean);
		if (!names.length) {
			frappe.msgprint(__("Select a DocType first"));
			return;
		}
		const doctype = names[0];
		const options = `include_standard=${standardFlag()}&customisations_only=${customisedFlag()}`;
		const path =
			names.length > 1
				? `download_pack?doctypes=${encodeURIComponent(JSON.stringify(names))}&${options}`
				: `download_workbook?doctype=${encodeURIComponent(doctype)}&${options}`;
		const url = frappe.urllib.get_full_url(`/api/method/${API}${path}`);
		const $button = page.btn_primary;
		$button.prop("disabled", true);
		try {
			// fetched here, not opened in a new tab, so a failure shows as a message and not an error page
			const response = await fetch(url, {
				credentials: "same-origin",
				headers: { "X-Frappe-CSRF-Token": frappe.csrf_token },
			});
			if (!response.ok) {
				let data = null;
				try {
					data = await response.json();
				} catch (e) {
					// not JSON: fall back to the status text below
				}
				frappe.msgprint({
					title: __("Export failed"),
					indicator: "red",
					message: escape(errorText(data, response.status)),
				});
				return;
			}

			const disposition = response.headers.get("Content-Disposition") || "";
			const match = disposition.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
			const filename = match
				? decodeURIComponent(match[1])
				: names.length > 1
				? `Fieldbook_${names.length}_DocTypes.zip`
				: `${doctype.replace(/\W+/g, "_")}_Global_Schema.xlsx`;

			const blobUrl = URL.createObjectURL(await response.blob());
			const link = document.createElement("a");
			link.href = blobUrl;
			link.download = filename;
			document.body.appendChild(link);
			link.click();
			link.remove();
			setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
			frappe.show_alert({ message: __("Downloaded {0}", [filename]), indicator: "green" });
		} catch (e) {
			frappe.msgprint({
				title: __("Export failed"),
				indicator: "red",
				message: __("Could not reach the server. Check your connection and try again."),
			});
		} finally {
			$button.prop("disabled", false);
		}
	}

	// typing makes the card stale at once; it is refreshed when you pick from the list or leave the box
	picker.$input.on("input", () => {
		previewToken++;
		previewKey = null;
		$preview.empty();
	});
	// Frappe's link field validates asynchronously on blur and may briefly empty the box, so look again
	// once it has settled
	picker.$input.on("blur", () => {
		setTimeout(refreshPreview, 200);
		setTimeout(refreshPreview, 1000);
	});

	page.set_primary_action(__("Export Workbook"), download);
	picker.$input.on("keydown", (e) => {
		if (e.key === "Enter") setTimeout(download, 300); // let the link field settle its value first
	});
};
