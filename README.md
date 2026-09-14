# NoteMark Web 0.2 — experimental

OneNote Web task pane using Office.js. Supports a selected paragraph with headings, bold, italic, strikethrough and highlight. Source records are scoped to notebook, page and paragraph in local browser storage. Restoration refuses changed text or formatting. A pending source backup is persisted before a write.

Upload manifest.xml from OneNote's Add-ins menu. The default writing scope is the page named `NoteMark Web Test`. Select a complete paragraph and use the conversion button. Source history is local; closing an incognito session or clearing browser data can remove it.

The keyboard/Enter bridge remains experimental and requires separate installation. Keep it disabled until host write, caret and undo behavior have been verified. API initialization and text reads have been verified; full write and persistence acceptance is still in progress. This is not a production release.

No analytics or remote note storage. Office.js is loaded from Microsoft's official CDN.
