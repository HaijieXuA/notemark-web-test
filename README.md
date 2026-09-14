# NoteMark API Test

Experimental OneNote on the web task-pane add-in. Not a finished Markdown editor.

Uses Microsoft Office.js to read selected text and write HTML. Only operates on a page titled `NoteMark Web Test` and sample selections beginning with `API-` followed by digits. Supports simple bold, italic, strikethrough and highlight samples. No keyboard hooks or automatic conversion.

Host files with GitHub Pages from the main branch root. Upload `manifest.xml` through OneNote: Insert → Office Add-ins → Manage My Add-ins → Upload My Add-in.

The add-in code does not send note text to GitHub or an application backend. GitHub serves static assets and Microsoft provides Office.js and OneNote. Logs stay in the task pane memory. Office API conversion, undo and persistence validation is pending.
