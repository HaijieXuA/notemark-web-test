# NoteMark Web 0.2.1 — experimental

A personal OneNote Web Markdown add-in. Select a complete paragraph to render headings, bold, italic and strikethrough using Office.js. The paired Chrome extension adds Option+Command+Enter, Command+Comma source restoration, and opt-in experimental Enter conversion. Highlight requires the bridge and OneNote's yellow highlighter.

## Setup
1. Upload `manifest.xml` using OneNote Web → Insert → Office Add-ins → Manage My Add-ins → Upload My Add-in.
2. For keyboard support, load the unpacked `extension` folder in Chrome's extension developer mode. For an incognito OneNote session, explicitly allow that extension in incognito.
3. Keep the NoteMark pane open and enable its extension checkbox. Automatic Enter is a separate, initially unchecked experimental option.
4. By default writes are restricted to a page named `NoteMark Web Test`.

## Source records
Original Markdown is stored locally, scoped to notebook, page and paragraph; 500 recent successful records are retained. A source backup is persisted before writing. Restoration checks the text and the formatting that OneNote exposes. OneNote's HTML export can omit highlight, so it cannot detect every manual formatting change. Browser storage is not part of notebook sync. Clearing browser data or closing an incognito session can erase it.

## Current validation and limitations
Official host initialization, selected text reads and HTML replacement have been exercised in OneNote Web. One measured HTML write took about 79 ms (not end-to-end latency). Bold and italic survived refresh. Native webpage highlight also survived refresh. `<s>` and `<del>` import preserve strikethrough; inline CSS strikethrough was discarded in this host.

Keyboard integration, complete source round-trips, heading timing, real IME input and rapid typing are still being validated. Do not treat this as a production release. If typing during an asynchronous conversion, current selection can change; use only test notes while evaluating the experimental Enter mode. If the current paragraph cannot be uniquely located, the bridge stops rather than guessing. Wrapped lines and complex nested Markdown are not yet accepted.

No analytics or remote note storage. Office.js is loaded from Microsoft's CDN; GitHub Pages serves the application code.


## Recovery update
The restore button uses an exact saved source when valid, or generates Markdown from exported paragraph formatting when no record exists. A separate Generate button explicitly uses current formatting when a saved record is stale. Unsupported content is refused. Formatting omitted by OneNote HTML export, including some highlight, cannot be reconstructed. Font family, size and color are not represented in Markdown.

The pane now displays extension connection state. Enter is disabled while disconnected or when keyboard support is unchecked. Clearing source records requires confirmation. This update has local unit and mock DOM coverage; live acceptance is pending.
