# Development Guidelines (shrimp-rules.md)

This standards document is designed exclusively for AI Agent operational use in the School Device Management System (IThings) codebase.

## 1. Project Standards

### Single File Constraint
- **Rule**: All Google Apps Script logic, CSS, and HTML template rendering MUST reside within a single file: `Code.gs` in the root.
- **Rule**: DO NOT create separate `.html` or `.js` files. HTML templates must be returned as string templates inside `createHtml_()` and `createErrorHtml_()`.

### Styling Rules
- **Rule**: MUST use pure Vanilla CSS with CSS variables defined in `:root`.
- **Rule**: DO NOT load external CSS libraries or frameworks (such as Tailwind CSS or Bootstrap) via CDN.
- **Rule**: All UI designs MUST target a mobile-first responsive layout (maximum container width: `480px`).

## 2. Code and Schema Standards

### Dynamic Column Search
- **Rule**: MUST NOT hardcode column indices (e.g. `col 1`, `col 2`).
- **Rule**: MUST use `getHeaderMapping_(sheet)` to resolve column indices dynamically by matching column header names (e.g., `관리번호`, `설치장소`, `종류`).

### Security & PC Sensitive Data
- **Rule**: `IP`, `비밀번호(1차)`, `비밀번호(2차)` columns are sensitive.
- **Rule**: These fields MUST only be fetched and editable if:
  1. The device `종류` is exactly `"PC"` (case-insensitive).
  2. The `[설정]` sheet has `PC민감정보수정허용` set to `"Y"` or `"TRUE"`.
- **Rule**: For all other cases, these columns MUST be filtered out from the `getDeviceById_` payload.
- **Rule**: In `updateDevice_`, modifications to these columns MUST be ignored if the conditions are not met.
- **Rule**: Passwords in browser summary screens MUST be masked using `••••••`.
- **Rule**: In `createLabelSheet()`, passwords MUST be completely excluded.

## 3. Database Operations

### Concurrency Lock
- **Rule**: All spreadsheet updates in `updateDevice_` MUST wrap write operations inside `LockService.getScriptLock()` with a `10000` ms wait time.
- **Rule**: The lock MUST be released in a `finally` block.

## 4. Key File Interaction

- **Rule**: When modifying the database schema, update `setupSheets()` inside `Code.gs`, `docs/prd.md`, and `docs/roadmap.md` simultaneously to maintain alignment.
