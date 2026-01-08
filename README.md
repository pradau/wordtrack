# WordTrack - AI-Powered Word Add-in with Track Changes

Word add-in that integrates Claude API to allow custom AI-powered edits with Track Changes in Microsoft Word.

## Project Status

**Current Phase: Phase 5 in progress**

- Phase 3: COMPLETE - Core functionality (text extraction, Claude API, insertion with Track Changes)
- Phase 4A: COMPLETE - User-friendly startup script with temp document hiding
- Phase 5: IN PROGRESS - Track Changes + Core Polish (formatting preservation implemented)

For detailed development plans and current status, see [DEV_PLAN.md](DEV_PLAN.md) and [DEV_PLAN_PHASE_5.md](DEV_PLAN_PHASE_5.md).

## Development

### Quick Start
The easiest way to start WordTrack:
```bash
./start.sh
```

This script automatically:
- Starts the proxy server
- Starts the add-in debugging session
- Opens Default.docx (if available)
- Hides the temporary document window

### Manual Setup (Alternative)
1. Install dependencies: `npm install`
2. Build: `npm run build:dev`
3. Start proxy server: `npm run proxy` (in one terminal)
4. Start dev server: `npx office-addin-debugging start manifest.xml` (in another terminal)

### Requirements
- Node.js v16+
- Microsoft Word for Mac
- Anthropic Claude API key with credits
- Track Changes must be manually enabled in Word before using the add-in
- Proxy server must be running for Claude API calls to work

## Project Structure

```
wordtrack/
├── manifest.xml          # Office Add-in manifest
├── package.json          # Dependencies and scripts
├── tsconfig.json        # TypeScript configuration
├── webpack.config.js    # Webpack bundler config
├── src/
│   ├── taskpane/       # Task pane UI and logic
│   └── commands/       # Command handlers
└── dist/               # Built files (gitignored)
```

## Notes

- Track Changes must be enabled manually in Word (Review tab) before using text insertion features
- The add-in works best when Track Changes is enabled before opening the task pane

