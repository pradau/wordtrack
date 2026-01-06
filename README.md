# WordTrack - AI-Powered Word Add-in with Track Changes

Word add-in that integrates Claude API to allow custom AI-powered edits with Track Changes in Microsoft Word.

## Project Status

**Current Phase: Phase 3 Complete**

- ✅ Phase 1: Basic add-in setup with task pane
- ✅ Phase 2: Text extraction and insertion with Track Changes
- ✅ Phase 3: Claude API integration
- ⏳ Phase 4: Track Changes implementation refinement (optional)
- ⏳ Phase 5: UI polish & error handling (optional)

## Git History

- `phase-3-complete` tag: Current state - Claude API integration working
- `phase-2-complete` tag: Text extraction and insertion working
- Initial commit: Complete Phase 2 implementation

## Rolling Back

To roll back to Phase 2 state:
```bash
git checkout phase-2-complete
```

To see what changed since Phase 2:
```bash
git diff phase-2-complete HEAD
```

## Development

### Setup
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

