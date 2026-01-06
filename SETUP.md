# Setup Instructions

## Prerequisites
- Node.js v16+ (tested with v24.12.0)
- npm (tested with v11.6.2)
- Microsoft Word for Mac

## Step 1: Fix npm Cache Permissions (if needed)

If you encounter permission errors when running `npm install`, fix the npm cache:

```bash
sudo chown -R 501:20 "/Users/$(whoami)/.npm"
```

Then try installing again.

## Step 2: Install Dependencies

Run the following command in the project directory:

```bash
cd /path/to/wordtrack
npm install
```

This will install all required packages including:
- @microsoft/office-js (Office.js runtime)
- @types/office-js (TypeScript types)
- TypeScript compiler
- Webpack bundler
- Office Add-in Developer Tools

**Note**: You may see a deprecation warning for @microsoft/office-js. This is expected - Office.js is loaded from CDN, but the package is still useful for TypeScript types.

## Step 3: Build the Project

Build the add-in for development:

```bash
npm run build:dev
```

This compiles TypeScript to JavaScript and bundles everything into the `dist/` folder.

## Step 4: Sideload the Add-in in Word

### Recommended Method: Use Office Add-in Developer Tools

The easiest and most reliable method is to use the command-line tools:

1. **Close Word completely** (if it's open)
2. Run this command in the project directory:

```bash
npx office-addin-debugging start manifest.xml
```

This will:
- Generate developer certificates for HTTPS localhost
- Start the webpack dev server automatically
- Sideload the add-in into Word
- Open Word automatically with the add-in loaded

**Important**: 
- You'll be prompted to accept a security certificate for localhost - this is safe for development
- Keep the terminal window running while using the add-in
- To stop, press `Ctrl+C` in the terminal

### Alternative: Manual Sideloading (if available)

If the above doesn't work, try manual sideloading:

1. Start the dev server manually:
   ```bash
   npm run dev-server
   ```
   Keep this running in a separate terminal.

2. In Word:
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - Click **Upload My Add-in** (if available)
   - Select `manifest.xml` from your project directory

### If "ADMIN MANAGED" Blocks Sideloading

If you see "ADMIN MANAGED" and can't upload add-ins:

1. **Use the command-line method above** - it often bypasses UI restrictions
2. Contact your IT department to enable sideloading for development
3. Use a personal Office installation if available

## Step 5: Open the Task Pane

After the add-in is loaded, open the task pane:

**Method 1: Via Insert Menu (Most Reliable)**
1. In Word, go to **Insert** > **Add-ins** > **My Add-ins**
2. Look for **WordTrack** in the list
3. Click on **WordTrack** to open the task pane

**Method 2: Via Home Tab Ribbon**
1. Go to the **Home** tab in Word
2. Look for the **WordTrack** group in the ribbon
3. Click the **Show Taskpane** button

The task pane should appear on the right side of Word.

## Step 6: Enable Track Changes (Important!)

**Before using text insertion features**, enable Track Changes in Word:

1. Go to the **Review** tab in Word
2. Click **Track Changes** (or "Track Changes for Everyone" depending on your Word version)
3. Make sure it's turned ON

**Note**: Track Changes must be enabled manually - the add-in cannot enable it programmatically due to Office.js limitations.

## Step 7: Test the Add-in

1. **Test Text Extraction:**
   - Type some text in your Word document
   - Select the text with your mouse
   - In the task pane, click **Get Selected Text**
   - You should see your selected text displayed

2. **Test Text Transformation:**
   - After extracting text, click **Capitalize All Words & Insert**
   - The text should be replaced with capitalized version
   - If Track Changes is enabled, you'll see the changes marked

## Troubleshooting

### npm Install Permission Errors

If you see `EACCES` or permission errors:

```bash
sudo chown -R 501:20 "/Users/$(whoami)/.npm"
npm cache clean --force
npm install
```

### Add-in doesn't appear in Word
- Make sure you used `npx office-addin-debugging start manifest.xml` (this starts the server automatically)
- Or manually start the dev server: `npm run dev-server`
- Try restarting Word completely
- Check that Word opened automatically after running the debugging command

### Task pane shows errors
- Right-click in the task pane area (if possible) and select **Inspect** to open developer tools
- Check the Console tab for error messages
- Verify the dev server is accessible at https://localhost:3000
- Make sure you accepted the security certificate for localhost

### "ADMIN MANAGED" - Cannot upload add-ins
This means your Office installation is managed by your organization's IT policies. Solutions:

1. **Use command-line sideloading** (recommended):
   ```bash
   npx office-addin-debugging start manifest.xml
   ```
   This often bypasses UI restrictions.

2. **Contact IT Support**: Request permission to sideload add-ins for development purposes

3. **Use Personal Office**: If you have a personal Microsoft 365 account, use that installation

### Text insertion doesn't show Track Changes
- Make sure Track Changes is enabled **before** using the add-in
- Go to **Review** tab > **Track Changes** and enable it
- Close and reopen the task pane after enabling Track Changes
- The add-in cannot programmatically enable Track Changes due to Office.js limitations

### Button doesn't work or shows errors
- Refresh the task pane: close it and reopen via Insert > Add-ins > My Add-ins > WordTrack
- Check the browser console for specific error messages
- Make sure you've selected text in the document before clicking "Get Selected Text"

## Development Workflow

1. Make code changes to files in `src/`
2. Run `npm run build:dev` to rebuild
3. Refresh the task pane in Word (close and reopen)
4. Or restart the debugging session: `npx office-addin-debugging start manifest.xml`

## Next Steps

Once setup is complete and you can extract and insert text, you're ready for Phase 3: Claude API Integration.

