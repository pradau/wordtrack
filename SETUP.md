# Setup Instructions

## Prerequisites
- Node.js v16+ (tested with v24.12.0) from https://nodejs.org/en/download  
- npm (tested with v11.6.2)
- Microsoft Word for Mac

Test that node and npm are installed by use these commands in the terminal: 
node -v
npm -v

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

## Step 4: Proxy Server Setup (Required for Claude API)

The WordTrack add-in requires a proxy server to communicate with the Claude API due to CORS (Cross-Origin Resource Sharing) restrictions in Office Add-ins. The proxy server must be running before the add-in can make API calls.

### Option 1: Use start.py (Recommended - Easiest)

The `start.py` script automatically starts both the proxy server and the add-in:

```bash
./start.py
```

Or:

```bash
python3 start.py
```

This will:
- Start the proxy server on port 3001 (in the background)
- Start the add-in debugging session
- Open Word with the add-in loaded
- Monitor Word and automatically save your work to `~/Downloads/Default.docx` when you close Word

**Keep the terminal window running** while using the add-in. Press `Ctrl+C` to stop both services.

**Note about Default.docx**: When you close Word, the script automatically finds the temp document and copies it to `~/Downloads/Default.docx`. You'll see a message showing where your work was saved.

### Option 2: Manual Setup (Two Terminals)

If you prefer more control or need to debug issues:

**Important:** Make sure you're in the project directory first:
```bash
cd /path/to/wordtrack
```

**Terminal 1 - Start the proxy server:**
```bash
cd /path/to/wordtrack
npm run proxy
```

You should see:
```
Claude API proxy server running on https://localhost:3001
Keep this running while using the WordTrack add-in
```

**Terminal 2 - Start the add-in:**
```bash
cd /path/to/wordtrack
npx office-addin-debugging start manifest.xml
```

**Note:** Both commands must be run from the project directory (where `package.json` and `manifest.xml` are located).

**Important Notes:**
- The proxy server must be running **before** you use any Claude API features in the add-in
- If the proxy server fails to start, you'll see "Add-in Error: sorry we can't load the add-in" in Word
- The proxy server uses HTTPS if certificates are available, otherwise falls back to HTTP
- If you see certificate warnings, run: `npx office-addin-dev-certs install`

### Verifying the Proxy Server is Running

To check if the proxy server is running:
```bash
lsof -i :3001
```

You should see a process listening on port 3001. If not, the proxy server isn't running.

## Step 5: Sideload the Add-in in Word

### Recommended Method: Use Office Add-in Developer Tools

The easiest and most reliable method is to use the command-line tools:

1. **Close Word completely** (if it's open)
2. **Navigate to the project directory:**
   ```bash
   cd /path/to/wordtrack
   ```
3. **Run this command:**
   ```bash
   npx office-addin-debugging start manifest.xml
   ```

**Important:** You must be in the project directory (where `package.json` and `manifest.xml` are located) for this command to work.

This will:
- Generate developer certificates for HTTPS localhost
- Start the webpack dev server automatically
- Sideload the add-in into Word
- Open Word automatically with the add-in loaded

**Important**: 
- You'll be prompted to accept a security certificate for localhost - this is safe for development
- Keep the terminal window running while using the add-in
- To stop, press `Ctrl+C` in the terminal
- **Saving your work**: When using `start.py`, your work is automatically saved to `~/Downloads/Default.docx` when you close Word. The script monitors Word and copies the temp document file when Word closes.

### Alternative: Manual Sideloading (if available)

If the above doesn't work, try manual sideloading:

1. **Navigate to the project directory:**
   ```bash
   cd /path/to/wordtrack
   ```

2. **Start the dev server manually:**
   ```bash
   npm run dev-server
   ```
   Keep this running in a separate terminal.

3. **In Word:**
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - Click **Upload My Add-in** (if available)
   - Select `manifest.xml` from your project directory

### If "ADMIN MANAGED" Blocks Sideloading

If you see "ADMIN MANAGED" and can't upload add-ins:

1. **Use the command-line method above** - it often bypasses UI restrictions
2. Contact your IT department to enable sideloading for development
3. Use a personal Office installation if available

## Step 6: Open the Task Pane

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

## Step 7: Enable Track Changes (Important!)

**Before using text insertion features**, enable Track Changes in Word:

1. Go to the **Review** tab in Word
2. Click **Track Changes** (or "Track Changes for Everyone" depending on your Word version)
3. Make sure it's turned ON

**Note**: Track Changes must be enabled manually - the add-in cannot enable it programmatically due to Office.js limitations.

## Step 8: Test the Add-in

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

### "ENOENT: no such file or directory, open '/Users/username/package.json'"

This error means you're running the command from the wrong directory:

1. **Navigate to the project directory:**
   ```bash
   cd /path/to/wordtrack
   ```
   Replace `/path/to/wordtrack` with the actual path to your project (e.g., `/Users/pradau/Dropbox/ChildrensHospital/IT/scripts/wordtrack`).

2. **Verify you're in the right directory:**
   ```bash
   ls package.json manifest.xml
   ```
   Both files should be listed. If not, you're in the wrong directory.

3. **Then run the command again:**
   ```bash
   npx office-addin-debugging start manifest.xml
   ```

### "Add-in Error: sorry we can't load the add-in" - Network Connectivity Error

This error can mean either the dev server (port 3000) or proxy server (port 3001) isn't running or accessible.

**First, check if the dev server is running (required for add-in to load):**

1. **Check if dev server is running on port 3000:**
   ```bash
   lsof -i :3000
   ```
   If nothing is shown, the dev server isn't running. This is the most common cause of this error.

2. **If dev server isn't running:**
   - The `office-addin-debugging` command should start it automatically, but sometimes it doesn't
   - **Solution: Start dev server manually in a separate terminal:**
     ```bash
     cd /path/to/wordtrack
     npm run dev-server
     ```
     Keep this terminal open. You should see webpack compilation output.
   - Verify certificates are installed: `npx office-addin-dev-certs install`
   - Once dev server is running, try accessing: `https://localhost:3000/taskpane.html` in Safari
   - **If prompted, accept the security certificate** (this is safe for localhost development)
   - After the dev server is running, try loading the add-in in Word again

**Then, check if proxy server is running (required for Claude API calls):**

3. **Check if proxy server is running on port 3001:**
   ```bash
   lsof -i :3001
   ```
   If nothing is shown, the proxy server isn't running.

4. **Start the proxy server (if not running):**
   ```bash
   npm run proxy
   ```
   Or use `./start.py` which starts both proxy and add-in.

5. **Check for port conflicts:**
   If port 3001 is already in use, you'll see an error. Find what's using it:
   ```bash
   lsof -i :3001
   ```
   Kill the process or change the port in `proxy-server.js` and `src/taskpane/taskpane.ts`.

6. **Check proxy server output:**
   When running `npm run proxy`, you should see:
   - `Claude API proxy server running on https://localhost:3001` (or `http://localhost:3001`)
   - If you see errors, check Node.js version (needs v16+)
   - Check that certificates exist if using HTTPS: `ls ~/.office-addin-dev-certs/`

7. **On older Macs:**
   - Ensure Node.js v16+ is installed: `node -v`
   - Try running services in foreground to see errors:
     - Terminal 1: `npm run proxy` (don't use `&`)
     - Terminal 2: `npm run dev-server` (don't use `&`)
   - Check firewall settings that might block localhost connections
   - Verify certificates are installed: `ls ~/.office-addin-dev-certs/`
   - If certificates are missing, run: `npx office-addin-dev-certs install`
   - If needed, try accessing `https://localhost:3000/taskpane.html` in Safari to accept the certificate (may not be required on all systems)

### "Invalid options object. Dev Server has been initialized using an options object that does not match the API schema" - Webpack Configuration Error

If you see this error when running `npm run dev-server`:

```
[webpack-cli] Invalid options object. Dev Server has been initialized using an options object that does not match the API schema.
 - options has an unknown property 'https'.
```

**This has been fixed in the current version of webpack.config.js**, but if you're seeing this error:

1. **Make sure you have the latest code** - the webpack config has been updated to use the correct API for webpack-dev-server 5
2. **If the error persists**, check your `webpack.config.js` - it should use `server: { type: 'https', ... }` instead of `https: true`
3. **Update dependencies** if needed:
   ```bash
   npm install
   ```

### Task pane shows errors
- Right-click in the task pane area (if possible) and select **Inspect** to open developer tools
- Check the Console tab for error messages
- Verify the dev server is accessible at https://localhost:3000
- Verify the proxy server is accessible at https://localhost:3001 (or http://localhost:3001)
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

