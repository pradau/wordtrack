# Phase 1 Setup Instructions

## Prerequisites
- Node.js v16+ (you have v24.12.0 - perfect!)
- npm (you have v11.6.2 - perfect!)
- Microsoft Word for Mac

## Step 1: Install Dependencies

Run the following command in the project directory:

```bash
npm install
```

This will install all required packages including:
- Office.js types and runtime
- TypeScript compiler
- Webpack bundler
- Development tools

## Step 2: Build the Project

Build the add-in for development:

```bash
npm run build:dev
```

This compiles TypeScript to JavaScript and bundles everything into the `dist/` folder.

## Step 3: Start Development Server

In a separate terminal, start the webpack dev server:

```bash
npm run dev-server
```

This starts a local HTTPS server at `https://localhost:3000` which serves your add-in files.

**Important**: Keep this server running while testing the add-in in Word.

## Step 4: Sideload the Add-in in Word

### Option A: Standard Sideloading (if available)

1. Open Microsoft Word for Mac
2. Go to **Insert** > **Add-ins** > **My Add-ins**
3. Click the **Upload My Add-in** button (folder icon at the bottom)
4. Navigate to your project directory and select `manifest.xml`
5. Click **Upload**

### Option B: If "ADMIN MANAGED" blocks sideloading

If you see "ADMIN MANAGED" and the Upload button is missing/disabled, try these alternatives:

**1. Contact Your IT Department:**
   - Request permission to sideload add-ins for development
   - Ask them to enable "My Custom Apps" permission in Exchange Admin Center
   - They may need to adjust Group Policy settings

**2. Use a Personal Office Installation:**
   - If you have a personal Microsoft 365 subscription, install Word separately
   - Personal installations typically allow sideloading without restrictions

**3. Try Developer Mode (if available):**
   - In Word, go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
   - Look for "Developer Mode" or "Enable add-in debugging"
   - Enable if available (may still be blocked by admin policies)

**4. Use Office Add-in Developer Tools (after installing dependencies):**
   ```bash
   npm install
   npx office-addin-debugging start manifest.xml
   ```
   This may bypass some UI restrictions by using command-line sideloading.

## Step 5: Open the Task Pane

1. In Word, go to the **Home** tab
2. Look for the **WordTrack** group in the ribbon
3. Click the **Show Taskpane** button
4. The task pane should appear on the right side

## Step 6: Test the Button

1. In the task pane, you should see a blue button labeled "Click Me"
2. Click the button
3. You should see an alert dialog saying "Hello World"

## Troubleshooting

### Add-in doesn't appear in Word
- Make sure the dev server is running (`npm run dev-server`)
- Check that you selected the `manifest.xml` file (not a folder)
- Try restarting Word

### Task pane shows errors
- Open the browser developer tools (right-click in task pane > Inspect)
- Check the Console tab for error messages
- Verify the dev server is accessible at https://localhost:3000

### Build errors
- Make sure all dependencies are installed: `npm install`
- Check that Node.js version is 16 or higher: `node --version`

### "ADMIN MANAGED" - Cannot upload add-ins
This means your Office installation is managed by your organization's IT policies. Solutions:

1. **Contact IT Support**: Request permission to sideload add-ins for development purposes
2. **Use Personal Office**: If you have a personal Microsoft 365 account, use that installation
3. **Command-line sideloading**: After `npm install`, try:
   ```bash
   npx office-addin-debugging start manifest.xml
   ```
4. **Check Office Settings**: Go to **File** > **Account** and verify if you can switch to a personal account

## Next Steps

Once Phase 1 is working (you can see the task pane and clicking the button shows "Hello World"), you're ready for Phase 2: Text Extraction.

