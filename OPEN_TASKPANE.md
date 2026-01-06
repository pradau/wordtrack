# How to Open the WordTrack Task Pane

## Method 1: Through Insert Menu (Easiest)

1. In Word, go to **Insert** menu (top menu bar)
2. Click **Add-ins**
3. Click **My Add-ins**
4. Look for **WordTrack** in the list
5. Click on **WordTrack** - this should open the task pane on the right side

## Method 2: Check Home Tab Ribbon

1. Go to the **Home** tab in Word
2. Look carefully at the ribbon for a group called **WordTrack**
3. If you see it, click the **Show Taskpane** button

## Method 3: If Add-in Doesn't Appear

If WordTrack doesn't appear in "My Add-ins":

1. Make sure the dev server is running:
   ```bash
   npm run dev-server
   ```
   OR
   ```bash
   npx office-addin-debugging start manifest.xml
   ```

2. Restart Word completely (quit and reopen)

3. Try reloading the add-in:
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - If WordTrack appears but is grayed out, click on it anyway
   - Or remove it and re-sideload using the debugging command

## Method 4: Direct URL (Advanced)

If nothing else works, you can try opening the task pane directly by:
1. Making sure dev server is running at https://localhost:3000
2. The task pane URL is: https://localhost:3000/taskpane.html
3. But Word needs to load it through the add-in system, so this won't work directly

## Troubleshooting

- **"Add-in Error" message**: The dev server might not be running or the manifest has issues
- **Add-in not in list**: The sideloading might have failed - try running `npx office-addin-debugging start manifest.xml` again
- **Button not in ribbon**: Missing icons might prevent the button from showing - we'll fix this next

