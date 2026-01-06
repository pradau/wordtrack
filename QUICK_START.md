# Quick Start - Get WordTrack Working

## Step 1: Close Word Completely
- Quit Microsoft Word completely (Word > Quit Word, or Cmd+Q)
- Make sure no Word windows are open

## Step 2: Start the Add-in Server
Open Terminal and run:
```bash
cd /Users/pradau/Dropbox/ChildrensHospital/IT/scripts/wordtrack
npx office-addin-debugging start manifest.xml
```

**Wait for it to say "Debugging started" and Word should open automatically.**

## Step 3: Find the Add-in in Word

### Option A: Check the Home Tab Ribbon
1. Look at the **Home** tab in Word
2. Look for a button or group called **WordTrack** or **Show Taskpane**
3. Click it if you see it

### Option B: Use Insert Menu
1. Click **Insert** in the top menu
2. Click **Add-ins**
3. Click **My Add-ins**
4. Look for **WordTrack** in the list
5. Click on **WordTrack** to open the task pane

## Step 4: If You Still Don't See It

1. **Check the terminal** - Does it say "Debugging started" without errors?
2. **Check Word** - Did a new Word window open? Look for any error messages
3. **Try this**: In Word, go to **File** > **Options** > **Add-ins**
   - At the bottom, change the dropdown to **Disabled Items**
   - If WordTrack is there, select it and click **Enable**

## Still Stuck?

Run this command and share the output:
```bash
cd /Users/pradau/Dropbox/ChildrensHospital/IT/scripts/wordtrack
npx office-addin-debugging start manifest.xml
```

Look for any error messages in red and share them.

