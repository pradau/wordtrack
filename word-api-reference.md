# Word Add-In API Quick Reference

A concise guide to key Office.js Word API capabilities for your LLM-powered add-in.

**Main Documentation**: https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview

---

## Document Object

The top-level object representing the Word document.

### Track Changes
```javascript
// Enable/disable track changes
context.document.trackRevisions = true;

// Get current tracking mode
context.document.changeTrackingMode; // "Off" | "TrackAll" | "TrackMineOnly"

// Access revisions
const revisions = context.document.revisions;
```

### Body & Content
```javascript
// Main document body (excludes headers/footers)
const body = context.document.body;

// Get entire document range
const content = context.document.content;

// Get current selection
const selection = context.document.getSelection();
```

### Document Properties
```javascript
// Built-in properties
const props = context.document.properties;
props.load("title,author,creationDate");

// Custom properties
const customProps = context.document.customDocumentProperties;
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.document

---

## Range Object

Represents a contiguous area of content (text, tables, images, etc.)

### Text Manipulation
```javascript
// Insert text
range.insertText("New text", "Replace" | "Start" | "End" | "Before" | "After");

// Get/set text
const text = range.text;
range.text = "Updated text";

// Delete content
range.delete();

// Search within range
const searchResults = range.search("searchTerm", { matchCase: true });
```

### Formatting
```javascript
// Font formatting
range.font.bold = true;
range.font.size = 12;
range.font.name = "Arial";
range.font.color = "#FF0000";

// Paragraph formatting
range.paragraphFormat.alignment = "Centered";
range.paragraphFormat.firstLineIndent = 36;
range.paragraphFormat.spaceAfter = 10;
```

### Selection & Navigation
```javascript
// Select the range
range.select();

// Get paragraphs in range
const paragraphs = range.paragraphs;

// Expand range
range.expandTo(otherRange);
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.range

---

## Paragraphs

### Access & Iterate
```javascript
// Get all document paragraphs
const paragraphs = context.document.body.paragraphs;
paragraphs.load("items");
await context.sync();

// Access specific paragraph
const firstPara = paragraphs.getFirst();

// Iterate
paragraphs.items.forEach(para => {
    console.log(para.text);
});
```

### Manipulation
```javascript
// Insert paragraph
body.insertParagraph("New paragraph", "End");

// Delete paragraph
paragraph.delete();

// Get/set text
paragraph.text = "Updated text";
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph

---

## Content Controls

Content controls are bounded regions that can contain text, images, tables, etc.

### Create & Access
```javascript
// Insert content control
const contentControl = range.insertContentControl();
contentControl.title = "MyControl";
contentControl.tag = "myTag";

// Find by tag
const controls = context.document.contentControls.getByTag("myTag");

// Find by title
const controlsByTitle = context.document.selectContentControlsByTitle("MyControl");
```

### Manipulation
```javascript
// Get/set text
contentControl.insertText("Text content", "Replace");

// Delete but keep content
contentControl.delete(true); // keepContent = true

// Delete including content
contentControl.delete(false);
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol

---

## Tables

### Create & Access
```javascript
// Insert table
const table = body.insertTable(3, 4, "End", ["Header1", "Header2", "Header3", "Header4"]);

// Access tables
const tables = context.document.body.tables;

// Get specific cell
const cell = table.getCell(0, 1); // row, column
```

### Manipulation
```javascript
// Add rows/columns
table.addRows("End", 2);
table.addColumns("End", 1);

// Set cell value
cell.value = "Cell content";

// Delete table
table.delete();
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.table

---

## Search & Replace

### Search
```javascript
// Basic search
const results = context.document.body.search("searchTerm");
results.load("items");

// Advanced search
const results = range.search("term", {
    matchCase: true,
    matchWholeWord: true,
    matchWildcards: false
});
```

### Replace
```javascript
// Search and replace all
results.items.forEach(result => {
    result.insertText("replacement", "Replace");
});
```

**Reference**: https://learn.microsoft.com/en-us/office/dev/add-ins/word/search-option-guidance

---

## Formatting

### Character Formatting (Font)
```javascript
range.font.bold = true;
range.font.italic = true;
range.font.underline = "Single";
range.font.size = 14;
range.font.color = "#0000FF";
range.font.highlightColor = "Yellow";
```

### Paragraph Formatting
```javascript
paragraph.alignment = "Left" | "Centered" | "Right" | "Justified";
paragraph.firstLineIndent = 36;
paragraph.leftIndent = 72;
paragraph.spaceAfter = 10;
paragraph.spaceBefore = 10;
paragraph.lineSpacing = 12;
```

### Styles
```javascript
// Apply style
paragraph.style = "Heading 1";
range.style = "Intense Quote";

// Get document styles
const styles = context.document.getStyles();
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.font

---

## Comments

### Add & Access Comments
```javascript
// Add comment
const comment = range.insertComment("This is a comment");

// Get all comments
const comments = context.document.comments;
comments.load("items");

// Reply to comment
comment.reply("This is a reply");
```

### Manage Comments
```javascript
// Delete comment
comment.delete();

// Get comment content
comment.content; // The comment text
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.comment

---

## Lists

### Create Lists
```javascript
// Create numbered list
const list = range.insertList(Word.ListLevelType.level1);

// Create bulleted list
paragraph.insertBullet("Start");
```

### Access Lists
```javascript
// Get all lists
const lists = context.document.lists;

// Get paragraphs in a list
const listParagraphs = context.document.listParagraphs;
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.list

---

## Headers & Footers

### Access
```javascript
// Get first section
const section = context.document.sections.getFirst();

// Access headers
const primaryHeader = section.getHeader("Primary");
const firstPageHeader = section.getHeader("FirstPage");

// Access footers
const primaryFooter = section.getFooter("Primary");
```

### Manipulation
```javascript
// Add content to header
primaryHeader.insertText("Document Title", "End");

// Add page numbers to footer
primaryFooter.insertText("Page ", "End");
primaryFooter.insertField("Page", "End");
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.headerfootercollection

---

## Images

### Insert Images
```javascript
// Insert from base64
const image = range.insertInlinePicture(base64String, "End");

// Insert from URL (requires internet access)
const image = range.insertInlinePicture(imageUrl, "End");
```

### Manipulate Images
```javascript
// Set dimensions
image.width = 200;
image.height = 150;

// Lock aspect ratio
image.lockAspectRatio = true;
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.inlinepicture

---

## Fields

Insert dynamic fields like page numbers, dates, etc.

```javascript
// Insert page number
range.insertField("Page", "End");

// Insert date
range.insertField("Date", "End");

// Insert other fields
range.insertField("NumPages", "End"); // Total pages
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.field

---

## Document Saving & Closing

### Save
```javascript
// Save document
await context.document.save();

// Note: In Word Online, save is automatic
// In desktop, this prompts user or saves to existing location
```

### Close (Desktop only)
```javascript
// Close with save
context.document.close("Save");

// Close without saving
context.document.close("SkipSave");

// Not supported in Word Online
```

---

## Loading Properties

The Office.js API uses explicit loading for performance.

```javascript
// Load specific properties
paragraph.load("text,style");

// Load all properties
paragraph.load("*");

// Load nested properties
range.load("font/size,font/name");

// Must sync to retrieve values
await context.sync();
console.log(paragraph.text);
```

---

## Common Patterns

### Basic Word.run Pattern
```javascript
await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("Hello World", "End");
    await context.sync();
});
```

### Error Handling
```javascript
try {
    await Word.run(async (context) => {
        // Your code
    });
} catch (error) {
    if (error.code === "ItemNotFound") {
        console.log("Item not found");
    }
    console.error(error);
}
```

### Batch Operations
```javascript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Queue multiple operations
    body.insertParagraph("Para 1", "End");
    body.insertParagraph("Para 2", "End");
    body.insertParagraph("Para 3", "End");
    
    // Execute all at once
    await context.sync();
});
```

---

## Events

Subscribe to document events.

```javascript
// Content control added
context.document.onContentControlAdded.add(eventHandler);

// Paragraph added/changed/deleted
context.document.onParagraphAdded.add(eventHandler);
context.document.onParagraphChanged.add(eventHandler);
context.document.onParagraphDeleted.add(eventHandler);

// Selection changed (Common API)
Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    eventHandler
);
```

**Reference**: https://learn.microsoft.com/en-us/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member

---

## Requirement Sets

Office.js features are versioned by requirement sets. Check support:

```javascript
if (Office.context.requirements.isSetSupported("WordApi", "1.3")) {
    // Use WordApi 1.3 features
}
```

**Latest Requirement Sets**:
- WordApi 1.1 (2016) - Basic document manipulation
- WordApi 1.3 (2019) - Comments, document properties
- WordApi 1.4 (2021) - Track changes, revision management
- WordApi 1.5 (Current) - Latest features

**Reference**: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

---

## Tips for LLM Integration

1. **Use Track Changes** - Let users review LLM suggestions before accepting
2. **Content Controls** - Useful for marking LLM-generated sections
3. **Batch Operations** - Queue multiple text changes, sync once
4. **Selection Context** - Get user's current selection to understand context
5. **Comments** - Add explanatory comments for complex changes
6. **Preserve Formatting** - Use `insertText` carefully to maintain existing styles

---

## Additional Resources

- **Full API Reference**: https://learn.microsoft.com/en-us/javascript/api/word
- **Word Add-ins Overview**: https://learn.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins-programming-overview
- **Code Samples**: https://github.com/OfficeDev/Office-Add-in-samples
- **Script Lab** (for testing): https://learn.microsoft.com/en-us/office/dev/add-ins/overview/explore-with-script-lab