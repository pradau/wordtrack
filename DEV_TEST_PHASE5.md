# Phase 5 Testing Guide: Formatting Preservation

## Implementation Analysis

### What We Capture
The current implementation captures and preserves:
- Font name (`name`)
- Font size (`size`)
- Bold (`bold`)
- Italic (`italic`)
- Underline (`underline`)
- Text color (`color`)
- Highlight color (`highlightColor`)

### What We Don't Capture
These formatting properties are **not** currently preserved:
- Hyperlinks (URLs and link formatting)
- Subscript/superscript
- Strikethrough
- Text effects (shadow, outline, glow, etc.)
- Character spacing/kerning
- Word styles (we capture font properties but not the style name itself)

---

## Test Case Categories

### Category A: Will Definitely Fail

These test cases will fail because the properties are not captured:

1. **Hyperlinked Text**
   - **Test:** Select text with a hyperlink → edit → check if link is preserved
   - **Expected:** Link will be lost (only text formatting preserved)
   - **Status:** FAIL - Not implemented

2. **Subscript/Superscript**
   - **Test:** Select text with subscript (H₂O) or superscript (x²) → edit
   - **Expected:** Subscript/superscript will be lost
   - **Status:** FAIL - Not implemented

3. **Strikethrough Text**
   - **Test:** Select strikethrough text → edit
   - **Expected:** Strikethrough will be lost
   - **Status:** FAIL - Not implemented

4. **Text Effects (Shadow, Outline, etc.)**
   - **Test:** Select text with shadow or outline effects → edit
   - **Expected:** Effects will be lost
   - **Status:** FAIL - Not implemented

---

### Category B: High Probability of Failure

These may fail depending on Office.js behavior with mixed formatting:

5. **Mixed Formatting Within Selection**
   - **Test:** Select text where some words are bold, others italic, others plain
   - **Expected:** May get dominant style OR null/undefined properties
   - **Status:** MAY FAIL - Depends on Office.js behavior
   - **Note:** Office.js may return `null` for properties that are mixed

6. **Multiple Paragraphs with Different Formatting**
   - **Test:** Select multiple paragraphs where each has different formatting
   - **Expected:** Only one set of formatting properties captured (from first paragraph or dominant)
   - **Status:** MAY FAIL - Will use range's overall formatting, may not match all paragraphs

7. **Style-Based Formatting with Additional Properties**
   - **Test:** Select text with a Word style (e.g., "Heading 1, Italic") that has style-level formatting
   - **Expected:** Font properties preserved, but style name and style-specific properties may be lost
   - **Status:** PARTIAL - Font properties work, but style context may be lost
   - **Note:** Your test showed this works for italicized Heading 1, so font properties are preserved

8. **Very Long Selections (500+ words)**
   - **Test:** Select very long text that likely has mixed formatting
   - **Expected:** May have mixed formatting issues
   - **Status:** MAY FAIL - Depends on whether selection has uniform formatting

---

### Category C: Low Probability of Failure

These should work but may have edge cases:

9. **Combined Formatting (Bold + Italic + Color)**
   - **Test:** Select text that is bold, italic, and colored
   - **Expected:** All properties should be preserved
   - **Status:** SHOULD WORK - All properties are captured
   - **Edge Case:** If Office.js returns null for any property, that one will be skipped

10. **Underlined Text**
    - **Test:** Select underlined text → edit
    - **Expected:** Underline should be preserved
    - **Status:** SHOULD WORK - Underline is captured
    - **Edge Case:** Different underline types might not all be preserved

11. **Highlighted Text**
    - **Test:** Select highlighted text → edit
    - **Expected:** Highlight color should be preserved
    - **Status:** SHOULD WORK - Highlight color is captured
    - **Edge Case:** Some highlight colors might not map correctly

12. **Font Name and Size**
    - **Test:** Select text in specific font (Arial, Times New Roman) and size (12pt, 14pt) → edit
    - **Expected:** Font name and size should be preserved
    - **Status:** SHOULD WORK - Both are captured
    - **Edge Case:** Uncommon fonts might not be available

---

### Category D: Will Never Fail

These are guaranteed to work:

13. **Plain Text (No Formatting)**
    - **Test:** Select plain text with default formatting → edit
    - **Expected:** Stays plain (nothing to preserve, so no failure possible)
    - **Status:** ALWAYS WORKS

14. **Bold Text (Uniform)**
    - **Test:** Select text that is uniformly bold → edit
    - **Expected:** Stays bold
    - **Status:** ALWAYS WORKS - Bold is explicitly captured and applied

15. **Italic Text (Uniform)**
    - **Test:** Select text that is uniformly italic → edit
    - **Expected:** Stays italic
    - **Status:** ALWAYS WORKS - Italic is explicitly captured and applied

16. **Text Color (Uniform)**
    - **Test:** Select text with uniform color → edit
    - **Expected:** Color is preserved
    - **Status:** ALWAYS WORKS - Color is explicitly captured and applied

17. **Single Paragraph with Uniform Formatting**
    - **Test:** Select a single paragraph where all text has the same formatting
    - **Expected:** All formatting preserved
    - **Status:** ALWAYS WORKS - No mixed formatting issues

---

## Testing Protocol

### Setup
1. Create a test document with various formatting scenarios
2. Start WordTrack using `./start.sh`
3. Open the test document in Word

### For Each Test Case

1. **Select the formatted text** in Word
2. **Click "Get Selected Text"** in WordTrack
3. **Verify the text displays correctly** in the task pane
4. **Select a prompt** (e.g., "Improve clarity and readability")
5. **Click "Send to Claude"**
6. **Wait for response**
7. **Click "Insert Claude's Response"**
8. **Verify formatting** in Word:
   - Check if bold/italic/color/etc. are preserved
   - Compare with original formatting
   - Note any differences

### Documentation Template

For each test, record:
```
Test #: [Number]
Category: [A/B/C/D]
Formatting Type: [e.g., Bold + Italic]
Selection: [Brief description]
Prompt Used: [e.g., "Improve clarity"]
Result: [PASS / FAIL / PARTIAL]
Notes: [Any observations, edge cases, or issues]
```

---

## Test Document Creation

Create a Word document with these sections:

### Section 1: Basic Formatting
- **Bold text** - uniformly bold paragraph
- *Italic text* - uniformly italic paragraph
- **Bold and *italic* combined** - mixed within paragraph
- <u>Underlined text</u> - uniformly underlined
- Colored text (red, blue, green) - different colors
- Highlighted text - yellow highlight

### Section 2: Advanced Formatting
- [Hyperlinked text](https://example.com) - text with hyperlink
- H₂O - subscript example
- x² - superscript example
- ~~Strikethrough text~~ - strikethrough example
- Text with shadow effect (if available)

### Section 3: Style-Based Formatting
- Heading 1 style
- Heading 1, Italic (modified style)
- Heading 2 style
- Normal style with bold

### Section 4: Edge Cases
- Very long paragraph (500+ words) with uniform formatting
- Multiple paragraphs with different formatting selected together
- Mixed formatting: some bold, some italic, some plain in one selection

### Section 5: Combined Formatting
- Bold + Italic + Colored text
- Underlined + Highlighted text
- Font name (Arial) + Size (14pt) + Bold

---

## Expected Results Summary

| Test Case | Expected Result | Confidence |
|-----------|----------------|------------|
| Plain text | ALWAYS WORKS | 100% |
| Uniform bold | ALWAYS WORKS | 100% |
| Uniform italic | ALWAYS WORKS | 100% |
| Uniform color | ALWAYS WORKS | 100% |
| Combined formatting | SHOULD WORK | 95% |
| Underline | SHOULD WORK | 95% |
| Highlight | SHOULD WORK | 95% |
| Font name/size | SHOULD WORK | 95% |
| Style-based (font props) | SHOULD WORK | 90% |
| Mixed formatting | MAY USE DOMINANT STYLE | 70% |
| Multiple paragraphs | MAY NOT MATCH ALL | 60% |
| Hyperlinks | WILL FAIL | 0% |
| Subscript/superscript | WILL FAIL | 0% |
| Strikethrough | WILL FAIL | 0% |
| Text effects | WILL FAIL | 0% |

---

## Next Steps After Testing

1. **Document all failures** - Create a list of what doesn't work
2. **Prioritize fixes** - Decide which failures are critical vs. acceptable
3. **Enhance implementation** - Add support for high-priority missing features
4. **Update DEV_PLAN_PHASE5.md** - Mark Task 1 as complete with known limitations

---

## Notes

- The current implementation uses Office.js `range.font` properties, which gives the formatting of the range as a whole
- For mixed formatting, Office.js may return `null` or `undefined` for properties that vary
- The code handles `undefined` properties gracefully by only applying defined properties
- Hyperlinks require separate Office.js APIs (`range.hyperlinks`) and are not currently implemented
- Subscript/superscript use `range.font.subscript` and `range.font.superscript` properties (not currently captured)
