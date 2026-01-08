# Phase 5: Track Changes + Core Polish

## Objective
Make Track Changes reliable and preserve formatting so AI edits feel native to Word.

## Success Criteria
- [ ] Text formatting (bold, italic, etc.) preserved after edits
- [ ] Track Changes enabled and working programmatically
- [ ] Changes appear in Word Review tab and can be accepted/rejected
- [ ] Edge cases handled gracefully (empty selections, large docs, etc.)
- [ ] Tested successfully on 5+ real documents from target users
- [ ] No crashes or unexpected behavior

## Task 1: Formatting Preservation
**Goal:** When Claude edits styled text, preserve the original formatting.

**Current problem:** Selected text with bold/italic/colors → API returns plain text → insertion loses all formatting.

**Desired behavior:**
- User selects text with formatting (e.g., some words bold, some italic)
- Claude edits the text
- Inserted text preserves the dominant formatting style from original selection
- Acceptable trade-off: If original had mixed formatting, new text gets the most common style

**Implementation approach:**
- Before sending to Claude: capture and store the selected range's font properties (name, size, bold, italic, color, etc.)
- After receiving Claude's response: apply stored formatting to the newly inserted text
- Use Office.js font properties API

**Test cases:**
- Bold text → edit → stays bold
- Italic text → edit → stays italic  
- Mixed formatting selection → edit → gets dominant style
- Hyperlinked text → edit → hyperlink preserved
- Highlighted text → edit → highlighting preserved

---

## Task 2: Track Changes Reliability
**Goal:** Ensure Track Changes is enabled programmatically and changes appear correctly in Word's Review interface.

**Current state verification needed:**
- Does current implementation actually enable Track Changes?
- Do edits appear as tracked changes (deletions red/strikethrough, insertions blue/underline)?
- Can users accept/reject changes using Word's Review tab?

**Desired behavior:**
- Add-in programmatically enables Track Changes before making edits
- All text replacements appear as tracked changes
- Changes are attributed to the add-in (or appropriate author)
- Users can review and accept/reject each change individually
- If Track Changes was already on, respect that state

**Test cases:**
- Document with Track Changes OFF → edit → changes tracked, can accept/reject
- Document with Track Changes ON → edit → changes tracked like manual edits  
- Multiple sequential edits → each appears as separate reviewable change
- Accept one change, reject another → both operations work correctly

---

## Task 3: Edge Case Handling
**Goal:** Handle problematic inputs gracefully with helpful error messages.

**Edge cases to handle:**

1. **Empty or whitespace-only selection**
   - Show error: "Please select some text to edit"
   - Don't make API call

2. **Very large selections (>5000 words)**
   - Show warning with word count: "This is 6,847 words. Processing may be slow and use significant API credits. Continue?"
   - Let user confirm or cancel
   - Consider suggesting smaller selections

3. **Claude returns empty or invalid response**
   - Show error: "Claude returned an empty response. Try rephrasing your prompt."
   - Don't modify document
   - Log error for debugging

4. **Selection spans complex structures**
   - Tables: Show error "Please select only text (tables not supported)"
   - Images: Show error "Please select only text (images not supported)"  
   - For MVP, graceful degradation is acceptable

5. **Network/API failures**
   - Timeout after 60 seconds
   - Show clear error message
   - Don't leave document in partial edit state

**Test document to create:**
- Create comprehensive test document with:
  - Normal paragraph
  - Heavily formatted text (multiple styles)
  - Very long paragraph (2000+ words)
  - Multiple short paragraphs
  - Bulleted list
  - Numbered list
  - Table
  - Image
- Run each preset prompt on each section
- Verify appropriate behavior for each

---

## Task 4: Real Document Testing
**Goal:** Validate add-in works on actual documents from target users.

**Documents to test:**
- Wife's student papers (2-3 recent papers)
- Employee's business documents (1-2 examples)
- Your own documents (2-3 examples)
- Mix of lengths: short (1 page), medium (3-5 pages), long (10+ pages)

**Testing protocol for each document:**
For each document, test:
1. Select single paragraph → Grammar check preset → Verify changes accurate and formatting preserved
2. Select paragraph with formatting → Simplify preset → Verify bold/italic preserved
3. Select sentence with hyperlink → Edit → Verify link preserved
4. Large selection (500+ words) → Verify no timeout or crash
5. Multiple edits in same document → Verify all tracked separately
6. Review tab → Accept some changes, reject others → Verify both work

**Issue tracking:**
- Document each problem found with:
  - Document type
  - Selection type  
  - Prompt used
  - Expected behavior
  - Actual behavior
  - Screenshot if relevant
- Fix critical issues before proceeding to Phase 4B