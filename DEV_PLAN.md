# WordTrack Development Plan

## Purpose
WordTrack is a Microsoft Word add-in designed to enhance document editing by 
integrating AI-powered text improvement (via Claude API) with Word's native 
editing tools. The add-in leverages Word's Track Changes feature, formatting 
capabilities, and document structure to provide seamless AI-assisted editing 
while maintaining full compatibility with Word's existing functionality.

## Current Status
**Phase 3 Complete** - Core functionality working (extraction, API, insertion)
**Phase 4A Complete** - User-friendly startup with temp document hiding
**Phase 5 Updates Complete** - Removed buggy Track Changes auto-enable, added context menu, improved prompts, added Help system
**Total time invested: ~12-15 hours**

## Next Phases (Recommended Order)

### Phase 5: Track Changes + Core Polish (UPDATED)

**Completed:**
- Removed programmatic Track Changes enabling (was buggy/unreliable)
- Added Context menu (Personal/Formal/Documentation) for tone guidance
- Optimized prompts for token efficiency (12 useful options)
- Added Help system with concise instructions
- Added debug log download feature
- Removed Test button (replaced with Help)

**Remaining:**
- Fix table selection bug (BUG-001) - General Exception when inserting into table selections
- Test with real documents from target users:
  - Wife's student papers (varied lengths, formatting)
  - Employee's documents (business writing)
  - Your own use cases
- Handle edge cases (tables, complex formatting, large documents)
- Verify Track Changes behavior when manually enabled by users

**Decision point:** Track Changes is now user-controlled (manual enable). Focus on ensuring 
WordTrack works well when Track Changes is enabled by users.

**Deliverable:** Stable editing experience with proper error handling for edge cases

### Phase 6: Context and Category System
**Priority: MEDIUM - Context menu already implemented**

**Completed:**
- Context selector (Personal/Formal/Documentation) - DONE
- Context integrated into system prompts for tone guidance

**Potential Future Enhancements:**
- Category selector (dynamically filtered by Context)
  - Personal: Email, Notes, Creative Writing
  - Professional: MRI/Neuroradiology, Personal Training, Business
- `categories.json` config with context+category-specific:
  - Guidelines for Claude
  - Curated prompt sets
  - Terminology preferences
- Prompt dropdown filtered by context+category
- localStorage persistence for user preferences

**Design consideration:** Current context system provides basic tone guidance. 
Category system can be added if needed after real-world usage testing.

**Deliverable:** Current context system is sufficient for MVP. Category system 
can be added later if user feedback indicates need.

### Phase 7: Optional Polish 
**Priority: LOW - Only if highly valuable**

Potential features (prioritize based on user feedback):
- UI styling improvements
- Additional preset prompts per category
- Performance optimizations (caching, batching)
- Advanced error recovery
- Technical dictionaries (MRI terms, etc.)
- Usage statistics/tracking
- Export/import category configurations

**Decision point:** After 2-4 weeks of real use, evaluate what polish actually 
matters vs. what's just nice-to-have.

## Decision Points

**After Phase 5:**
- ✅ Track Changes approach: User-controlled (manual enable) - DECIDED
- ✅ Context menu implemented - DONE
- ⚠️ Table selection bug needs fixing before production use

**After Phase 6:**
- Are context/categories genuinely useful? (Test with wife/employee)
- Do category-specific prompts improve results meaningfully?
- Is current feature set sufficient for daily use?

**After 2-4 weeks of real use:**
- What friction points remain?
- What features would 10x the value?
- Is Phase 6 worth 8-12 more hours?

## Revised Time Estimates
- **Phase 5:** 3-5 hours (Track Changes polish) - PARTIALLY COMPLETE
  - Completed: Context menu, prompt optimization, Help system (~2 hours)
  - Remaining: Bug fixes, edge case handling (~2-3 hours)
- **Phase 6:** 6-8 hours (Category system) - DEFERRED (basic context done)
- **Phase 7:** 8-12 hours (Optional polish)
- **Total remaining:** 10-20 hours (reduced due to Phase 5/6 progress)

**Total project time if completed:** 22-30 hours

## Bug Tracking

See **BUGS.md** for identified bugs and issues.

## Principles
- ✅ Validate foundation before building structure 
- ✅ Real user testing drives priorities
- ✅ Stop when "good enough" for target users
- ✅ Time estimates are speculative - adjust as you learn