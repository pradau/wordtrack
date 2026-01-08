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
**Total time invested: ~8-10 hours over 2 days**

## Next Phases (Recommended Order)

### Phase 5: Track Changes + Core Polish 

Core improvements:
- Refine Track Changes reliability (handle edge cases)
- Test with real documents from target users:
  - Wife's student papers (varied lengths, formatting)
  - Employee's documents (business writing)
  - Your own use cases
- Fix formatting preservation issues
- Handle large insertions/deletions gracefully
- Verify Track Changes behavior across document types

**Decision point:** If Track Changes has fundamental limitations, address them 
before building category system on top.

**Deliverable:** Rock-solid Track Changes that works reliably for target users

### Phase 6: Context and Category System
**Priority: HIGH - Build on proven foundation**

Features:
- Context selector (Personal/Professional)
- Category selector (dynamically filtered by Context)
  - Personal: Email, Notes, Creative Writing
  - Professional: MRI/Neuroradiology, Personal Training, Business
- `categories.json` config with context+category-specific:
  - Guidelines for Claude
  - Curated prompt sets
  - Terminology preferences
- Prompt dropdown filtered by context+category
- localStorage persistence for user preferences
- API calls include category-specific context

**Design consideration:** After Phase 5 testing, you'll know what editing 
patterns work best, so category prompts can be optimized accordingly.

**Deliverable:** Tailored editing experience for different document types

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
- ✅ Track Changes reliable? → Proceed 
- ❌ Fundamental issues? → Pivot approach or stop

**After Phase 6:**
- Are context/categories genuinely useful? (Test with wife/employee)
- Do category-specific prompts improve results meaningfully?
- Is current feature set sufficient for daily use?

**After 2-4 weeks of real use:**
- What friction points remain?
- What features would 10x the value?
- Is Phase 6 worth 8-12 more hours?

## Revised Time Estimates
- **Phase 5:** 3-5 hours (Track Changes polish)
- **Phase 6:** 6-8 hours (Category system)
- **Phase 7:** 8-12 hours (Optional polish)
- **Total remaining:** 17-25 hours

**Total project time if completed:** 25-35 hours

## Principles
- ✅ Validate foundation before building structure 
- ✅ Real user testing drives priorities
- ✅ Stop when "good enough" for target users
- ✅ Time estimates are speculative - adjust as you learn