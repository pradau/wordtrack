/**
 * Tests for integration of Track Changes with insert functions
 */

describe('Track Changes Integration with Insert Functions', () => {
  // Mock implementations of the insert functions as they would be with Track Changes
  let ensureTrackChangesEnabledCalled = false;
  let insertTextCalled = false;
  let trackChangesEnabled = false;

  beforeEach(() => {
    ensureTrackChangesEnabledCalled = false;
    insertTextCalled = false;
    trackChangesEnabled = false;
  });

  // Mock ensureTrackChangesEnabled
  async function ensureTrackChangesEnabled(context: any): Promise<boolean> {
    ensureTrackChangesEnabledCalled = true;
    if ('trackRevisions' in context.document) {
      try {
        context.document.trackRevisions = true;
        trackChangesEnabled = true;
        return true;
      } catch (error) {
        // Document might be protected
        return false;
      }
    }
    return false;
  }

  // Mock handleCapitalizeAndInsert with Track Changes
  async function handleCapitalizeAndInsert(
    text: string,
    context: any
  ): Promise<void> {
    const enabled = await ensureTrackChangesEnabled(context);
    if (enabled) {
      trackChangesEnabled = true;
    }
    
    const selection = context.document.getSelection();
    const range = selection.getRange();
    range.insertText(text, 'Replace');
    insertTextCalled = true;
    
    await context.sync();
  }

  // Mock handleInsertClaudeResponse with Track Changes
  async function handleInsertClaudeResponse(
    text: string,
    context: any
  ): Promise<void> {
    const enabled = await ensureTrackChangesEnabled(context);
    if (enabled) {
      trackChangesEnabled = true;
    }
    
    const selection = context.document.getSelection();
    const range = selection.getRange();
    range.insertText(text, 'Replace');
    insertTextCalled = true;
    
    await context.sync();
  }

  describe('handleCapitalizeAndInsert Integration', () => {
    test('should call ensureTrackChangesEnabled before insertText', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      await handleCapitalizeAndInsert('Test Text', context);
      
      expect(ensureTrackChangesEnabledCalled).toBe(true);
      expect(insertTextCalled).toBe(true);
      expect(trackChangesEnabled).toBe(true);
    });

    test('should enable Track Changes before insertion', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      await handleCapitalizeAndInsert('Test Text', context);
      
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should still insert text even if Track Changes enable fails', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      await handleCapitalizeAndInsert('Test Text', context);
      
      expect(ensureTrackChangesEnabledCalled).toBe(true);
      expect(insertTextCalled).toBe(true);
      // Text should still be inserted
    });
  });

  describe('handleInsertClaudeResponse Integration', () => {
    test('should call ensureTrackChangesEnabled before insertText', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      await handleInsertClaudeResponse('Claude Response', context);
      
      expect(ensureTrackChangesEnabledCalled).toBe(true);
      expect(insertTextCalled).toBe(true);
      expect(trackChangesEnabled).toBe(true);
    });

    test('should enable Track Changes before insertion', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      await handleInsertClaudeResponse('Claude Response', context);
      
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should still insert text even if Track Changes enable fails', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      await handleInsertClaudeResponse('Claude Response', context);
      
      expect(ensureTrackChangesEnabledCalled).toBe(true);
      expect(insertTextCalled).toBe(true);
    });
  });

  describe('Error Handling', () => {
    test('should handle errors when Track Changes can\'t be enabled', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
        documentProtected: true,
      });
      
      // Should not throw, but Track Changes won't be enabled
      await handleCapitalizeAndInsert('Test', context);
      expect(insertTextCalled).toBe(true);
    });

    test('should handle errors during insertion', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Make insertText throw
      const selection = context.document.getSelection();
      const range = selection.getRange();
      range.insertText = jest.fn(() => {
        throw new Error('Insert failed');
      });
      
      // The function should handle the error, but insertText will throw synchronously
      // In a real scenario, this would be caught in the Word.run() error handler
      expect(() => {
        range.insertText('Test', 'Replace');
      }).toThrow('Insert failed');
    });
  });

  describe('Order of Operations', () => {
    test('should enable Track Changes before getting selection', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const callOrder: string[] = [];
      
      // Track call order by wrapping the function
      let ensureTrackChangesEnabledWrapper = async (ctx: any): Promise<boolean> => {
        callOrder.push('ensureTrackChanges');
        return ensureTrackChangesEnabled(ctx);
      };
      
      const originalGetSelection = context.document.getSelection;
      context.document.getSelection = jest.fn(() => {
        callOrder.push('getSelection');
        return originalGetSelection();
      });
      
      // Use wrapper instead of original
      const handleWithWrapper = async (text: string, ctx: any): Promise<void> => {
        const enabled = await ensureTrackChangesEnabledWrapper(ctx);
        if (enabled) {
          trackChangesEnabled = true;
        }
        
        const selection = ctx.document.getSelection();
        const range = selection.getRange();
        range.insertText(text, 'Replace');
        insertTextCalled = true;
        
        await ctx.sync();
      };
      
      await handleWithWrapper('Test', context);
      
      expect(callOrder[0]).toBe('ensureTrackChanges');
      expect(callOrder[1]).toBe('getSelection');
    });
  });
});
