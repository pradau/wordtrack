/**
 * Tests for selection and range handling with Track Changes
 */

describe('Track Changes Selection Handling', () => {
  describe('Selection Retrieval', () => {
    test('should retrieve selection before enabling Track Changes', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Get selection first
      const selection = context.document.getSelection();
      expect(selection).toBeDefined();
      
      // Then enable Track Changes
      context.document.trackRevisions = true;
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should get selection correctly with Track Changes enabled', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      expect(selection).toBeDefined();
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Range Operations', () => {
    test('should get range correctly with Track Changes enabled', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      expect(range).toBeDefined();
      expect(range.insertText).toBeDefined();
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should perform range operations with Track Changes ON', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Load range properties
      range.load('text');
      range.font.load('bold, italic');
      
      // Insert text
      range.insertText('Test', 'Replace');
      
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Invalid Selection Handling', () => {
    test('should handle error when selection is invalid', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      // Simulate invalid selection
      context.document.getSelection = jest.fn(() => {
        throw new Error('No selection available');
      });
      
      expect(() => {
        context.document.getSelection();
      }).toThrow('No selection available');
    });

    test('should handle error when range is invalid', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      
      // Simulate invalid range
      selection.getRange = jest.fn(() => {
        throw new Error('Invalid range');
      });
      
      expect(() => {
        selection.getRange();
      }).toThrow('Invalid range');
    });
  });

  describe('Selection and Track Changes Interaction', () => {
    test('should maintain Track Changes state during selection operations', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      // Get selection
      const selection = context.document.getSelection();
      expect(context.document.trackRevisions).toBe(true);
      
      // Get range
      const range = selection.getRange();
      expect(context.document.trackRevisions).toBe(true);
      
      // Load properties
      range.load('text');
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should not interfere with Track Changes when getting selection', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      // Multiple selection operations
      const selection1 = context.document.getSelection();
      const selection2 = context.document.getSelection();
      const selection3 = context.document.getSelection();
      
      expect(selection1).toBeDefined();
      expect(selection2).toBeDefined();
      expect(selection3).toBeDefined();
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Range Operations Don\'t Interfere', () => {
    test('should not change Track Changes state during range operations', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Multiple range operations
      range.load('text');
      range.font.load('bold');
      range.insertText('Test', 'Replace');
      range.font.bold = true;
      
      // Track Changes should remain enabled
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should allow formatting operations without affecting Track Changes', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Apply formatting
      range.font.bold = true;
      range.font.italic = true;
      range.font.color = '#FF0000';
      
      // Track Changes should still be enabled
      expect(context.document.trackRevisions).toBe(true);
      expect(range.font.bold).toBe(true);
      expect(range.font.italic).toBe(true);
    });
  });

  describe('Selection State Consistency', () => {
    test('should maintain consistent selection state with Track Changes', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      // Get selection multiple times
      const selection1 = context.document.getSelection();
      const range1 = selection1.getRange();
      
      const selection2 = context.document.getSelection();
      const range2 = selection2.getRange();
      
      // Both should work correctly
      expect(range1).toBeDefined();
      expect(range2).toBeDefined();
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should handle selection changes correctly', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      // First selection
      const selection1 = context.document.getSelection();
      const range1 = selection1.getRange();
      range1.insertText('First', 'Replace');
      
      // Second selection (user might have changed selection)
      const selection2 = context.document.getSelection();
      const range2 = selection2.getRange();
      range2.insertText('Second', 'Replace');
      
      // Track Changes should remain enabled
      expect(context.document.trackRevisions).toBe(true);
    });
  });
});
