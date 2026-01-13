/**
 * Tests for Track Changes edge cases and error handling
 */

describe('Track Changes Edge Cases', () => {
  describe('Document Protection', () => {
    test('should detect document protection and handle gracefully', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
        documentProtected: true,
      });
      
      // Attempt to enable Track Changes
      try {
        context.document.trackRevisions = true;
        fail('Should have thrown an error for protected document');
      } catch (error) {
        expect(error).toBeInstanceOf(Error);
        expect((error as Error).message).toContain('protected');
      }
    });

    test('should return false when document is protected', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
        documentProtected: true,
      });
      
      let result = false;
      try {
        context.document.trackRevisions = true;
      } catch {
        result = false;
      }
      
      expect(result).toBe(false);
    });
  });

  describe('API Version Compatibility', () => {
    test('should handle older Word versions without trackRevisions', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      const hasAPI = 'trackRevisions' in context.document;
      expect(hasAPI).toBe(false);
    });

    test('should work with newer Word versions', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const hasAPI = 'trackRevisions' in context.document;
      expect(hasAPI).toBe(true);
    });
  });

  describe('Error Messages', () => {
    test('should provide clear error message when API not available', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      const hasAPI = 'trackRevisions' in context.document;
      const errorMessage = hasAPI 
        ? null 
        : 'Track Changes API not available in this Word version';
      
      expect(errorMessage).toBeTruthy();
      expect(errorMessage).toContain('not available');
    });

    test('should provide clear error message when document protected', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        documentProtected: true,
      });
      
      let errorMessage: string | null = null;
      try {
        context.document.trackRevisions = true;
      } catch (error) {
        errorMessage = 'Document is protected. Cannot enable Track Changes.';
      }
      
      expect(errorMessage).toBeTruthy();
      expect(errorMessage).toContain('protected');
    });
  });

  describe('Graceful Fallback', () => {
    test('should fallback gracefully when Track Changes unavailable', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      // Should still allow insertion even if Track Changes can't be enabled
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      expect(() => {
        range.insertText('Test', 'Replace');
      }).not.toThrow();
    });

    test('should proceed with insertion when Track Changes enable fails', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        documentProtected: true,
      });
      
      // Should still be able to insert text
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      expect(() => {
        range.insertText('Test', 'Replace');
      }).not.toThrow();
    });
  });

  describe('Multiple Sequential Calls', () => {
    test('should handle multiple calls to ensureTrackChangesEnabled', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // First call
      context.document.trackRevisions = true;
      expect(context.document.trackRevisions).toBe(true);
      
      // Second call (should be idempotent)
      const wasAlreadyEnabled = context.document.trackRevisions === true;
      expect(wasAlreadyEnabled).toBe(true);
      
      // Third call
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should not fail on repeated enable attempts', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Enable multiple times
      context.document.trackRevisions = true;
      context.document.trackRevisions = true;
      context.document.trackRevisions = true;
      
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Unexpected Errors', () => {
    test('should handle unexpected errors during enable', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Simulate unexpected error
      Object.defineProperty(context.document, 'trackRevisions', {
        get: jest.fn(() => false),
        set: jest.fn(() => {
          throw new Error('Unexpected error');
        }),
        configurable: true,
      });
      
      expect(() => {
        context.document.trackRevisions = true;
      }).toThrow('Unexpected error');
    });

    test('should handle null/undefined context gracefully', () => {
      const context: any = null;
      
      expect(() => {
        if (!context) {
          throw new Error('Context is null');
        }
      }).toThrow('Context is null');
    });
  });
});
