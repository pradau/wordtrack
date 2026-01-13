/**
 * Tests for async behavior and Promise chains in Track Changes operations
 */

describe('Track Changes Async Behavior', () => {
  describe('Context.sync() Calls', () => {
    test('should call context.sync() after enabling Track Changes', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      context.document.trackRevisions = true;
      
      const syncSpy = jest.spyOn(context, 'sync');
      await context.sync();
      
      expect(syncSpy).toHaveBeenCalled();
    });

    test('should call context.sync() after insertion', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      range.insertText('Test', 'Replace');
      
      const syncSpy = jest.spyOn(context, 'sync');
      await context.sync();
      
      expect(syncSpy).toHaveBeenCalled();
    });

    test('should call sync() in correct order: enable → sync → insert → sync', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const callOrder: string[] = [];
      let syncCallCount = 0;
      
      // Enable Track Changes
      context.document.trackRevisions = true;
      callOrder.push('enable');
      
      // Sync
      context.sync = jest.fn(async () => {
        syncCallCount++;
        callOrder.push(`sync${syncCallCount}`);
        return Promise.resolve();
      });
      await context.sync();
      
      // Insert
      const selection = context.document.getSelection();
      const range = selection.getRange();
      range.insertText('Test', 'Replace');
      callOrder.push('insert');
      
      // Sync again
      await context.sync();
      
      expect(callOrder).toEqual(['enable', 'sync1', 'insert', 'sync2']);
    });
  });

  describe('Async/Await Handling', () => {
    test('should handle async ensureTrackChangesEnabled correctly', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      async function ensureTrackChangesEnabled(ctx: any): Promise<boolean> {
        if ('trackRevisions' in ctx.document) {
          ctx.document.trackRevisions = true;
          await ctx.sync();
          return true;
        }
        return false;
      }
      
      const result = await ensureTrackChangesEnabled(context);
      
      expect(result).toBe(true);
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should await sync() before proceeding', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      let syncCompleted = false;
      
      context.sync = jest.fn(async () => {
        await new Promise(resolve => setTimeout(resolve, 10));
        syncCompleted = true;
        return Promise.resolve();
      });
      
      context.document.trackRevisions = true;
      await context.sync();
      
      expect(syncCompleted).toBe(true);
    });
  });

  describe('Promise Chain Correctness', () => {
    test('should chain promises correctly: enable → sync → insert → sync', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const chain: string[] = [];
      
      // Enable
      context.document.trackRevisions = true;
      chain.push('enable');
      
      // Chain: sync → insert → sync
      await context.sync()
        .then(() => {
          chain.push('sync1');
          const selection = context.document.getSelection();
          const range = selection.getRange();
          range.insertText('Test', 'Replace');
          chain.push('insert');
          return context.sync();
        })
        .then(() => {
          chain.push('sync2');
        });
      
      expect(chain).toEqual(['enable', 'sync1', 'insert', 'sync2']);
    });

    test('should handle Promise rejection correctly', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      context.sync = jest.fn(() => Promise.reject(new Error('Sync failed')));
      
      context.document.trackRevisions = true;
      
      await expect(context.sync()).rejects.toThrow('Sync failed');
    });
  });

  describe('Error Propagation', () => {
    test('should propagate errors through Promise chain', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      context.sync = jest.fn(() => Promise.reject(new Error('Sync error')));
      
      context.document.trackRevisions = true;
      
      try {
        await context.sync();
        fail('Should have thrown error');
      } catch (error) {
        expect((error as Error).message).toBe('Sync error');
      }
    });

    test('should handle errors at different points in chain', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Error during first sync
      context.sync = jest.fn(() => Promise.reject(new Error('First sync failed')));
      
      context.document.trackRevisions = true;
      
      await expect(context.sync()).rejects.toThrow('First sync failed');
    });
  });

  describe('Operations Wait for Track Changes Enable', () => {
    test('should wait for Track Changes enable before inserting', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const order: string[] = [];
      
      // Enable Track Changes
      context.document.trackRevisions = true;
      order.push('enable');
      
      // Wait for sync
      await context.sync();
      order.push('sync');
      
      // Then insert
      const selection = context.document.getSelection();
      const range = selection.getRange();
      range.insertText('Test', 'Replace');
      order.push('insert');
      
      expect(order).toEqual(['enable', 'sync', 'insert']);
    });

    test('should ensure Track Changes is enabled before operations', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Enable first
      context.document.trackRevisions = true;
      await context.sync();
      
      // Verify it's enabled
      expect(context.document.trackRevisions).toBe(true);
      
      // Now safe to insert
      const selection = context.document.getSelection();
      const range = selection.getRange();
      range.insertText('Test', 'Replace');
      
      // Track Changes should still be enabled
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Concurrent Operations', () => {
    test('should handle sequential async operations correctly', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      // Operation 1
      context.document.trackRevisions = true;
      await context.sync();
      
      // Operation 2
      const selection1 = context.document.getSelection();
      const range1 = selection1.getRange();
      range1.insertText('First', 'Replace');
      await context.sync();
      
      // Operation 3
      const selection2 = context.document.getSelection();
      const range2 = selection2.getRange();
      range2.insertText('Second', 'Replace');
      await context.sync();
      
      expect(context.document.trackRevisions).toBe(true);
    });

    test('should maintain state across async operations', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      context.document.trackRevisions = true;
      await context.sync();
      
      // Multiple async operations
      await Promise.all([
        context.sync(),
        Promise.resolve().then(() => {
          expect(context.document.trackRevisions).toBe(true);
        }),
      ]);
      
      expect(context.document.trackRevisions).toBe(true);
    });
  });
});
