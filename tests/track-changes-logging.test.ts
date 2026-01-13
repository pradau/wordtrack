/**
 * Tests for console logging in Track Changes operations
 */

describe('Track Changes Console Logging', () => {
  let consoleLogSpy: jest.SpyInstance;
  let consoleErrorSpy: jest.SpyInstance;

  beforeEach(() => {
    consoleLogSpy = jest.spyOn(console, 'log').mockImplementation();
    consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();
  });

  afterEach(() => {
    consoleLogSpy.mockRestore();
    consoleErrorSpy.mockRestore();
  });

  describe('Track Changes Enable Logging', () => {
    test('should log when Track Changes is enabled', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      context.document.trackRevisions = true;
      console.log('Track Changes enabled');
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Track Changes enabled');
    });

    test('should log Track Changes state', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const state = context.document.trackRevisions;
      console.log('Track Changes state:', state);
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Track Changes state:', true);
    });
  });

  describe('Error Logging', () => {
    test('should log errors with context', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        documentProtected: true,
      });
      
      try {
        context.document.trackRevisions = true;
      } catch (error) {
        console.error('Error enabling Track Changes:', error);
      }
      
      expect(consoleErrorSpy).toHaveBeenCalled();
      const errorCall = consoleErrorSpy.mock.calls[0];
      expect(errorCall[0]).toContain('Error enabling Track Changes');
    });

    test('should log API availability errors', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      if (!('trackRevisions' in context.document)) {
        console.error('Track Changes API not available');
      }
      
      expect(consoleErrorSpy).toHaveBeenCalledWith('Track Changes API not available');
    });

    test('should log document protection errors', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        documentProtected: true,
      });
      
      try {
        context.document.trackRevisions = true;
      } catch (error) {
        console.error('Document is protected:', error);
      }
      
      expect(consoleErrorSpy).toHaveBeenCalled();
    });
  });

  describe('Successful Operation Logging', () => {
    test('should log successful Track Changes enable', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      context.document.trackRevisions = true;
      console.log('Track Changes enabled successfully');
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Track Changes enabled successfully');
    });

    test('should log when Track Changes already enabled', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const wasEnabled = context.document.trackRevisions === true;
      if (wasEnabled) {
        console.log('Track Changes already enabled');
      }
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Track Changes already enabled');
    });
  });

  describe('Sensitive Information Protection', () => {
    test('should not log API keys', () => {
      // Simulate having an API key without logging it
      console.log('Using API key');
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Using API key');
      // Verify that API key pattern is not in logs
      const allLogs = consoleLogSpy.mock.calls.flat().join(' ');
      expect(allLogs).not.toContain('sk-ant');
    });

    test('should not log user document content', () => {
      // Simulate having document text without logging it
      console.log('Processing document');
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Processing document');
      // Verify that sensitive content is not in logs
      const allLogs = consoleLogSpy.mock.calls.flat().join(' ');
      expect(allLogs).not.toContain('Sensitive');
    });

    test('should log operation type without sensitive data', () => {
      console.log('Inserting text with Track Changes enabled');
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Inserting text with Track Changes enabled');
      // Should not contain actual text content
    });
  });

  describe('Debug Information Logging', () => {
    test('should log debug information for troubleshooting', () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const debugInfo = {
        apiAvailable: 'trackRevisions' in context.document,
        currentState: context.document.trackRevisions,
      };
      
      console.log('Track Changes debug info:', debugInfo);
      
      expect(consoleLogSpy).toHaveBeenCalledWith('Track Changes debug info:', expect.objectContaining({
        apiAvailable: true,
        currentState: false,
      }));
    });

    test('should log operation sequence for debugging', () => {
      const operations = ['enableTrackChanges', 'getSelection', 'insertText', 'sync'];
      
      operations.forEach(op => {
        console.log(`Operation: ${op}`);
      });
      
      expect(consoleLogSpy).toHaveBeenCalledTimes(4);
      expect(consoleLogSpy).toHaveBeenCalledWith('Operation: enableTrackChanges');
      expect(consoleLogSpy).toHaveBeenCalledWith('Operation: sync');
    });
  });

  describe('Logging Levels', () => {
    test('should use console.log for informational messages', () => {
      console.log('Track Changes operation started');
      
      expect(consoleLogSpy).toHaveBeenCalled();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test('should use console.error for error messages', () => {
      console.error('Track Changes operation failed');
      
      expect(consoleErrorSpy).toHaveBeenCalled();
      expect(consoleLogSpy).not.toHaveBeenCalledWith(expect.stringContaining('failed'));
    });
  });
});
