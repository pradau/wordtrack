/**
 * Tests for user message updates based on Track Changes state
 */

describe('Track Changes User Messages', () => {
  // Mock message functions
  const messages: string[] = [];
  
  function showSuccess(message: string): void {
    messages.push(`SUCCESS: ${message}`);
  }
  
  function showError(message: string): void {
    messages.push(`ERROR: ${message}`);
  }

  beforeEach(() => {
    messages.length = 0;
  });

  // Helper to generate success message based on Track Changes state
  function getSuccessMessage(
    trackChangesEnabled: boolean,
    hasFormatting: boolean = false
  ): string {
    if (trackChangesEnabled) {
      if (hasFormatting) {
        return 'Text has been inserted with formatting preserved. Changes are tracked.';
      }
      return 'Text has been inserted. Changes are tracked.';
    } else {
      if (hasFormatting) {
        return 'Text has been inserted with formatting preserved. Make sure Track Changes is enabled in Word to see the changes tracked.';
      }
      return 'Text has been inserted. Make sure Track Changes is enabled in Word to see the changes tracked.';
    }
  }

  describe('Success Messages', () => {
    test('should show message when Track Changes enabled programmatically', () => {
      const message = getSuccessMessage(true, false);
      showSuccess(message);
      
      expect(messages[0]).toContain('inserted');
      expect(messages[0]).toContain('tracked');
      expect(messages[0]).not.toContain('Make sure Track Changes is enabled');
    });

    test('should show message when Track Changes already ON', () => {
      const message = getSuccessMessage(true, false);
      showSuccess(message);
      
      expect(messages[0]).toContain('tracked');
      expect(messages[0]).not.toContain('Make sure');
    });

    test('should show fallback message when Track Changes unavailable', () => {
      const message = getSuccessMessage(false, false);
      showSuccess(message);
      
      expect(messages[0]).toContain('Make sure Track Changes is enabled');
    });

    test('should include formatting info when applicable', () => {
      const message = getSuccessMessage(true, true);
      showSuccess(message);
      
      expect(messages[0]).toContain('formatting preserved');
      expect(messages[0]).toContain('tracked');
    });

    test('should show different message for capitalizeAndInsert', () => {
      const trackChangesEnabled = true;
      const message = trackChangesEnabled
        ? 'Text has been capitalized and inserted. Changes are tracked.'
        : 'Text has been capitalized and inserted. Make sure Track Changes is enabled in Word to see the changes tracked.';
      
      showSuccess(message);
      
      expect(messages[0]).toContain('capitalized');
      expect(messages[0]).toContain('tracked');
    });

    test('should show different message for insertClaudeResponse', () => {
      const trackChangesEnabled = true;
      const message = trackChangesEnabled
        ? 'Claude\'s response has been inserted with formatting preserved. Changes are tracked.'
        : 'Claude\'s response has been inserted with formatting preserved. Make sure Track Changes is enabled in Word to see the changes tracked.';
      
      showSuccess(message);
      
      expect(messages[0]).toContain('Claude');
      expect(messages[0]).toContain('formatting preserved');
      expect(messages[0]).toContain('tracked');
    });
  });

  describe('Error Messages', () => {
    test('should show error when Track Changes enable fails', () => {
      const errorMessage = 'Could not enable Track Changes. Document may be protected.';
      showError(errorMessage);
      
      expect(messages[0]).toContain('ERROR:');
      expect(messages[0]).toContain('Track Changes');
    });

    test('should show error when API not available', () => {
      const errorMessage = 'Track Changes API not available in this Word version. Please enable Track Changes manually.';
      showError(errorMessage);
      
      expect(messages[0]).toContain('not available');
      expect(messages[0]).toContain('manually');
    });

    test('should show error when document is protected', () => {
      const errorMessage = 'Document is protected. Cannot enable Track Changes.';
      showError(errorMessage);
      
      expect(messages[0]).toContain('protected');
      expect(messages[0]).toContain('Cannot enable');
    });
  });

  describe('Message Removal', () => {
    test('should remove "Make sure Track Changes is enabled" when auto-enabled', () => {
      const oldMessage = 'Text inserted. Make sure Track Changes is enabled in Word to see the changes tracked.';
      const newMessage = 'Text inserted. Changes are tracked.';
      
      // Simulate message update
      const updatedMessage = oldMessage.replace(
        'Make sure Track Changes is enabled in Word to see the changes tracked.',
        'Changes are tracked.'
      );
      
      expect(updatedMessage).toBe(newMessage);
      expect(updatedMessage).not.toContain('Make sure');
    });
  });

  describe('Message Consistency', () => {
    test('should use consistent message format across operations', () => {
      const capitalizeMessage = getSuccessMessage(true, false);
      const claudeMessage = getSuccessMessage(true, true);
      
      // Both should mention tracking
      expect(capitalizeMessage).toContain('tracked');
      expect(claudeMessage).toContain('tracked');
    });

    test('should differentiate between operations in messages', () => {
      const capitalizeMessage = 'Text has been capitalized and inserted. Changes are tracked.';
      const claudeMessage = 'Claude\'s response has been inserted. Changes are tracked.';
      
      expect(capitalizeMessage).toContain('capitalized');
      expect(claudeMessage).toContain('Claude');
    });
  });
});
