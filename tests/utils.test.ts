/**
 * Tests for utility functions that can be tested in isolation
 */

// Import functions we want to test
// Since these are not exported, we'll need to extract them or test them indirectly
// For now, let's create testable versions

describe('Utility Functions', () => {
  describe('capitalizeWords', () => {
    // Function from taskpane.ts:396
    const capitalizeWords = (text: string): string => {
      return text.split(/\s+/).map(word => {
        if (word.length === 0) return word;
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      }).join(' ');
    };

    test('should capitalize first letter of each word', () => {
      expect(capitalizeWords('hello world')).toBe('Hello World');
    });

    test('should handle single word', () => {
      expect(capitalizeWords('hello')).toBe('Hello');
    });

    test('should handle empty string', () => {
      expect(capitalizeWords('')).toBe('');
    });

    test('should handle multiple spaces', () => {
      expect(capitalizeWords('hello   world')).toBe('Hello World');
    });

    test('should handle already capitalized text', () => {
      expect(capitalizeWords('HELLO WORLD')).toBe('Hello World');
    });

    test('should handle mixed case', () => {
      expect(capitalizeWords('hElLo WoRlD')).toBe('Hello World');
    });

    test('should handle text with numbers', () => {
      expect(capitalizeWords('hello 123 world')).toBe('Hello 123 World');
    });

    test('should handle text with punctuation', () => {
      expect(capitalizeWords('hello, world!')).toBe('Hello, World!');
    });

    test('should handle leading/trailing spaces', () => {
      // Note: capitalizeWords preserves leading/trailing spaces from split/join
      expect(capitalizeWords('  hello world  ')).toBe(' Hello World ');
    });
  });

  describe('escapeHtml', () => {
    // Function from taskpane.ts:465
    const escapeHtml = (text: string): string => {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    };

    test('should escape HTML entities', () => {
      expect(escapeHtml('<div>test</div>')).toBe('&lt;div&gt;test&lt;/div&gt;');
    });

    test('should escape ampersands', () => {
      expect(escapeHtml('A & B')).toBe('A &amp; B');
    });

    test('should handle quotes', () => {
      // Note: textContent doesn't escape quotes, only HTML tags and ampersands
      expect(escapeHtml('"quoted"')).toBe('"quoted"');
    });

    test('should handle plain text', () => {
      expect(escapeHtml('plain text')).toBe('plain text');
    });

    test('should handle empty string', () => {
      expect(escapeHtml('')).toBe('');
    });

    test('should escape multiple special characters', () => {
      const result = escapeHtml('<script>alert("xss")</script>');
      expect(result).not.toContain('<script>');
      // Note: textContent doesn't escape quotes, only HTML tags
      expect(result).toContain('"');
      expect(result).toContain('&lt;');
      expect(result).toContain('&gt;');
    });
  });
});
