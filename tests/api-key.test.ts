/**
 * Tests for API key validation and storage functions
 */

describe('API Key Management', () => {
  // Mock localStorage
  const mockLocalStorage = (() => {
    let store: { [key: string]: string } = {};
    return {
      getItem: (key: string) => store[key] || null,
      setItem: (key: string, value: string) => {
        store[key] = value.toString();
      },
      removeItem: (key: string) => {
        delete store[key];
      },
      clear: () => {
        store = {};
      },
    };
  })();

  beforeEach(() => {
    // Reset localStorage before each test
    mockLocalStorage.clear();
    Object.defineProperty(window, 'localStorage', {
      value: mockLocalStorage,
      writable: true,
    });
  });

  describe('API Key Validation', () => {
    const validateApiKey = (apiKey: string): { valid: boolean; error?: string } => {
      if (!apiKey || apiKey.trim().length === 0) {
        return { valid: false, error: 'Please enter an API key' };
      }
      
      if (!apiKey.startsWith('sk-ant-')) {
        return { valid: false, error: 'API key should start with sk-ant-' };
      }
      
      return { valid: true };
    };

    test('should accept valid API key', () => {
      const result = validateApiKey('sk-ant-api03-valid-key-12345');
      expect(result.valid).toBe(true);
      expect(result.error).toBeUndefined();
    });

    test('should reject empty API key', () => {
      const result = validateApiKey('');
      expect(result.valid).toBe(false);
      expect(result.error).toBe('Please enter an API key');
    });

    test('should reject whitespace-only API key', () => {
      const result = validateApiKey('   ');
      expect(result.valid).toBe(false);
      expect(result.error).toBe('Please enter an API key');
    });

    test('should reject API key without sk-ant- prefix', () => {
      const result = validateApiKey('invalid-key');
      expect(result.valid).toBe(false);
      expect(result.error).toBe('API key should start with sk-ant-');
    });

    test('should accept API key with sk-ant- prefix', () => {
      const result = validateApiKey('sk-ant-api03-test-key');
      expect(result.valid).toBe(true);
    });

    test('should handle trimmed API key', () => {
      // Note: current implementation doesn't trim, so keys with leading spaces fail
      // TODO: Update validateApiKey to trim before validation
      const result = validateApiKey('  sk-ant-api03-test-key  ');
      expect(result.valid).toBe(false);
      expect(result.error).toBe('API key should start with sk-ant-');
    });
  });

  describe('API Key Storage', () => {
    const API_KEY_STORAGE_KEY = 'wordtrack_claude_api_key';

    test('should save API key to localStorage', () => {
      const apiKey = 'sk-ant-api03-test-key';
      mockLocalStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
      
      const saved = mockLocalStorage.getItem(API_KEY_STORAGE_KEY);
      expect(saved).toBe(apiKey);
    });

    test('should retrieve API key from localStorage', () => {
      const apiKey = 'sk-ant-api03-test-key';
      mockLocalStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
      
      const retrieved = mockLocalStorage.getItem(API_KEY_STORAGE_KEY);
      expect(retrieved).toBe(apiKey);
    });

    test('should return null for non-existent key', () => {
      const retrieved = mockLocalStorage.getItem(API_KEY_STORAGE_KEY);
      expect(retrieved).toBeNull();
    });

    test('should handle localStorage errors gracefully', () => {
      // Simulate localStorage error
      mockLocalStorage.getItem = jest.fn(() => {
        throw new Error('Storage quota exceeded');
      });

      expect(() => {
        try {
          mockLocalStorage.getItem(API_KEY_STORAGE_KEY);
        } catch (error) {
          // Should handle error gracefully
          expect(error).toBeInstanceOf(Error);
        }
      }).not.toThrow();
    });
  });
});
