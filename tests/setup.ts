// Test setup file for Jest
// Mock Office.js global object
(global as any).Office = {
  onReady: jest.fn((callback) => {
    callback({ host: 'Word' });
  }),
  HostType: {
    Word: 'Word',
  },
};

// Helper to create a mock Word context with customizable properties
function createMockWordContext(options: {
  trackRevisions?: boolean;
  trackRevisionsAvailable?: boolean;
  documentProtected?: boolean;
} = {}) {
  const {
    trackRevisions = false,
    trackRevisionsAvailable = true,
    documentProtected = false,
  } = options;

  const mockDocument: any = {
    getSelection: jest.fn(() => ({
      getRange: jest.fn(() => ({
        load: jest.fn(),
        text: '',
        font: {
          load: jest.fn(),
          name: '',
          size: 0,
          bold: false,
          italic: false,
          underline: '',
          color: '',
          highlightColor: '',
        },
        insertText: jest.fn(),
        getOoxml: jest.fn(() => ({
          value: '',
        })),
      })),
    })),
  };

  // Add trackRevisions property if available
  if (trackRevisionsAvailable) {
    Object.defineProperty(mockDocument, 'trackRevisions', {
      get: jest.fn(() => trackRevisions),
      set: jest.fn((value: boolean) => {
        if (documentProtected) {
          throw new Error('Document is protected');
        }
        // Update the value for subsequent gets
        Object.defineProperty(mockDocument, 'trackRevisions', {
          get: jest.fn(() => value),
          set: jest.fn(),
          configurable: true,
        });
      }),
      configurable: true,
    });
  }

  return {
    document: mockDocument,
    sync: jest.fn(() => Promise.resolve()),
  };
}

// Mock Word API
(global as any).Word = {
  run: jest.fn((callback) => {
    const mockContext = createMockWordContext();
    return callback(mockContext);
  }),
  InsertLocation: {
    replace: 'Replace',
  },
  UnderlineType: {
    none: 'none',
    single: 'single',
  },
};

// Export helper for use in tests
(global as any).createMockWordContext = createMockWordContext;

// Mock DOM APIs that might be needed
if (typeof DOMParser === 'undefined') {
  (global as any).DOMParser = class DOMParser {
    parseFromString(_str: string, _type: string): Document {
      // Basic mock - in real tests we'll use jsdom's DOMParser
      return {} as Document;
    }
  };
}
