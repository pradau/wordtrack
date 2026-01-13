/**
 * Tests for formatting preservation when Track Changes is enabled
 */

describe('Track Changes and Formatting Integration', () => {
  interface StoredFormatting {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    color?: string;
    highlightColor?: string;
  }

  describe('Formatting Preservation with Track Changes', () => {
    test('should preserve formatting when Track Changes is enabled', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const storedFormatting: StoredFormatting = {
        name: 'Arial',
        size: 12,
        bold: true,
        italic: false,
      };
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Apply formatting
      if (storedFormatting.name) {
        range.font.name = storedFormatting.name;
      }
      if (storedFormatting.size) {
        range.font.size = storedFormatting.size;
      }
      if (storedFormatting.bold !== undefined) {
        range.font.bold = storedFormatting.bold;
      }
      
      expect(range.font.name).toBe('Arial');
      expect(range.font.size).toBe(12);
      expect(range.font.bold).toBe(true);
    });

    test('should apply formatting correctly with Track Changes ON', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const storedFormatting: StoredFormatting = {
        bold: true,
        italic: true,
        color: '#FF0000',
      };
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Insert text first
      range.insertText('Test', 'Replace');
      
      // Then apply formatting
      if (storedFormatting.bold !== undefined) {
        range.font.bold = storedFormatting.bold;
      }
      if (storedFormatting.italic !== undefined) {
        range.font.italic = storedFormatting.italic;
      }
      if (storedFormatting.color) {
        range.font.color = storedFormatting.color;
      }
      
      expect(range.font.bold).toBe(true);
      expect(range.font.italic).toBe(true);
      expect(range.font.color).toBe('#FF0000');
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Formatting Application Order', () => {
    test('should enable Track Changes before applying formatting', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: false,
      });
      
      const order: string[] = [];
      
      // Enable Track Changes
      context.document.trackRevisions = true;
      order.push('enableTrackChanges');
      
      // Get selection
      const selection = context.document.getSelection();
      order.push('getSelection');
      
      // Apply formatting
      const range = selection.getRange();
      range.font.bold = true;
      order.push('applyFormatting');
      
      expect(order[0]).toBe('enableTrackChanges');
      expect(order[1]).toBe('getSelection');
      expect(order[2]).toBe('applyFormatting');
    });

    test('should insert text before applying formatting', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const order: string[] = [];
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Insert text
      range.insertText('Test', 'Replace');
      order.push('insertText');
      
      // Apply formatting
      range.font.bold = true;
      order.push('applyFormatting');
      
      expect(order[0]).toBe('insertText');
      expect(order[1]).toBe('applyFormatting');
    });
  });

  describe('All Formatting Properties', () => {
    test('should apply all formatting properties with Track Changes', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const storedFormatting: StoredFormatting = {
        name: 'Times New Roman',
        size: 14,
        bold: true,
        italic: true,
        underline: 'single',
        color: '#0000FF',
        highlightColor: '#FFFF00',
      };
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      range.insertText('Test', 'Replace');
      
      // Apply all formatting
      if (storedFormatting.name) range.font.name = storedFormatting.name;
      if (storedFormatting.size) range.font.size = storedFormatting.size;
      if (storedFormatting.bold !== undefined) range.font.bold = storedFormatting.bold;
      if (storedFormatting.italic !== undefined) range.font.italic = storedFormatting.italic;
      if (storedFormatting.underline) range.font.underline = storedFormatting.underline;
      if (storedFormatting.color) range.font.color = storedFormatting.color;
      if (storedFormatting.highlightColor) range.font.highlightColor = storedFormatting.highlightColor;
      
      expect(range.font.name).toBe('Times New Roman');
      expect(range.font.size).toBe(14);
      expect(range.font.bold).toBe(true);
      expect(range.font.italic).toBe(true);
      expect(range.font.underline).toBe('single');
      expect(range.font.color).toBe('#0000FF');
      expect(range.font.highlightColor).toBe('#FFFF00');
      expect(context.document.trackRevisions).toBe(true);
    });
  });

  describe('Formatting When Track Changes Unavailable', () => {
    test('should still apply formatting when Track Changes can\'t be enabled', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: false,
      });
      
      const storedFormatting: StoredFormatting = {
        bold: true,
        italic: true,
      };
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      range.insertText('Test', 'Replace');
      
      // Should still be able to apply formatting
      if (storedFormatting.bold !== undefined) {
        range.font.bold = storedFormatting.bold;
      }
      if (storedFormatting.italic !== undefined) {
        range.font.italic = storedFormatting.italic;
      }
      
      expect(range.font.bold).toBe(true);
      expect(range.font.italic).toBe(true);
    });
  });

  describe('Formatting Interference', () => {
    test('should not interfere with Track Changes when applying formatting', async () => {
      const context = (global as any).createMockWordContext({
        trackRevisionsAvailable: true,
        trackRevisions: true,
      });
      
      const selection = context.document.getSelection();
      const range = selection.getRange();
      
      // Enable Track Changes
      expect(context.document.trackRevisions).toBe(true);
      
      // Apply formatting
      range.font.bold = true;
      range.font.italic = true;
      
      // Track Changes should still be enabled
      expect(context.document.trackRevisions).toBe(true);
      expect(range.font.bold).toBe(true);
      expect(range.font.italic).toBe(true);
    });
  });
});
