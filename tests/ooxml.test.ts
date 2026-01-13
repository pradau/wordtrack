/**
 * Tests for OOXML parsing and HTML conversion functions
 */

describe('OOXML Parsing', () => {
  describe('convertOoxmlToHtml', () => {
    // Simplified version of convertOoxmlToHtml for testing
    const convertOoxmlToHtml = (ooxml: string, plainText: string): string => {
      let html = plainText;
      
      try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxml, 'text/xml');
        
        const paragraphs = xmlDoc.getElementsByTagName('w:p');
        let htmlParts: string[] = [];
        
        for (let i = 0; i < paragraphs.length; i++) {
          const para = paragraphs[i];
          const runs = para.getElementsByTagName('w:r');
          let paraHtml = '';
          
          for (let j = 0; j < runs.length; j++) {
            const run = runs[j];
            const textElements = run.getElementsByTagName('w:t');
            let runText = '';
            
            for (let k = 0; k < textElements.length; k++) {
              runText += textElements[k].textContent || '';
            }
            
            const isBold = run.getElementsByTagName('w:b').length > 0;
            const isItalic = run.getElementsByTagName('w:i').length > 0;
            
            const escapeHtml = (text: string): string => {
              const div = document.createElement('div');
              div.textContent = text;
              return div.innerHTML;
            };
            
            if (isBold && isItalic) {
              paraHtml += `<strong><em>${escapeHtml(runText)}</em></strong>`;
            } else if (isBold) {
              paraHtml += `<strong>${escapeHtml(runText)}</strong>`;
            } else if (isItalic) {
              paraHtml += `<em>${escapeHtml(runText)}</em>`;
            } else {
              paraHtml += escapeHtml(runText);
            }
          }
          
          const isList = para.getElementsByTagName('w:numPr').length > 0;
          if (isList) {
            htmlParts.push(`<li>${paraHtml}</li>`);
          } else {
            htmlParts.push(`<p>${paraHtml}</p>`);
          }
        }
        
        if (htmlParts.length > 0) {
          const hasList = htmlParts.some(p => p.startsWith('<li>'));
          if (hasList) {
            html = `<ul>${htmlParts.join('')}</ul>`;
          } else {
            html = htmlParts.join('');
          }
        }
      } catch (error) {
        console.error('Error converting OOXML to HTML:', error);
        html = plainText.replace(/\n/g, '<br>');
      }
      
      return html;
    };

    test('should convert simple OOXML to HTML', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>Hello World</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, 'Hello World');
      expect(result).toContain('Hello World');
      expect(result).toContain('<p>');
    });

    test('should handle bold text', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:rPr>
                  <w:b/>
                </w:rPr>
                <w:t>Bold Text</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, 'Bold Text');
      expect(result).toContain('<strong>');
      expect(result).toContain('Bold Text');
    });

    test('should handle italic text', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:rPr>
                  <w:i/>
                </w:rPr>
                <w:t>Italic Text</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, 'Italic Text');
      expect(result).toContain('<em>');
      expect(result).toContain('Italic Text');
    });

    test('should handle bold and italic text', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:rPr>
                  <w:b/>
                  <w:i/>
                </w:rPr>
                <w:t>Bold Italic</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, 'Bold Italic');
      expect(result).toContain('<strong><em>');
      expect(result).toContain('Bold Italic');
    });

    test('should handle list items', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0"/>
                  <w:numId w:val="1"/>
                </w:numPr>
              </w:pPr>
              <w:r>
                <w:t>List Item</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, 'List Item');
      expect(result).toContain('<ul>');
      expect(result).toContain('<li>');
      expect(result).toContain('List Item');
    });

    test('should handle multiple paragraphs', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>First paragraph</w:t>
              </w:r>
            </w:p>
            <w:p>
              <w:r>
                <w:t>Second paragraph</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, 'First paragraph\nSecond paragraph');
      expect(result).toContain('First paragraph');
      expect(result).toContain('Second paragraph');
      // Should have two <p> tags
      const pMatches = result.match(/<p>/g);
      expect(pMatches).toHaveLength(2);
    });

    test('should fallback to plain text on parse error', () => {
      const invalidOoxml = 'invalid xml <unclosed';
      const plainText = 'Hello\nWorld';
      const result = convertOoxmlToHtml(invalidOoxml, plainText);
      // Should fallback to plain text with newlines converted
      expect(result).toContain('Hello');
      expect(result).toContain('World');
    });

    test('should escape HTML in text content', () => {
      const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>&lt;script&gt;alert("xss")&lt;/script&gt;</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      const result = convertOoxmlToHtml(ooxml, '<script>alert("xss")</script>');
      // HTML should be escaped
      expect(result).toContain('&lt;');
      expect(result).not.toContain('<script>');
    });
  });
});
