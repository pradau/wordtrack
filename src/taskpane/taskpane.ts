import './taskpane.css';
import guidelinesData from './guidelines.json';

let currentSelectedText: string = '';
let currentClaudeResponse: string = '';
// Store formatting properties as plain values (not Office.js objects)
interface StoredFormatting {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: Word.UnderlineType | string;
  color?: string;
  highlightColor?: Word.UnderlineType | string;
}
let storedFormatting: StoredFormatting | null = null;

const API_KEY_STORAGE_KEY = 'wordtrack_claude_api_key';

function initializeTaskPane(): void {
  console.log('Initializing task pane...');
  console.log('Office available:', typeof Office !== 'undefined');
  console.log('Word available:', typeof Word !== 'undefined');
  
  loadApiKey();
  
  // Enable Track Changes automatically when add-in loads
  // Wait a bit to ensure Office.js is fully initialized
  console.log('Scheduling Track Changes enable on startup...');
  enableTrackChangesOnStartup();
  
  const getSelectedTextButton = document.getElementById('get-selected-text-button');
  const capitalizeButton = document.getElementById('capitalize-button');
  const saveApiKeyButton = document.getElementById('save-api-key-button');
  const sendToClaudeButton = document.getElementById('send-to-claude-button');
  const insertClaudeResponseButton = document.getElementById('insert-claude-response-button');
  
  if (getSelectedTextButton) {
    getSelectedTextButton.addEventListener('click', handleGetSelectedText);
    console.log('Get selected text button found and event listener added');
  } else {
    console.error('Button with id "get-selected-text-button" not found');
    setTimeout(initializeTaskPane, 100);
  }
  
  if (capitalizeButton) {
    capitalizeButton.addEventListener('click', handleCapitalizeAndInsert);
    console.log('Capitalize button found and event listener added');
  }
  
  if (saveApiKeyButton) {
    saveApiKeyButton.addEventListener('click', handleSaveApiKey);
  }
  
  const testTrackChangesButton = document.getElementById('test-track-changes-button');
  if (testTrackChangesButton) {
    testTrackChangesButton.addEventListener('click', handleTestTrackChanges);
  }
  
  if (sendToClaudeButton) {
    sendToClaudeButton.addEventListener('click', handleSendToClaude);
  }
  
  if (insertClaudeResponseButton) {
    insertClaudeResponseButton.addEventListener('click', handleInsertClaudeResponse);
  }
}

/**
 * Test function to manually enable Track Changes (for debugging)
 */
function handleTestTrackChanges(): void {
  console.log('Test Track Changes button clicked');
  updateApiKeyStatus('Testing Track Changes...', false);
  
  // Check if Word API is available
  if (typeof Word === 'undefined') {
    console.error('Test: Word API not available');
    updateApiKeyStatus('Word API not available', true);
    showError('Word API is not available. Make sure the add-in is loaded in Word.');
    return;
  }
  
  console.log('Test: Word API is available');
  console.log('Test: Office.context available:', typeof Office !== 'undefined' && typeof Office.context !== 'undefined');
  
  // Check requirement set
  if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
    const hasWordApi14 = Office.context.requirements.isSetSupported('WordApi', '1.4');
    console.log('Test: WordApi 1.4 supported:', hasWordApi14);
    if (!hasWordApi14) {
      updateApiKeyStatus('WordApi 1.4 not supported', true);
      showError('Your Word version does not support Track Changes API (needs WordApi 1.4+). Please enable Track Changes manually in the Review tab.');
      return;
    }
  }
  
  Word.run(async (context) => {
    console.log('Test: Word.run() started');
    console.log('Test: context.document available:', typeof context.document !== 'undefined');
    console.log('Test: trackRevisions in document:', 'trackRevisions' in context.document);
    
    try {
      const enabled = await ensureTrackChangesEnabled(context);
      console.log('Test: ensureTrackChangesEnabled returned:', enabled);
      
      if (enabled) {
        // Verify it's actually enabled
        context.document.load('trackRevisions');
        await context.sync();
        const actualState = context.document.trackRevisions;
        console.log('Test: Actual trackRevisions state after enable:', actualState);
        
        if (actualState === true) {
          console.log('Test: Track Changes enabled successfully');
          updateApiKeyStatus('Track Changes enabled! Check Review tab.', false);
          showSuccess('Track Changes has been enabled. Check the Review tab to verify.');
        } else {
          console.log('Test: Track Changes enable failed - state is still false');
          updateApiKeyStatus('Track Changes: Enable failed (state still false)', true);
          showError('Track Changes could not be enabled. The state is still false after attempting to enable.');
        }
      } else {
        console.log('Test: Track Changes could not be enabled');
        updateApiKeyStatus('Track Changes: Could not enable (check console)', true);
        showError('Track Changes could not be enabled. Check browser console for details.');
      }
    } catch (innerError) {
      console.error('Test: Error in ensureTrackChangesEnabled:', innerError);
      if (innerError instanceof Error) {
        console.error('Test: Inner error name:', innerError.name);
        console.error('Test: Inner error message:', innerError.message);
      }
      updateApiKeyStatus('Track Changes: Error in enable function', true);
      showError('Error in enable function: ' + (innerError instanceof Error ? innerError.message : String(innerError)));
    }
  }).catch((error) => {
    console.error('Test: Error in Word.run():', error);
    if (error instanceof Error) {
      console.error('Test: Error name:', error.name);
      console.error('Test: Error message:', error.message);
      if ((error as any).code) {
        console.error('Test: Error code:', (error as any).code);
      }
      if ((error as any).debugInfo) {
        console.error('Test: Debug info:', (error as any).debugInfo);
      }
    }
    updateApiKeyStatus('Track Changes: Error in Word.run()', true);
    showError('Error: ' + (error instanceof Error ? error.message : String(error)));
  });
}

/**
 * Enable Track Changes when the add-in starts up.
 * This runs automatically when the task pane loads.
 */
function enableTrackChangesOnStartup(): void {
  console.log('enableTrackChangesOnStartup() called');
  
  // Wait for Office.js and Word API to be fully ready
  const checkAndEnable = () => {
    console.log('Checking if Word API is available...');
    console.log('Word API available:', typeof Word !== 'undefined');
    console.log('Office API available:', typeof Office !== 'undefined');
    
    if (typeof Word === 'undefined') {
      console.log('Word API not available yet, retrying in 500ms...');
      setTimeout(checkAndEnable, 500);
      return;
    }
    
    console.log('Word API is available, attempting to enable Track Changes...');
    
    Word.run(async (context) => {
      console.log('Word.run() started for Track Changes enable');
      const enabled = await ensureTrackChangesEnabled(context);
      if (enabled) {
        console.log('Track Changes enabled automatically on add-in startup');
        // Show status message
        updateApiKeyStatus('Track Changes enabled automatically', false);
      } else {
        console.log('Track Changes could not be enabled automatically (API may not be available)');
        updateApiKeyStatus('Track Changes: Enable manually in Review tab', false);
      }
    }).catch((error) => {
      console.error('Error enabling Track Changes on startup:', error);
      if (error instanceof Error) {
        console.error('Error name:', error.name);
        console.error('Error message:', error.message);
        if ((error as any).code) {
          console.error('Error code:', (error as any).code);
        }
      }
      updateApiKeyStatus('Track Changes: Enable manually in Review tab', false);
    });
  };
  
  // Start checking after a short delay
  setTimeout(checkAndEnable, 1000);
}

function loadApiKey(): void {
  try {
    const savedKey = localStorage.getItem(API_KEY_STORAGE_KEY);
    if (savedKey) {
      const apiKeyInput = document.getElementById('api-key-input') as HTMLInputElement;
      if (apiKeyInput) {
        apiKeyInput.value = savedKey;
        updateApiKeyStatus('API key loaded');
      }
    }
  } catch (error) {
    console.error('Error loading API key:', error);
  }
}

function handleSaveApiKey(): void {
  const apiKeyInput = document.getElementById('api-key-input') as HTMLInputElement;
  if (!apiKeyInput) {
    return;
  }
  
  const apiKey = apiKeyInput.value.trim();
  
  if (!apiKey) {
    updateApiKeyStatus('Please enter an API key', true);
    return;
  }
  
  if (!apiKey.startsWith('sk-ant-')) {
    updateApiKeyStatus('API key should start with sk-ant-', true);
    return;
  }
  
  try {
    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
    updateApiKeyStatus('API key saved');
    console.log('API key saved to localStorage');
  } catch (error) {
    console.error('Error saving API key:', error);
    updateApiKeyStatus('Error saving API key', true);
  }
}

function updateApiKeyStatus(message: string, isError: boolean = false): void {
  const statusElement = document.getElementById('api-key-status');
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.style.color = isError ? '#dc3545' : '#28a745';
    console.log('Status updated:', message);
  } else {
    console.warn('Status element not found');
  }
}

function getApiKey(): string | null {
  try {
    return localStorage.getItem(API_KEY_STORAGE_KEY);
  } catch (error) {
    console.error('Error getting API key:', error);
    return null;
  }
}

function handleGetSelectedText(): void {
  console.log('Get selected text button clicked');
  
  hideMessages();
  
  Word.run(async (context) => {
    // Ensure Track Changes is enabled (in case it was turned off after accepting changes)
    await ensureTrackChangesEnabled(context);
    
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    // Load text and formatting properties
    range.load('text');
    range.font.load('name, size, bold, italic, underline, color, highlightColor');
    const ooxmlResult = range.getOoxml();
    
    return context.sync().then(() => {
      const selectedText = range.text.trim();
      
      if (selectedText.length === 0) {
        showError('No text is currently selected. Please select some text in your document and try again.');
        return;
      }
      
      currentSelectedText = selectedText;
      
      // Use majority vote approach: analyze OOXML to count formatting from each run
      const formattingCounts: {
        name: Map<string, number>;
        size: Map<number, number>;
        bold: { true: number; false: number };
        italic: { true: number; false: number };
        underline: Map<string, number>;
        color: Map<string, number>;
        highlightColor: Map<string, number>;
      } = {
        name: new Map(),
        size: new Map(),
        bold: { true: 0, false: 0 },
        italic: { true: 0, false: 0 },
        underline: new Map(),
        color: new Map(),
        highlightColor: new Map()
      };
      
      let totalWords = 0;
      
      // Parse OOXML to count formatting from each run, weighted by word count
      try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(ooxmlResult.value, 'text/xml');
        const paragraphs = xmlDoc.getElementsByTagName('w:p');
        
        for (let i = 0; i < paragraphs.length; i++) {
          const para = paragraphs[i];
          const runs = para.getElementsByTagName('w:r');
          
          for (let j = 0; j < runs.length; j++) {
            const run = runs[j];
            const textElements = run.getElementsByTagName('w:t');
            let runText = '';
            
            for (let k = 0; k < textElements.length; k++) {
              runText += textElements[k].textContent || '';
            }
            
            // Only count runs with actual text
            if (runText.trim().length > 0) {
              // Count words in this run (split by whitespace, filter empty strings)
              const words = runText.trim().split(/\s+/).filter(word => word.length > 0);
              const wordCount = words.length;
              
              if (wordCount > 0) {
                totalWords += wordCount;
                
                // Get run properties (rPr = run properties)
                const rPr = run.getElementsByTagName('w:rPr')[0];
                
                if (rPr) {
                  // Count font name (weighted by word count)
                  const fontName = rPr.getElementsByTagName('w:rFonts')[0];
                  if (fontName) {
                    const name = fontName.getAttribute('w:ascii') || fontName.getAttribute('w:hAnsi') || '';
                    if (name) {
                      const count = formattingCounts.name.get(name) || 0;
                      formattingCounts.name.set(name, count + wordCount);
                    }
                  }
                  
                  // Count font size (weighted by word count)
                  const fontSize = rPr.getElementsByTagName('w:sz')[0];
                  if (fontSize) {
                    const sizeAttr = fontSize.getAttribute('w:val');
                    if (sizeAttr) {
                      const size = parseInt(sizeAttr) / 2; // Office.js uses half-points
                      const count = formattingCounts.size.get(size) || 0;
                      formattingCounts.size.set(size, count + wordCount);
                    }
                  }
                  
                  // Count bold (weighted by word count)
                  const isBold = rPr.getElementsByTagName('w:b').length > 0;
                  if (isBold) {
                    formattingCounts.bold.true += wordCount;
                  } else {
                    formattingCounts.bold.false += wordCount;
                  }
                  
                  // Count italic (weighted by word count)
                  const isItalic = rPr.getElementsByTagName('w:i').length > 0;
                  if (isItalic) {
                    formattingCounts.italic.true += wordCount;
                  } else {
                    formattingCounts.italic.false += wordCount;
                  }
                  
                  // Count underline (weighted by word count)
                  const underline = rPr.getElementsByTagName('w:u')[0];
                  if (underline) {
                    const underlineVal = underline.getAttribute('w:val') || 'single';
                    const count = formattingCounts.underline.get(underlineVal) || 0;
                    formattingCounts.underline.set(underlineVal, count + wordCount);
                  } else {
                    const count = formattingCounts.underline.get('none') || 0;
                    formattingCounts.underline.set('none', count + wordCount);
                  }
                  
                  // Count color (weighted by word count)
                  const color = rPr.getElementsByTagName('w:color')[0];
                  if (color) {
                    const colorVal = color.getAttribute('w:val') || '';
                    if (colorVal) {
                      const count = formattingCounts.color.get(colorVal) || 0;
                      formattingCounts.color.set(colorVal, count + wordCount);
                    }
                  }
                  
                  // Count highlight color (weighted by word count)
                  const highlight = rPr.getElementsByTagName('w:highlight')[0];
                  if (highlight) {
                    const highlightVal = highlight.getAttribute('w:val') || '';
                    if (highlightVal) {
                      const count = formattingCounts.highlightColor.get(highlightVal) || 0;
                      formattingCounts.highlightColor.set(highlightVal, count + wordCount);
                    }
                  }
                } else {
                  // No run properties - count as default formatting (weighted by word count)
                  formattingCounts.bold.false += wordCount;
                  formattingCounts.italic.false += wordCount;
                  const count = formattingCounts.underline.get('none') || 0;
                  formattingCounts.underline.set('none', count + wordCount);
                }
              }
            }
          }
        }
      } catch (error) {
        console.error('Error parsing OOXML for majority vote:', error);
        // Fall back to range-level formatting if OOXML parsing fails
        totalWords = 0;
      }
      
      // Determine majority formatting (based on word count, not run count)
      storedFormatting = {};
      
      if (totalWords > 0) {
        // Font name - most common (by word count)
        let maxCount = 0;
        let majorityName: string | undefined;
        formattingCounts.name.forEach((count, name) => {
          if (count > maxCount) {
            maxCount = count;
            majorityName = name;
          }
        });
        if (majorityName && maxCount > totalWords / 2) {
          storedFormatting.name = majorityName;
        }
        
        // Font size - most common (by word count)
        maxCount = 0;
        let majoritySize: number | undefined;
        formattingCounts.size.forEach((count, size) => {
          if (count > maxCount) {
            maxCount = count;
            majoritySize = size;
          }
        });
        if (majoritySize !== undefined && maxCount > totalWords / 2) {
          storedFormatting.size = majoritySize;
        }
        
        // Bold - majority vote (by word count)
        if (formattingCounts.bold.true + formattingCounts.bold.false > 0) {
          storedFormatting.bold = formattingCounts.bold.true > formattingCounts.bold.false;
        }
        
        // Italic - majority vote (by word count)
        if (formattingCounts.italic.true + formattingCounts.italic.false > 0) {
          storedFormatting.italic = formattingCounts.italic.true > formattingCounts.italic.false;
        }
        
        // Underline - most common (by word count, if majority)
        maxCount = 0;
        let majorityUnderline: string | undefined;
        formattingCounts.underline.forEach((count, underline) => {
          if (count > maxCount) {
            maxCount = count;
            majorityUnderline = underline;
          }
        });
        if (majorityUnderline && maxCount > totalWords / 2) {
          storedFormatting.underline = majorityUnderline as Word.UnderlineType;
        }
        
        // Color - most common (by word count, if majority)
        maxCount = 0;
        let majorityColor: string | undefined;
        formattingCounts.color.forEach((count, color) => {
          if (count > maxCount) {
            maxCount = count;
            majorityColor = color;
          }
        });
        if (majorityColor && maxCount > totalWords / 2) {
          storedFormatting.color = majorityColor;
        }
        
        // Highlight color - most common (by word count, if majority)
        maxCount = 0;
        let majorityHighlight: string | undefined;
        formattingCounts.highlightColor.forEach((count, highlight) => {
          if (count > maxCount) {
            maxCount = count;
            majorityHighlight = highlight;
          }
        });
        if (majorityHighlight && maxCount > totalWords / 2) {
          storedFormatting.highlightColor = majorityHighlight as Word.UnderlineType;
        }
      }
      
      const ooxml = ooxmlResult.value;
      const htmlContent = convertOoxmlToHtml(ooxml, selectedText);
      showSelectedText(selectedText, htmlContent);
    });
  }).catch((error) => {
    console.error('Error getting selected text:', error);
    showError('Error retrieving selected text: ' + error.message);
  });
}

function handleCapitalizeAndInsert(): void {
  if (!currentSelectedText) {
    showError('Please select text first using "Get Selected Text" button.');
    return;
  }
  
  console.log('Capitalize and insert button clicked');
  
  hideMessages();
  
  const capitalizedText = capitalizeWords(currentSelectedText);
  
  Word.run(async (context) => {
    // Enable Track Changes before inserting
    const trackChangesEnabled = await ensureTrackChangesEnabled(context);
    
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    range.insertText(capitalizedText, Word.InsertLocation.replace);
    
    return context.sync().then(async () => {
      // Verify Track Changes is still enabled after insertion
      // (Word might have turned it off, so re-enable if needed)
      if (trackChangesEnabled) {
        await ensureTrackChangesEnabled(context);
      }
      
      if (trackChangesEnabled) {
        showSuccess('Text has been capitalized and inserted. Changes are tracked.');
      } else {
        showSuccess('Text has been capitalized and inserted. Make sure Track Changes is enabled in Word to see the changes tracked.');
      }
      console.log('Text replaced successfully');
    });
  }).catch((error) => {
    console.error('Error inserting text:', error);
    showError('Error inserting text: ' + error.message);
  });
}

/**
 * Ensures Track Changes is enabled in the document.
 * Returns true if Track Changes was enabled (or already was enabled), false if not possible.
 */
async function ensureTrackChangesEnabled(context: Word.RequestContext): Promise<boolean> {
  console.log('ensureTrackChangesEnabled() called');
  try {
    // Check Word API requirement set first (but don't fail if check isn't available)
    if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
      try {
        const hasWordApi14 = Office.context.requirements.isSetSupported('WordApi', '1.4');
        console.log('WordApi 1.4 supported:', hasWordApi14);
        if (!hasWordApi14) {
          console.warn('WordApi 1.4 not supported - Track Changes API may be unavailable');
          // Don't return false here - try anyway in case the check is wrong
        }
      } catch (reqError) {
        console.warn('Could not check requirement set, proceeding anyway:', reqError);
      }
    }
    
    // Check if trackRevisions property exists (available in WordApi 1.4+)
    if ('trackRevisions' in context.document) {
      console.log('trackRevisions property exists in document');
      // Check current state
      context.document.load('trackRevisions');
      await context.sync();
      const currentState = context.document.trackRevisions;
      console.log('Current trackRevisions state:', currentState);
      
      // If already enabled, return true
      if (currentState === true) {
        console.log('Track Changes already enabled');
        return true;
      }
      
      // Enable Track Changes (it was OFF, so we need to turn it ON)
      console.log('Track Changes is OFF - enabling it...');
      context.document.trackRevisions = true;
      await context.sync();
      
      // Verify it was actually enabled
      context.document.load('trackRevisions');
      await context.sync();
      const newState = context.document.trackRevisions;
      console.log('Track Changes state after enabling:', newState);
      
      if (newState === true) {
        console.log('Track Changes enabled successfully');
        return true;
      } else {
        console.warn('Track Changes enable failed - state is still false');
        return false;
      }
    } else {
      // API not available (older Word version)
      console.warn('Track Changes API not available in this Word version (trackRevisions property not found)');
      return false;
    }
  } catch (error) {
    // Handle errors (e.g., document protected)
    console.error('Error enabling Track Changes:', error);
    if (error instanceof Error) {
      console.error('Error name:', error.name);
      console.error('Error message:', error.message);
      if ((error as any).code) {
        console.error('Error code:', (error as any).code);
      }
    }
    return false;
  }
}

function capitalizeWords(text: string): string {
  return text.split(/\s+/).map(word => {
    if (word.length === 0) return word;
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
}

function convertOoxmlToHtml(ooxml: string, plainText: string): string {
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
}

function escapeHtml(text: string): string {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function showSelectedText(text: string, htmlContent?: string): void {
  const container = document.getElementById('selected-text-container');
  const display = document.getElementById('selected-text-display');
  const errorMessage = document.getElementById('error-message');
  
  if (errorMessage) {
    errorMessage.style.display = 'none';
  }
  
  if (container && display) {
    if (htmlContent && htmlContent.trim().length > 0) {
      display.innerHTML = htmlContent;
    } else {
      display.textContent = text;
    }
    container.style.display = 'block';
    console.log('Selected text displayed, length:', text.length, 'has HTML:', !!htmlContent);
  }
}

function showError(message: string): void {
  const container = document.getElementById('selected-text-container');
  const errorMessage = document.getElementById('error-message');
  
  if (container) {
    container.style.display = 'none';
  }
  
  if (errorMessage) {
    errorMessage.style.display = 'block';
    errorMessage.style.cssText = 'display: block; padding: 15px; margin-top: 15px; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; color: #721c24;';
    errorMessage.textContent = message;
    console.error('Error message displayed:', message);
  }
}

function showSuccess(message: string): void {
  const successMessage = document.getElementById('success-message');
  const errorMessage = document.getElementById('error-message');
  
  if (errorMessage) {
    errorMessage.style.display = 'none';
  }
  
  if (successMessage) {
    successMessage.style.display = 'block';
    successMessage.style.cssText = 'display: block; padding: 15px; margin-top: 15px; background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 4px; color: #155724;';
    successMessage.textContent = message;
    console.log('Success message displayed:', message);
  }
}

function handleSendToClaude(): void {
  if (!currentSelectedText) {
    showError('Please select text first using "Get Selected Text" button.');
    return;
  }
  
  const apiKey = getApiKey();
  if (!apiKey) {
    showError('Please enter and save your Claude API key first.');
    return;
  }
  
  const promptSelect = document.getElementById('prompt-select') as HTMLSelectElement;
  const customPrompt = promptSelect ? promptSelect.value.trim() : 'Make this 2x longer with verbosity.';
  
  if (!customPrompt) {
    showError('Please select a prompt.');
    return;
  }
  
  console.log('Sending to Claude...');
  
  hideMessages();
  showLoading(true);
  
  callClaudeAPI(currentSelectedText, customPrompt, apiKey)
    .then((response) => {
      showLoading(false);
      currentClaudeResponse = response;
      showClaudeResponse(response);
    })
    .catch((error) => {
      showLoading(false);
      console.error('Error calling Claude API:', error);
      showError('Error calling Claude API: ' + error.message);
    });
}

async function callClaudeAPI(text: string, prompt: string, apiKey: string): Promise<string> {
  // Condense guidelines into a compact system message
  const guidelinesText = guidelinesData.guidelines.join(' ');
  
  // Build messages array with system message for guidelines (more token-efficient)
  // System messages are typically cheaper and don't need to be repeated
  const messages: Array<{role: string, content: string}> = [];
  
  // Add system message with condensed guidelines
  messages.push({
    role: 'system',
    content: `Text editing guidelines: ${guidelinesText}. Apply these unless the user's instruction conflicts.`
  });
  
  // User message is just the instruction and text (much shorter)
  const userPrompt = `${prompt}\n\nText:\n${text}\n\nReturn only the edited text.`;
  
  messages.push({
    role: 'user',
    content: userPrompt
  });
  
  const requestBody = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4096,
    messages: messages,
    apiKey: apiKey
  };
  
  let proxyUrl = 'https://localhost:3001/api/claude';
  
  try {
    const response = await fetch(proxyUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(requestBody),
      mode: 'cors'
    }).catch(async (httpsError) => {
      console.log('HTTPS proxy failed, trying HTTP:', httpsError);
      proxyUrl = 'http://localhost:3001/api/claude';
      return fetch(proxyUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody),
        mode: 'cors'
      });
    });
    
    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      const errorMessage = errorData.error?.message || response.statusText;
      
      if (response.status === 401) {
        throw new Error('Invalid API key. Please check your API key and try again.');
      } else if (response.status === 402 || errorMessage.includes('credit balance') || errorMessage.includes('too low')) {
        throw new Error('Insufficient credits. Please add credits to your Anthropic account at https://console.anthropic.com/');
      } else if (response.status === 429) {
        throw new Error('Rate limit exceeded. Please try again later.');
      } else if (response.status === 500) {
        throw new Error('Proxy server error. Make sure the proxy server is running (npm run proxy).');
      } else {
        throw new Error(`API error: ${errorMessage}`);
      }
    }
    
    const data = await response.json();
    
    if (data.content && data.content.length > 0 && data.content[0].text) {
      return data.content[0].text;
    } else {
      throw new Error('Unexpected response format from Claude API');
    }
  } catch (error) {
    if (error instanceof TypeError && error.message.includes('Failed to fetch')) {
      throw new Error('Cannot connect to proxy server. Make sure the proxy is running: npm run proxy');
    }
    throw error;
  }
}

function showClaudeResponse(response: string): void {
  const container = document.getElementById('claude-response-container');
  const display = document.getElementById('claude-response-display');
  
  if (container && display) {
    display.textContent = response;
    container.style.display = 'block';
    console.log('Claude response displayed, length:', response.length);
  }
}

function handleInsertClaudeResponse(): void {
  if (!currentClaudeResponse) {
    showError('No Claude response available. Please send text to Claude first.');
    return;
  }
  
  console.log('Inserting Claude response...');
  
  hideMessages();
  
  Word.run(async (context) => {
    // Enable Track Changes before inserting
    const trackChangesEnabled = await ensureTrackChangesEnabled(context);
    
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    // Insert the text
    range.insertText(currentClaudeResponse, Word.InsertLocation.replace);
    
    // Apply stored formatting if available
    if (storedFormatting && Object.keys(storedFormatting).length > 0) {
      // Apply formatting to the newly inserted text
      // Only apply properties that are defined and not null
      if (storedFormatting.name !== undefined && storedFormatting.name !== null) {
        range.font.name = storedFormatting.name;
      }
      if (storedFormatting.size !== undefined && storedFormatting.size !== null) {
        range.font.size = storedFormatting.size;
      }
      if (storedFormatting.bold !== undefined && storedFormatting.bold !== null) {
        range.font.bold = storedFormatting.bold;
      }
      if (storedFormatting.italic !== undefined && storedFormatting.italic !== null) {
        range.font.italic = storedFormatting.italic;
      }
      if (storedFormatting.underline !== undefined && storedFormatting.underline !== null) {
        range.font.underline = storedFormatting.underline as Word.UnderlineType;
      }
      if (storedFormatting.color !== undefined && storedFormatting.color !== null) {
        range.font.color = storedFormatting.color;
      }
      if (storedFormatting.highlightColor !== undefined && storedFormatting.highlightColor !== null) {
        range.font.highlightColor = storedFormatting.highlightColor as Word.UnderlineType;
      }
    }
    
    return context.sync().then(() => {
      if (storedFormatting) {
        if (trackChangesEnabled) {
          showSuccess('Claude\'s response has been inserted with formatting preserved. Changes are tracked.');
        } else {
          showSuccess('Claude\'s response has been inserted with formatting preserved. Make sure Track Changes is enabled in Word to see the changes tracked.');
        }
        console.log('Claude response inserted successfully with formatting');
      } else {
        if (trackChangesEnabled) {
          showSuccess('Claude\'s response has been inserted. Changes are tracked.');
        } else {
          showSuccess('Claude\'s response has been inserted. Make sure Track Changes is enabled in Word to see the changes tracked.');
        }
        console.log('Claude response inserted successfully');
      }
    });
  }).catch((error) => {
    console.error('Error inserting Claude response:', error);
    showError('Error inserting text: ' + error.message);
  });
}

function showLoading(show: boolean): void {
  const loadingIndicator = document.getElementById('loading-indicator');
  if (loadingIndicator) {
    loadingIndicator.style.display = show ? 'block' : 'none';
  }
}

function hideMessages(): void {
  const container = document.getElementById('selected-text-container');
  const claudeContainer = document.getElementById('claude-response-container');
  const errorMessage = document.getElementById('error-message');
  const successMessage = document.getElementById('success-message');
  
  if (container) {
    container.style.display = 'none';
  }
  
  if (claudeContainer) {
    claudeContainer.style.display = 'none';
  }
  
  if (errorMessage) {
    errorMessage.style.display = 'none';
  }
  
  if (successMessage) {
    successMessage.style.display = 'none';
  }
}

if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    console.log('Office.onReady called, host:', info.host);
    if (info.host === Office.HostType.Word) {
      console.log('Word host confirmed, initializing task pane...');
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeTaskPane);
      } else {
        initializeTaskPane();
      }
    } else {
      console.error('This add-in only works in Word. Current host:', info.host);
    }
  });
} else {
  console.error('Office.js not loaded');
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeTaskPane);
  } else {
    initializeTaskPane();
  }
}

