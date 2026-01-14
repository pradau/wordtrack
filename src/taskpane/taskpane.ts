import './taskpane.css';
import guidelinesData from './guidelines.json';

let currentSelectedText: string = '';
let currentClaudeResponse: string = '';

type LogLevel = 'debug' | 'info' | 'warn' | 'error';
interface DebugLogEntry {
  ts: string;
  level: LogLevel;
  message: string;
}

const debugLogBuffer: DebugLogEntry[] = [];
const DEBUG_LOG_MAX_ENTRIES = 500;

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

function appendDebugLog(level: LogLevel, message: string): void {
  const entry: DebugLogEntry = {
    ts: new Date().toISOString(),
    level,
    message
  };
  debugLogBuffer.push(entry);
  if (debugLogBuffer.length > DEBUG_LOG_MAX_ENTRIES) {
    debugLogBuffer.splice(0, debugLogBuffer.length - DEBUG_LOG_MAX_ENTRIES);
  }
}

function safeStringify(value: unknown): string {
  try {
    if (typeof value === 'string') {
      return value;
    }
    if (value instanceof Error) {
      return `${value.name}: ${value.message}`;
    }
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function installConsoleLogCapture(): void {
  const original = {
    log: console.log.bind(console),
    info: console.info.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console)
  };

  console.log = (...args: unknown[]) => {
    appendDebugLog('info', args.map(safeStringify).join(' '));
    original.log(...args);
  };
  console.info = (...args: unknown[]) => {
    appendDebugLog('info', args.map(safeStringify).join(' '));
    original.info(...args);
  };
  console.warn = (...args: unknown[]) => {
    appendDebugLog('warn', args.map(safeStringify).join(' '));
    original.warn(...args);
  };
  console.error = (...args: unknown[]) => {
    appendDebugLog('error', args.map(safeStringify).join(' '));
    original.error(...args);
  };

  window.addEventListener('error', (event) => {
    appendDebugLog('error', `Window error: ${event.message} (${event.filename}:${event.lineno}:${event.colno})`);
  });

  window.addEventListener('unhandledrejection', (event) => {
    appendDebugLog('error', `Unhandled promise rejection: ${safeStringify(event.reason)}`);
  });
}

function buildDebugLogText(): string {
  const lines: string[] = [];
  lines.push('WordTrack Debug Log');
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push('');
  lines.push('Environment');
  lines.push(`User agent: ${navigator.userAgent}`);
  lines.push(`Office available: ${typeof Office !== 'undefined'}`);
  lines.push(`Word available: ${typeof Word !== 'undefined'}`);
  try {
    if (typeof Office !== 'undefined' && Office.context) {
      lines.push(`Host: ${String(Office.context.host)}`);
      lines.push(`Platform: ${String(Office.context.platform)}`);
    }
  } catch (e) {
    lines.push(`Office context read failed: ${safeStringify(e)}`);
  }
  lines.push('');
  lines.push(`Recent logs (last ${debugLogBuffer.length} entries; max ${DEBUG_LOG_MAX_ENTRIES})`);
  for (const entry of debugLogBuffer) {
    lines.push(`[${entry.ts}] ${entry.level.toUpperCase()}: ${entry.message}`);
  }
  lines.push('');
  return lines.join('\n');
}

function downloadDebugLog(): void {
  const text = buildDebugLogText();
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `wordtrack-debug-log-${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showSuccess('Debug log downloaded.');
}

function initializeTaskPane(): void {
  installConsoleLogCapture();
  console.log('Initializing task pane...');
  console.log('Office available:', typeof Office !== 'undefined');
  console.log('Word available:', typeof Word !== 'undefined');
  
  loadApiKey();
  
  const getSelectedTextButton = document.getElementById('get-selected-text-button');
  const capitalizeButton = document.getElementById('capitalize-button');
  const saveApiKeyButton = document.getElementById('save-api-key-button');
  const sendToClaudeButton = document.getElementById('send-to-claude-button');
  const insertClaudeResponseButton = document.getElementById('insert-claude-response-button');
  const downloadDebugLogButton = document.getElementById('download-debug-log-button');
  const helpButton = document.getElementById('help-button');
  const helpModal = document.getElementById('help-modal');
  const helpCloseButton = document.getElementById('help-close-button');
  
  if (getSelectedTextButton) {
    getSelectedTextButton.addEventListener('click', handleGetSelectedText);
    console.log('Get selected text button found and event listener added');
  } else {
    console.error('Button with id "get-selected-text-button" not found');
    setTimeout(initializeTaskPane, 100);
  }
  
  if (helpButton) {
    helpButton.addEventListener('click', () => {
      if (helpModal) {
        helpModal.style.display = 'block';
      }
    });
  }
  
  if (helpCloseButton && helpModal) {
    helpCloseButton.addEventListener('click', () => {
      helpModal.style.display = 'none';
    });
    // Close modal when clicking outside
    helpModal.addEventListener('click', (e) => {
      if (e.target === helpModal) {
        helpModal.style.display = 'none';
      }
    });
  }
  
  if (capitalizeButton) {
    capitalizeButton.addEventListener('click', handleCapitalizeAndInsert);
    console.log('Capitalize button found and event listener added');
  }
  
  if (saveApiKeyButton) {
    saveApiKeyButton.addEventListener('click', handleSaveApiKey);
  }
  if (downloadDebugLogButton) {
    downloadDebugLogButton.addEventListener('click', downloadDebugLog);
  }
  
  if (sendToClaudeButton) {
    sendToClaudeButton.addEventListener('click', handleSendToClaude);
  }
  
  if (insertClaudeResponseButton) {
    insertClaudeResponseButton.addEventListener('click', handleInsertClaudeResponse);
  }
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
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    range.insertText(capitalizedText, Word.InsertLocation.replace);
    
    return context.sync().then(async () => {
      showSuccess('Text has been capitalized and inserted.');
      console.log('Text replaced successfully');
    });
  }).catch((error) => {
    console.error('Error inserting text:', error);
    showError('Error inserting text: ' + error.message);
  });
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
  const promptKey = promptSelect ? promptSelect.value.trim() : 'clarity';
  
  // Map user-friendly prompt keys to token-efficient LLM prompts
  const promptMap: Record<string, string> = {
    'concise': 'Make concise',
    'expand': 'Expand 2x',
    'formal': 'Make formal',
    'casual': 'Make casual',
    'clarity': 'Improve clarity',
    'simplify': 'Simplify jargon',
    'professional': 'Make professional',
    'persuasive': 'Make persuasive',
    'structure': 'Improve structure',
    'detail': 'Add detail'
  };
  
  const customPrompt = promptMap[promptKey] || promptMap['clarity'];
  
  const contextSelect = document.getElementById('context-select') as HTMLSelectElement;
  const context = contextSelect ? contextSelect.value : 'personal';
  
  console.log('Sending to Claude...');
  
  hideMessages();
  showLoading(true);
  
  callClaudeAPI(currentSelectedText, customPrompt, context, apiKey)
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

async function callClaudeAPI(text: string, prompt: string, context: string, apiKey: string): Promise<string> {
  // Condense guidelines into compact system message
  const guidelinesText = guidelinesData.guidelines.join(' ');
  
  // Map context to minimal description
  const contextMap: Record<string, string> = {
    'personal': 'Personal: friends/colleagues',
    'formal': 'Formal: strangers/vendors',
    'documentation': 'Documentation: ethics/grants'
  };
  const contextDesc = contextMap[context] || contextMap['personal'];
  
  // Build messages array - system messages are token-efficient
  const messages: Array<{role: string, content: string}> = [];
  
  // System message: compact format
  messages.push({
    role: 'system',
    content: `${guidelinesText} Context: ${contextDesc}. Apply unless user instruction conflicts.`
  });
  
  // User message: minimal structure
  const userPrompt = `${prompt}\n\n${text}`;
  
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
        showSuccess('Claude\'s response has been inserted with formatting preserved.');
        console.log('Claude response inserted successfully with formatting');
      } else {
        showSuccess('Claude\'s response has been inserted.');
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

