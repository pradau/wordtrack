import './taskpane.css';

let currentSelectedText: string = '';
let currentClaudeResponse: string = '';

const API_KEY_STORAGE_KEY = 'wordtrack_claude_api_key';

function initializeTaskPane(): void {
  console.log('Initializing task pane...');
  
  loadApiKey();
  
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
  
  Word.run((context) => {
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    range.load('text');
    const ooxmlResult = range.getOoxml();
    
    return context.sync().then(() => {
      const selectedText = range.text.trim();
      
      if (selectedText.length === 0) {
        showError('No text is currently selected. Please select some text in your document and try again.');
        return;
      }
      
      currentSelectedText = selectedText;
      
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
  
  Word.run((context) => {
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    range.insertText(capitalizedText, Word.InsertLocation.replace);
    
    return context.sync().then(() => {
      showSuccess('Text has been capitalized and inserted. Make sure Track Changes is enabled in Word to see the changes tracked.');
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
  const fullPrompt = `${prompt}\n\nText to edit:\n${text}\n\nReturn only the edited text. Do not add titles, headers, explanations, or any other text before or after the edited content.`;
  
  const requestBody = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4096,
    messages: [
      {
        role: 'user',
        content: fullPrompt
      }
    ],
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
  
  Word.run((context) => {
    const selection = context.document.getSelection();
    const range = selection.getRange();
    
    range.insertText(currentClaudeResponse, Word.InsertLocation.replace);
    
    return context.sync().then(() => {
      showSuccess('Claude\'s response has been inserted. Make sure Track Changes is enabled in Word to see the changes tracked.');
      console.log('Claude response inserted successfully');
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

