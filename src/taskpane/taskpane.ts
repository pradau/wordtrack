import './taskpane.css';

let currentSelectedText: string = '';

function initializeTaskPane(): void {
  console.log('Initializing task pane...');
  
  const getSelectedTextButton = document.getElementById('get-selected-text-button');
  const capitalizeButton = document.getElementById('capitalize-button');
  
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

function hideMessages(): void {
  const container = document.getElementById('selected-text-container');
  const errorMessage = document.getElementById('error-message');
  const successMessage = document.getElementById('success-message');
  
  if (container) {
    container.style.display = 'none';
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

