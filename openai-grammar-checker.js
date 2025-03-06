/**
 * Creates a custom menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Grammar Checker')
    .addItem('Configure API Key', 'showApiKeyDialog')
    .addItem('Fix Typos', 'showTypoFixDialog')
    .addItem('Improve Grammar', 'showGrammarImproveDialog')
    .addItem('Manage Custom Dictionary', 'showDictionaryManager')
    .addItem('Language Settings', 'showLanguageSettings')
    .addSeparator()
    .addItem('Configure Triggers', 'configTriggers')
    .addToUi();
}

/**
 * Runs when the spreadsheet is opened, can be set as a trigger
 */
function onOpenDocument() {
  // Check if API key is set
  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    showApiKeyDialog();
  }
}

/**
 * Shows a dialog to set the OpenAI API key
 */
function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'OpenAI API Key',
    'Please enter your OpenAI API key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText().trim();
    PropertiesService.getUserProperties().setProperty('OPENAI_API_KEY', apiKey);
    ui.alert('API key saved successfully!');
  }
}

/**
 * Shows the language settings dialog
 */
function showLanguageSettings() {
  const ui = SpreadsheetApp.getUi();
  const currentLanguage = PropertiesService.getUserProperties().getProperty('GRAMMAR_CHECKER_LANGUAGE') || 'english';
  
  const htmlOutput = HtmlService.createHtmlOutput(
    `<form>
      <label>Select Language:</label><br>
      <select id="language">
        <option value="english" ${currentLanguage === 'english' ? 'selected' : ''}>English</option>
        <option value="spanish" ${currentLanguage === 'spanish' ? 'selected' : ''}>Spanish</option>
      </select>
      <br><br>
      <input type="button" value="Save" onclick="saveLanguage()" />
    </form>
    
    <script>
      function saveLanguage() {
        const language = document.getElementById('language').value;
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .setLanguagePreference(language);
      }
    </script>`
  )
  .setWidth(300)
  .setHeight(200);
  
  ui.showModalDialog(htmlOutput, 'Language Settings');
}

/**
 * Saves the selected language preference
 */
function setLanguagePreference(language) {
  PropertiesService.getUserProperties().setProperty('GRAMMAR_CHECKER_LANGUAGE', language);
  const ui = SpreadsheetApp.getUi();
  ui.alert(`Language set to ${language.charAt(0).toUpperCase() + language.slice(1)}`);
}

/**
 * Shows the typo fix dialog
 */
function showTypoFixDialog() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  
  if (!range) {
    ui.alert('Please select a range first.');
    return;
  }
  
  const text = range.getValue().toString();
  if (!text) {
    ui.alert('Selected cell is empty.');
    return;
  }
  
  // Find potential typos
  findTypos(text, range);
}

/**
 * Finds typos in text using OpenAI API
 */
function findTypos(text, range) {
  const language = PropertiesService.getUserProperties().getProperty('GRAMMAR_CHECKER_LANGUAGE') || 'english';
  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Please set your OpenAI API key first.');
    showApiKeyDialog();
    return;
  }
  
  // Get custom dictionary words
  const customDict = getCustomDictionary();
  
  try {
    const languagePrompt = language === 'spanish' ? 
      'Identifica palabras con posibles errores ortográficos en el siguiente texto en español. Para cada palabra, proporciona posibles correcciones:' :
      'Identify words with potential spelling mistakes in the following English text. For each word, provide possible corrections:';
    
    const response = callOpenAI(apiKey, [
      {
        "role": "system", 
        "content": `You are a spelling assistant. ${languagePrompt} Return your response as a JSON object with the structure: {"typos":[{"word":"misspelled", "replacements":["correct1", "correct2"]}]}`
      },
      {"role": "user", "content": text}
    ]);
    
    if (!response) {
      throw new Error("Failed to get response from OpenAI");
    }
    
    let result;
    try {
      // Extract JSON from the response
      const jsonMatch = response.match(/```json\s*([\s\S]*?)\s*```/) || 
                      response.match(/{[\s\S]*}/);
      
      if (jsonMatch) {
        result = JSON.parse(jsonMatch[0].replace(/```json|```/g, '').trim());
      } else {
        result = JSON.parse(response);
      }
    } catch (e) {
      console.error("Error parsing JSON:", e);
      console.log("Raw response:", response);
      throw new Error("Failed to parse OpenAI response");
    }
    
    if (!result || !result.typos || result.typos.length === 0) {
      SpreadsheetApp.getUi().alert('No typos found in the selected text.');
      return;
    }
    
    // Show dialog with potential typos
    showTypoSuggestions(text, result.typos, customDict, range);
    
  } catch (error) {
    console.error('Error:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * Shows dialog with typo suggestions
 */
function showTypoSuggestions(originalText, typos, customDict, range) {
  const htmlContent = `
  <style>
    .typo-container { margin-bottom: 15px; }
    .word { font-weight: bold; color: red; }
    .replacements { margin-top: 5px; }
    .replacement-btn { margin-right: 5px; margin-bottom: 5px; }
    .add-to-dict { margin-top: 10px; }
  </style>
  
  <h3>Potential Typos Found</h3>
  <div id="typos-list">
    ${typos.map((typo, index) => {
      // Add custom dictionary words to replacements if they're similar
      const allReplacements = [...typo.replacements];
      customDict.forEach(dictWord => {
        if (similarWords(typo.word, dictWord) && !allReplacements.includes(dictWord)) {
          allReplacements.push(dictWord);
        }
      });
      
      return `
      <div class="typo-container" id="typo-${index}">
        <div class="word">"${typo.word}"</div>
        <div class="replacements">
          ${allReplacements.map(replacement => 
            `<button class="replacement-btn" onclick="replaceTypo(${index}, '${typo.word}', '${replacement}')">${replacement}</button>`
          ).join('')}
          <button class="replacement-btn" onclick="keepOriginal(${index}, '${typo.word}')">Keep Original</button>
        </div>
        <div class="add-to-dict">
          <input type="text" id="new-word-${index}" placeholder="Add custom replacement">
          <button onclick="addAndReplace(${index}, '${typo.word}')">Add & Replace</button>
        </div>
      </div>
      `;
    }).join('')}
  </div>
  <div style="margin-top: 20px;">
    <button onclick="google.script.host.close()">Done</button>
  </div>
  
  <script>
    // Store original text and corrections
    const originalText = ${JSON.stringify(originalText)};
    let correctedText = originalText;
    const typos = ${JSON.stringify(typos)};
    
    function replaceTypo(index, original, replacement) {
      correctedText = correctedText.replace(new RegExp(original, 'g'), replacement);
      document.getElementById('typo-' + index).style.display = 'none';
      updateText();
    }
    
    function keepOriginal(index) {
      document.getElementById('typo-' + index).style.display = 'none';
    }
    
    function addAndReplace(index, original) {
      const newWord = document.getElementById('new-word-' + index).value.trim();
      if (newWord) {
        google.script.run.addToCustomDictionary(newWord);
        replaceTypo(index, original, newWord);
      }
    }
    
    function updateText() {
      google.script.run.updateCellWithCorrectedText(correctedText);
    }
  </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(400);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Fix Typos');
  
  // Store the range for later use
  CacheService.getUserCache().put('ACTIVE_RANGE_A1', range.getA1Notation());
  CacheService.getUserCache().put('ACTIVE_SHEET_ID', range.getSheet().getSheetId().toString());
}

/**
 * Check if two words are similar (for custom dictionary suggestions)
 */
function similarWords(word1, word2) {
  word1 = word1.toLowerCase();
  word2 = word2.toLowerCase();
  
  // Simple check - first character the same and length difference ≤ 2
  return word1[0] === word2[0] && Math.abs(word1.length - word2.length) <= 2;
}

/**
 * Updates the cell with corrected text
 */
function updateCellWithCorrectedText(correctedText) {
  const rangeA1 = CacheService.getUserCache().get('ACTIVE_RANGE_A1');
  const sheetId = CacheService.getUserCache().get('ACTIVE_SHEET_ID');
  
  if (!rangeA1 || !sheetId) return;
  
  const sheet = getSheetById(sheetId);
  if (sheet) {
    const range = sheet.getRange(rangeA1);
    range.setValue(correctedText);
  }
}

/**
 * Get sheet by ID
 */
function getSheetById(id) {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId().toString() === id) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Shows the grammar improvement dialog
 */
function showGrammarImproveDialog() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  
  if (!range) {
    ui.alert('Please select a range first.');
    return;
  }
  
  const text = range.getValue().toString();
  if (!text) {
    ui.alert('Selected cell is empty.');
    return;
  }
  
  improveGrammar(text, range);
}

/**
 * Improves grammar using OpenAI API
 */
function improveGrammar(text, range) {
  const language = PropertiesService.getUserProperties().getProperty('GRAMMAR_CHECKER_LANGUAGE') || 'english';
  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Please set your OpenAI API key first.');
    showApiKeyDialog();
    return;
  }
  
  try {
    const languagePrompt = language === 'spanish' ? 
      'Mejora la gramática y claridad de este texto en español:' :
      'Improve the grammar and clarity of this English text:';
    
    const response = callOpenAI(apiKey, [
      {
        "role": "system", 
        "content": `You are a grammar assistant. ${languagePrompt} Return only the improved text without any additional explanation.`
      },
      {"role": "user", "content": text}
    ]);
    
    if (!response) {
      throw new Error("Failed to get response from OpenAI");
    }
    
    const improvedText = response.trim();
    
    // Show dialog to compare and accept changes
    showGrammarComparison(text, improvedText, range);
    
  } catch (error) {
    console.error('Error:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * Shows dialog to compare original and improved grammar
 */
function showGrammarComparison(originalText, improvedText, range) {
  const htmlContent = `
  <style>
    .text-container { margin-bottom: 15px; }
    .text-label { font-weight: bold; margin-bottom: 5px; }
    .text-content { 
      border: 1px solid #ccc; 
      padding: 10px; 
      background-color: #f9f9f9;
      white-space: pre-wrap;
      min-height: 100px;
    }
    .buttons { margin-top: 20px; }
  </style>
  
  <h3>Grammar Improvement</h3>
  
  <div class="text-container">
    <div class="text-label">Original Text:</div>
    <div class="text-content">${originalText}</div>
  </div>
  
  <div class="text-container">
    <div class="text-label">Improved Text:</div>
    <div class="text-content">${improvedText}</div>
  </div>
  
  <div class="buttons">
    <button onclick="acceptChanges()">Accept Changes</button>
    <button onclick="google.script.host.close()">Cancel</button>
  </div>
  
  <script>
    function acceptChanges() {
      const improvedText = ${JSON.stringify(improvedText)};
      google.script.run
        .withSuccessHandler(() => google.script.host.close())
        .updateCellWithImprovedGrammar(improvedText);
    }
  </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(400);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Improve Grammar');
  
  // Store the range for later use
  CacheService.getUserCache().put('ACTIVE_RANGE_A1', range.getA1Notation());
  CacheService.getUserCache().put('ACTIVE_SHEET_ID', range.getSheet().getSheetId().toString());
}

/**
 * Updates the cell with improved grammar
 */
function updateCellWithImprovedGrammar(improvedText) {
  const rangeA1 = CacheService.getUserCache().get('ACTIVE_RANGE_A1');
  const sheetId = CacheService.getUserCache().get('ACTIVE_SHEET_ID');
  
  if (!rangeA1 || !sheetId) return;
  
  const sheet = getSheetById(sheetId);
  if (sheet) {
    const range = sheet.getRange(rangeA1);
    range.setValue(improvedText);
  }
}

/**
 * Calls the OpenAI API
 */
function callOpenAI(apiKey, messages) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    'model': 'gpt-4',
    'messages': messages,
    'temperature': 0.3,
    'max_tokens': 1000
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();
  
  if (responseCode !== 200) {
    console.error('API Error:', responseCode, responseBody);
    throw new Error(`API Error ${responseCode}: ${responseBody}`);
  }
  
  const parsedResponse = JSON.parse(responseBody);
  return parsedResponse.choices[0].message.content;
}

/**
 * Manages custom dictionary
 */
function showDictionaryManager() {
  const customDict = getCustomDictionary();
  
  const htmlContent = `
  <style>
    .word-list {
      margin: 15px 0;
      max-height: 200px;
      overflow-y: auto;
      border: 1px solid #ccc;
      padding: 10px;
    }
    .word-item {
      margin-bottom: 5px;
      display: flex;
      justify-content: space-between;
    }
    .delete-btn {
      color: red;
      cursor: pointer;
    }
    .add-word {
      margin-top: 15px;
    }
  </style>
  
  <h3>Custom Dictionary</h3>
  
  <div class="word-list" id="word-list">
    ${customDict.length > 0 ? 
      customDict.map((word, index) => 
        `<div class="word-item">
          <span>${word}</span>
          <span class="delete-btn" onclick="deleteWord(${index})">&times;</span>
        </div>`
      ).join('') : 
      '<div>No words in custom dictionary yet.</div>'
    }
  </div>
  
  <div class="add-word">
    <input type="text" id="new-word" placeholder="Add new word">
    <button onclick="addWord()">Add</button>
  </div>
  
  <div style="margin-top: 20px;">
    <button onclick="google.script.host.close()">Done</button>
  </div>
  
  <script>
    function deleteWord(index) {
      google.script.run
        .withSuccessHandler(refreshList)
        .removeFromCustomDictionary(index);
    }
    
    function addWord() {
      const newWord = document.getElementById('new-word').value.trim();
      if (newWord) {
        google.script.run
          .withSuccessHandler(refreshList)
          .addToCustomDictionary(newWord);
        document.getElementById('new-word').value = '';
      }
    }
    
    function refreshList(newDict) {
      const wordList = document.getElementById('word-list');
      if (newDict.length > 0) {
        wordList.innerHTML = newDict.map((word, index) => 
          \`<div class="word-item">
            <span>\${word}</span>
            <span class="delete-btn" onclick="deleteWord(\${index})">&times;</span>
          </div>\`
        ).join('');
      } else {
        wordList.innerHTML = '<div>No words in custom dictionary yet.</div>';
      }
    }
  </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(350);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Custom Dictionary');
}

/**
 * Gets the custom dictionary
 */
function getCustomDictionary() {
  const dictString = PropertiesService.getUserProperties().getProperty('CUSTOM_DICTIONARY');
  return dictString ? JSON.parse(dictString) : [];
}

/**
 * Adds a word to the custom dictionary
 */
function addToCustomDictionary(word) {
  word = word.trim();
  if (!word) return getCustomDictionary();
  
  const dict = getCustomDictionary();
  if (!dict.includes(word)) {
    dict.push(word);
    PropertiesService.getUserProperties().setProperty('CUSTOM_DICTIONARY', JSON.stringify(dict));
  }
  return dict;
}

/**
 * Removes a word from the custom dictionary
 */
function removeFromCustomDictionary(index) {
  const dict = getCustomDictionary();
  if (index >= 0 && index < dict.length) {
    dict.splice(index, 1);
    PropertiesService.getUserProperties().setProperty('CUSTOM_DICTIONARY', JSON.stringify(dict));
  }
  return dict;
}

/**
 * Configures triggers for the spreadsheet
 */
function configTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  // Check if the onOpenDocument trigger already exists
  let triggerExists = false;
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpenDocument') {
      triggerExists = true;
      break;
    }
  }
  
  const result = ui.alert(
    'Configure Triggers',
    triggerExists ? 
      'The onOpenDocument trigger is already set up. Would you like to remove it?' : 
      'Would you like to set up the onOpenDocument trigger to run when the spreadsheet opens?',
    ui.ButtonSet.YES_NO
  );
  
  if (result == ui.Button.YES) {
    if (triggerExists) {
      // Remove the existing trigger
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'onOpenDocument') {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
      ui.alert('Trigger removed successfully!');
    } else {
      // Create a new trigger
      ScriptApp.newTrigger('onOpenDocument')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onOpen()
        .create();
      ui.alert('Trigger created successfully!');
    }
  }
}
