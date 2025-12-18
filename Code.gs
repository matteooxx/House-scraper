/**
 * UNIVERSAL AI HOUSE HUNTER (GitHub Edition) - V5.2
 * * A professional AI-powered real estate workstation.
 * * Updates:
 * - Fixed "Formula Parse Error": Switched to semicolon (;) for Italian/EU locales.
 * - Fixed Maps URL: Using official Google Maps Search API.
 * - Dual Inquiry Drafts (Local & Target Language).
 * - Full Geographic Breakdown & AI Scoring.
 */

const PROPERTIES = PropertiesService.getScriptProperties();
const API_KEY_NAME = 'AI_HOUSE_HUNTER_KEY';
const JINA_KEY_NAME = 'JINA_API_KEY';
const PROVIDER_NAME = 'AI_HOUSE_HUNTER_PROVIDER';
const MODEL_NAME = 'AI_HOUSE_HUNTER_MODEL';
const LANGUAGE_NAME = 'AI_HOUSE_HUNTER_LANG';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu('🏠 AI House Hunter');
  mainMenu.addItem('1. Initialize Sheets', 'setupSpreadsheet');
  mainMenu.addSeparator();
  
  const providerSubMenu = ui.createMenu('2. Select AI Provider');
  providerSubMenu.addItem('Google Gemini', 'setProviderGemini');
  providerSubMenu.addItem('OpenAI (GPT-4o)', 'setProviderOpenAI');
  mainMenu.addSubMenu(providerSubMenu);
  
  const apiSubMenu = ui.createMenu('3. API Management');
  apiSubMenu.addItem('Set AI API Key', 'setApiKey');
  apiSubMenu.addItem('Set Jina AI Key (Optional)', 'setJinaKey');
  apiSubMenu.addItem('Check Status', 'checkSettings');
  apiSubMenu.addItem('⚡ Test AI Connection', 'testConnection');
  mainMenu.addSubMenu(apiSubMenu);

  mainMenu.addSeparator();
  mainMenu.addItem('4. Set Output Language', 'setOutputLanguage');
  mainMenu.addItem('5. Process All New Items', 'processLinks');
  
  mainMenu.addSeparator();
  mainMenu.addItem('📦 Archive Processed Items', 'archiveListings');
  mainMenu.addItem('🗑️ Clear Database', 'clearDatabase');
  mainMenu.addToUi();
}

/** SETTINGS HANDLERS */
function setProviderGemini() { PROPERTIES.setProperties({ [PROVIDER_NAME]: 'gemini', [MODEL_NAME]: 'gemini-flash-latest' }); SpreadsheetApp.getActiveSpreadsheet().toast("Provider: Gemini", "Updated ✅"); }
function setProviderOpenAI() { PROPERTIES.setProperties({ [PROVIDER_NAME]: 'openai', [MODEL_NAME]: 'gpt-4o' }); SpreadsheetApp.getActiveSpreadsheet().toast("Provider: OpenAI", "Updated ✅"); }

function setOutputLanguage() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Output Language', 'Target language (e.g., Italian, English, Chinese):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) { PROPERTIES.setProperty(LANGUAGE_NAME, response.getResponseText().trim()); ui.alert('✅ Language saved.'); }
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('AI API Key', 'Paste Gemini/OpenAI key:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) { PROPERTIES.setProperty(API_KEY_NAME, response.getResponseText().trim()); ui.alert('✅ AI Key saved.'); }
}

function setJinaKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Jina AI Key', 'Paste Jina API Key:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) { PROPERTIES.setProperty(JINA_KEY_NAME, response.getResponseText().trim()); ui.alert('✅ Jina Key saved.'); }
}

function checkSettings() {
  const p = PROPERTIES.getProperty(PROVIDER_NAME) || "Not set";
  const l = PROPERTIES.getProperty(LANGUAGE_NAME) || "English";
  const aiK = PROPERTIES.getProperty(API_KEY_NAME) ? "✅ Stored" : "❌ Missing";
  SpreadsheetApp.getUi().alert(`Provider: ${p}\nLanguage: ${l}\nAI Key: ${aiK}`);
}

function testConnection() {
  const ui = SpreadsheetApp.getUi();
  const apiKey = PROPERTIES.getProperty(API_KEY_NAME);
  const provider = PROPERTIES.getProperty(PROVIDER_NAME);
  if (!apiKey || !provider) return ui.alert("Setup Missing.");
  ui.alert("Testing connection...");
  const testData = analyzeContent("Test: House in London.", apiKey, provider, "English", "Any preference.");
  if (testData.error) ui.alert("❌ Failed:\n" + testData.error);
  else ui.alert("✅ Success!");
}

/** UI AND SHEET INITIALIZATION */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Preferences Sheet
  let prefSheet = ss.getSheetByName('Preferences') || ss.insertSheet('Preferences');
  prefSheet.clear().getRange("A1:B1").setValues([["Attribute", "My Requirements (Be descriptive)"]]).setFontWeight("bold").setBackground("#e67e22").setFontColor("white");
  if (prefSheet.getLastRow() < 2) {
    prefSheet.appendRow(["Desired Location", "Close to metro stations, safe neighborhood"]);
    prefSheet.appendRow(["Max Budget", "Max 1200 Monthly inclusive of bills"]);
    prefSheet.appendRow(["Must Have", "Balcony, dishwasher, natural light"]);
    prefSheet.appendRow(["Inquiry Style", "Polite, mention I am a quiet professional with no pets"]);
  }
  prefSheet.setColumnWidth(1, 150);
  prefSheet.setColumnWidth(2, 600);

  // 2. Links Sheet
  let linkSheet = ss.getSheetByName('Links') || ss.insertSheet('Links');
  linkSheet.clear().getRange(1, 1, 1, 2).setValues([['Listing URL (Auto)', 'Manual Content (Paste text)']]).setFontWeight("bold").setFontColor("white").setBackground("#34495e");
  linkSheet.setColumnWidth(1, 300);
  linkSheet.setColumnWidth(2, 600);
  linkSheet.setFrozenRows(1);

  // 3. Database Sheet (The Dashboard)
  let dbSheet = ss.getSheetByName('Database') || ss.insertSheet('Database');
  dbSheet.clear().clearFormats();
  const oldBandings = dbSheet.getBandings();
  for (let i = 0; i < oldBandings.length; i++) oldBandings[i].remove();

  const headersDB = [['URL Source', 'Status', 'Score', 'M-Price', 'Cur', 'W-Price', 'Cur', 'Available', 'Type', 'Shared', 'Bath', 'Utilities', 'Country', 'City', 'Street', 'Neighborhood', 'Map', 'Contact', 'AI Summary', 'Draft (Local)', 'Draft (Target Lang)']];
  dbSheet.getRange(1, 1, 1, 21).setValues(headersDB).setFontWeight("bold").setFontColor("white").setBackground("#2c3e50").setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  dbSheet.setFrozenRows(1);
  const widths = [100, 85, 50, 65, 40, 65, 40, 85, 75, 50, 60, 140, 60, 90, 120, 120, 90, 140, 350, 400, 400];
  for (let i = 0; i < widths.length; i++) dbSheet.setColumnWidth(i + 1, widths[i]);
  
  dbSheet.getRange("A2:U1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle");
  dbSheet.getRange("A2:A1000").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
  dbSheet.getRange("B2:K1000").setHorizontalAlignment("center");
  dbSheet.getRange("B2:B1000").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['will review', 'in review', 'good', 'bad', 'applied'], true).build());
  
  dbSheet.getRange(2, 1, 500, 21).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  applyStatusFormatting(dbSheet);
  
  // 4. Archive Sheet
  let archiveSheet = ss.getSheetByName('Archive') || ss.insertSheet('Archive');
  if (archiveSheet.getLastRow() === 0) {
    dbSheet.getRange(1,1,1,21).copyTo(archiveSheet.getRange(1,1));
    archiveSheet.setFrozenRows(1);
  }

  ss.toast("Pro Dashboard V5.2 Fixed!", "Success ✅");
}

function applyStatusFormatting(sheet) {
  const range = sheet.getRange("A2:U1000");
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('good').setBackground('#dff0d8').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('bad').setBackground('#f2dede').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('applied').setBackground('#d9edf7').setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

/** AUTO-ARCHIVE FUNCTION */
function archiveListings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('Database');
  const archiveSheet = ss.getSheetByName('Archive');
  const data = dbSheet.getDataRange().getValues();
  
  let rowsDeleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    let status = data[i][1];
    if (status === 'bad' || status === 'applied') {
      archiveSheet.appendRow(data[i]);
      dbSheet.deleteRow(i + 1);
      rowsDeleted++;
    }
  }
  ss.toast(`Moved ${rowsDeleted} items to Archive.`, "Archive Complete");
}

function clearDatabase() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database');
  if (sheet && sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
}

/** CORE ENGINE */
function processLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = PROPERTIES.getProperty(API_KEY_NAME);
  const provider = PROPERTIES.getProperty(PROVIDER_NAME);
  const lang = PROPERTIES.getProperty(LANGUAGE_NAME) || "English";
  if (!apiKey || !provider) return SpreadsheetApp.getUi().alert("Setup Missing.");

  const prefSheet = ss.getSheetByName('Preferences');
  const userPrefs = prefSheet.getRange("A2:B" + prefSheet.getLastRow()).getValues().map(r => r[0] + ": " + r[1]).join(". ");

  const inputSheet = ss.getSheetByName('Links');
  const dbSheet = ss.getSheetByName('Database');
  const inputData = inputSheet.getRange("A2:B" + inputSheet.getLastRow()).getValues();
  const existingLinks = dbSheet.getRange("A:A").getValues().flat().filter(String);
  const skipList = new Set(existingLinks);

  for (let i = 0; i < inputData.length; i++) {
    let url = inputData[i][0].toString().trim();
    let manualContent = inputData[i][1].toString().trim();
    let source = url || "Manual Entry " + (i+1);
    if (!source || skipList.has(source)) continue;

    let textToAnalyze = manualContent || scrapeWithJina(url);
    if (!textToAnalyze || textToAnalyze.length < 50) continue;

    try {
      const result = analyzeContent(textToAnalyze, apiKey, provider, lang, userPrefs);
      if (result.error) break;
      const data = result.data;
      
      // OFFICIAL MAPS SEARCH URL
      const mapQuery = encodeURIComponent(`${data.street}, ${data.city}, ${data.country}`);
      const mapsUrl = `https://www.google.com/maps/search/?api=1&query=${mapQuery}`;
      
      // FORMULA FIX: Using semicolon (;) for Italian/European Sheet settings
      const mapsHyperlink = `=HYPERLINK("${mapsUrl}"; "📍 View Map")`;
      
      dbSheet.appendRow([
        source, 'will review', data.score,
        data.monthly_price, data.monthly_currency, data.weekly_price, data.weekly_currency, 
        data.available_from, data.property_type, data.is_shared, data.bathroom_type, 
        data.utilities, data.country, data.city, data.street, data.neighborhood,
        mapsHyperlink, data.contact_info, data.notes, data.inquiry_draft_local, data.inquiry_draft_target
      ]);
      skipList.add(source);
      Utilities.sleep(2000);
    } catch (e) { Logger.log(e); }
  }
}

function scrapeWithJina(url) {
  try {
    const jinaKey = PROPERTIES.getProperty(JINA_KEY_NAME);
    const headers = { "X-With-Links-Summary": "true", "X-Target-Selector": "body" };
    if (jinaKey) headers["Authorization"] = "Bearer " + jinaKey;
    const res = UrlFetchApp.fetch('https://r.jina.ai/' + url, { headers: headers, muteHttpExceptions: true });
    return res.getContentText().substring(0, 30000);
  } catch (e) { return null; }
}

function analyzeContent(text, apiKey, provider, lang, prefs) {
  try {
    const model = PROPERTIES.getProperty(MODEL_NAME) || (provider === 'openai' ? 'gpt-4o' : 'gemini-flash-latest');
    const prompt = `Act as a real estate assistant.
    1. Compare property with USER PREFERENCES: [${prefs}]. Assign "score" (1-100).
    2. Generate "inquiry_draft_local" in the LISTING'S LOCAL LANGUAGE (e.g. German if ad is German).
    3. Generate "inquiry_draft_target" being the same draft translated into ${lang}.
    4. Translate all other fields into ${lang}.
    Return ONLY JSON: {
      "score": number, "monthly_price": number, "monthly_currency": "ISO", "weekly_price": number, "weekly_currency": "ISO",
      "available_from": "string", "property_type": "string", "is_shared": "Yes/No", "bathroom_type": "string",
      "utilities": "string", "country": "string", "city": "string", "street": "string",
      "neighborhood": "string", "contact_info": "string", "notes": "detailed summary", 
      "inquiry_draft_local": "string", "inquiry_draft_target": "string"
    }. Text: ${text}`;
    
    let apiResponse;
    const options = { method: 'post', contentType: 'application/json', muteHttpExceptions: true };
    if (provider === 'openai') {
      options.headers = { Authorization: 'Bearer ' + apiKey };
      options.payload = JSON.stringify({ model: model, messages: [{ role: "user", content: prompt }], response_format: { type: "json_object" }});
      apiResponse = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    } else {
      const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      options.payload = JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] });
      apiResponse = UrlFetchApp.fetch(geminiUrl, options);
    }
    if (apiResponse.getResponseCode() !== 200) return { error: true };
    const json = JSON.parse(apiResponse.getContentText());
    let resText = (provider === 'openai') ? json.choices[0].message.content : json.candidates[0].content.parts[0].text;
    return { data: JSON.parse(resText.replace(/```json|```/g, '').trim()) };
  } catch (err) { return { error: true }; }
}