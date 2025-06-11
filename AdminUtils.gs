// File: AdminUtils.gs
// Description: Contains administrative utility functions for project setup and configuration,
// such as managing API keys stored in UserProperties.

// In AdminUtils.gs

/**
 * Provides a UI prompt to set the shared Gemini API Key in UserProperties.
 * Uses the GEMINI_API_KEY_PROPERTY constant from Config.gs.
 */
function setSharedGeminiApiKey_UI() {
  const MIN_KEY_LENGTH = 35;
  const propertyName = GEMINI_API_KEY_PROPERTY; // From Config.gs
  const userProps = PropertiesService.getUserProperties();
  const currentKey = userProps.getProperty(propertyName);

  let uiAvailable = false;
  let ui;

  try {
    ui = SpreadsheetApp.getUi();
    if (ui) {
        uiAvailable = true;
    }
  } catch (e) {
    Logger.log('SpreadsheetApp.getUi() is not available: ' + e.message);
    uiAvailable = false;
  }

  if (uiAvailable) {
    // --- Use SpreadsheetApp.getUi() ---
    try {
      const initialPromptMessage = `Enter the shared Google AI Gemini API Key for all modules.\nThis will be stored in UserProperties under the key: "${propertyName}".\nPlease ensure you paste the full and correct API key (typically ${MIN_KEY_LENGTH}+ characters).\n${currentKey ? 'An existing key is currently set and will be overwritten.' : 'No key currently set.'}`;
      const response = ui.prompt('Set Shared Gemini API Key', initialPromptMessage, ui.ButtonSet.OK_CANCEL);

      if (response.getSelectedButton() == ui.Button.OK) {
        const apiKey = response.getResponseText().trim();

        if (apiKey && apiKey.length >= MIN_KEY_LENGTH && /^[a-zA-Z0-9_~-]+$/.test(apiKey)) {
          // Key is syntactically valid, now check if we are overwriting
          if (currentKey && currentKey !== apiKey) { // Only ask for confirmation if a different key exists
            const confirmOverwrite = ui.alert(
              'Confirm Overwrite',
              `An API key already exists. Do you want to overwrite it with the new key?\n\nOld Key (masked): ${currentKey.substring(0,4)}...${currentKey.substring(currentKey.length-4)}\nNew Key (masked): ${apiKey.substring(0,4)}...${apiKey.substring(apiKey.length-4)}`,
              ui.ButtonSet.YES_NO
            );
            if (confirmOverwrite !== ui.Button.YES) {
              ui.alert('Operation Cancelled', 'The existing API key was not overwritten.', ui.ButtonSet.OK);
              return; // Exit if user cancels overwrite
            }
          } else if (currentKey && currentKey === apiKey) {
             ui.alert('No Change', 'The entered API key is the same as the currently stored key. No changes made.', ui.ButtonSet.OK);
             return;
          }

          // Proceed to set the property
          userProps.setProperty(propertyName, apiKey);
          ui.alert('API Key Saved', `The Gemini API Key has been saved successfully for the project under property: "${propertyName}".`);

        } else if (apiKey) { // Entered key failed validation
          let N = "";
          if(apiKey.length < MIN_KEY_LENGTH) N = `The key appears to be too short (should be at least ${MIN_KEY_LENGTH} characters). `
          if(!/^[a-zA-Z0-9_~-]+$/.test(apiKey)) N += `The key contains unexpected characters. API keys typically consist of letters, numbers, hyphens, underscores, or tildes.`
          ui.alert('API Key Not Saved', `The entered key does not appear to be valid. ${N}Please check and try again.`, ui.ButtonSet.OK);
        } else { // No API key entered
          ui.alert('API Key Not Saved', 'No API key was entered.', ui.ButtonSet.OK);
        }
      } else { // User cancelled the initial prompt
        ui.alert('API Key Setup Cancelled', 'The API key setup process was cancelled at the input stage.', ui.ButtonSet.OK);
      }
    } catch (uiError) {
      Logger.log('Error during SpreadsheetApp.getUi() interaction: ' + uiError.message);
      // Optionally, you could add a generic error alert here if the specific alerts above didn't cover it.
      // ui.alert('Error', 'An unexpected UI error occurred: ' + uiError.message, ui.ButtonSet.OK);
    }
  } else {
    // --- Fallback to Browser.inputBox ---
    Logger.log('Attempting Browser.inputBox for API Key input.');
    try {
      const currentKeyInfo = currentKey ? "(An existing key will be overwritten if you proceed)" : "(No key currently set)";
      const apiKeyInput = Browser.inputBox(`Set Shared Gemini API Key`, `Enter the shared Gemini API Key.\nStored as property: ${propertyName}.\nPlease ensure it's the full key (typically ${MIN_KEY_LENGTH}+ characters).\n${currentKeyInfo}`, Browser.Buttons.OK_CANCEL);

      if (apiKeyInput !== 'cancel' && apiKeyInput !== null) {
        const trimmedApiKey = apiKeyInput.trim();

        if (trimmedApiKey && trimmedApiKey.length >= MIN_KEY_LENGTH && /^[a-zA-Z0-9_~-]+$/.test(trimmedApiKey)) {
          // Key is syntactically valid, now check for overwrite confirmation (using msgBox which only has OK for yes/no effectively)
          if (currentKey && currentKey !== trimmedApiKey) {
             const confirmOverwriteResponse = Browser.msgBox(
                'Confirm Overwrite',
                `An API key already exists. Do you want to overwrite it with the new key?\n\nOld Key (masked): ${currentKey.substring(0,4)}...${currentKey.substring(currentKey.length-4)}\nNew Key (masked): ${trimmedApiKey.substring(0,4)}...${trimmedApiKey.substring(trimmedApiKey.length-4)}\n\n(Click 'Cancel' to keep the old key, 'OK' to overwrite)`,
                Browser.Buttons.OK_CANCEL // OK means "Yes, overwrite", Cancel means "No, don't"
            );
            if (confirmOverwriteResponse !== 'ok') { // 'ok' is the return for OK button
                Browser.msgBox('Operation Cancelled', 'The existing API key was not overwritten.', Browser.Buttons.OK);
                return; // Exit
            }
          } else if (currentKey && currentKey === trimmedApiKey) {
             Browser.msgBox('No Change', 'The entered API key is the same as the currently stored key. No changes made.', Browser.Buttons.OK);
             return;
          }

          userProps.setProperty(propertyName, trimmedApiKey);
          Browser.msgBox('API Key Saved', `The Gemini API Key has been saved successfully under property: "${propertyName}".`, Browser.Buttons.OK);
        } else if (trimmedApiKey) {
          let N = "";
          if(trimmedApiKey.length < MIN_KEY_LENGTH) N = `The key appears to be too short (should be at least ${MIN_KEY_LENGTH} characters). `
          if(!/^[a-zA-Z0-9_~-]+$/.test(trimmedApiKey)) N += `The key contains unexpected characters. API keys typically consist of letters, numbers, hyphens, underscores, or tildes.`
          Browser.msgBox('API Key Not Saved', `The entered key does not appear to be valid. ${N}Please check and try again.`, Browser.Buttons.OK);
        } else {
          Browser.msgBox('API Key Not Saved', 'No API key was entered.', Browser.Buttons.OK);
        }
      } else {
        Browser.msgBox('API Key Setup Cancelled', 'The API key setup process was cancelled.', Browser.Buttons.OK);
      }
    } catch (e2) {
      Logger.log('Browser.inputBox also failed: ' + e2.message);
    }
  }
}

/**
 * TEMPORARY: Manually sets the shared Gemini API Key in UserProperties.
 * Edit YOUR_GEMINI_KEY_HERE in the code before running.
 * REMOVE OR CLEAR THE KEY FROM CODE AFTER RUNNING FOR SECURITY.
 */
function TEMPORARY_manualSetSharedGeminiApiKey() { // Renamed
  const YOUR_GEMINI_KEY_HERE = 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; // <<< EDIT THIS LINE WITH YOUR KEY
  const propertyName = GEMINI_API_KEY_PROPERTY; // From Config.gs

  if (YOUR_GEMINI_KEY_HERE === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || YOUR_GEMINI_KEY_HERE.trim() === '') {
    const msg = `ERROR: Gemini API Key not set in TEMPORARY_manualSetSharedGeminiApiKey function. Edit the script code first with your key (variable YOUR_GEMINI_KEY_HERE). It should be stored under UserProperty: "${propertyName}".`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert('Action Required', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
    catch(e) { try { Browser.msgBox('Action Required', msg, Browser.Buttons.OK); } catch(e2) {} }
    return;
  }
  PropertiesService.getUserProperties().setProperty(propertyName, YOUR_GEMINI_KEY_HERE);
  const successMsg = `UserProperty "${propertyName}" has been MANUALLY SET with the hardcoded Gemini API Key. IMPORTANT: For security, now remove or comment out the TEMPORARY_manualSetSharedGeminiApiKey function, or at least clear the YOUR_GEMINI_KEY_HERE variable in the code.`;
  Logger.log(successMsg);
  try { SpreadsheetApp.getUi().alert('API Key Manually Set', successMsg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { try { Browser.msgBox('API Key Manually Set', successMsg, Browser.Buttons.OK); } catch(e2) {} }
}

/**
 * Displays all UserProperties set for this script project to the logs.
 * Sensitive values like API keys are partially masked.
 */
function showAllUserProperties() { // Renamed
  const userProps = PropertiesService.getUserProperties().getProperties();
  let logOutput = "Current UserProperties for this script project:\n";
  if (Object.keys(userProps).length === 0) {
    logOutput += "  (No UserProperties are currently set for this project)\n";
  } else {
    for (const key in userProps) {
      let value = userProps[key];
      // Mask sensitive values - adjust keywords if needed
      if (key.toLowerCase().includes('api') || key.toLowerCase().includes('key') || key.toLowerCase().includes('secret')) {
        if (value && typeof value === 'string' && value.length > 10) {
          value = value.substring(0, 4) + "..." + value.substring(value.length - 4);
        } else if (value && typeof value === 'string') {
            value = "**** (short value)";
        }
      }
      logOutput += `  ${key}: ${value}\n`;
    }
  }
  Logger.log(logOutput);
  const alertMsg = "Current UserProperties have been logged. Check Apps Script logs (View > Logs or Executions) to see them. Sensitive values are partially masked in the log.";
  try { SpreadsheetApp.getUi().alert("User Properties Logged", alertMsg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { try { Browser.msgBox("User Properties Logged", alertMsg, Browser.Buttons.OK); } catch(e2) {} }
}
