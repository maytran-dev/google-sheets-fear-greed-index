/**
 * Google Apps Script to fetch the latest Fear and Greed Index from CoinMarketCap API
 * with email alerts for extreme values
 * 
 * Setup Instructions:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this script
 * 4. Update EMAIL_RECIPIENTS with your email addresses
 * 5. Save and run the fetchLatestFearGreedIndex function
 * 6. Authorize the script when prompted (including email permissions)
 */

// Configuration
const API_KEY = '8c96cadd-7c73-41fa-b536-b33740a0d273'; // Your CoinMarketCap API key
const API_URL = 'https://pro-api.coinmarketcap.com/v3/fear-and-greed/latest';
const SHEET_NAME = 'Fear & Greed'; // Name of the sheet to paste data

// Email Alert Configuration
const EMAIL_RECIPIENTS = 'email1@example.com'; // Single email or comma-separated list: 'email1@example.com, email2@example.com'
const FEAR_THRESHOLD = 40; // Send alert when value < 40
const GREED_THRESHOLD = 60; // Send alert when value > 60
const ENABLE_EMAIL_ALERTS = true; // Set to false to disable email alerts

/**
 * Main function to fetch the latest Fear and Greed value
 */
function fetchLatestFearGreedIndex() {
  try {
    console.log('Fetching latest Fear and Greed Index...');
    
    // Fetch data from CoinMarketCap API
    const response = fetchFromAPI();
    
    if (!response || !response.data) {
      SpreadsheetApp.getUi().alert('No data received from API');
      return;
    }
    
    // Get the data object (no longer an array)
    const latestData = response.data;
    
    // Write to sheet
    writeToSheet(latestData);
    
    // Check if email alert should be sent
    if (ENABLE_EMAIL_ALERTS) {
      checkAndSendAlert(latestData);
    }
    
    // Show success message
    SpreadsheetApp.getUi().alert('Latest Fear & Greed Index updated successfully!');
    
  } catch (error) {
    console.error('Error:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Function to run silently (for triggers) without UI alerts
 */
function fetchLatestFearGreedIndexSilent() {
  try {
    console.log('Fetching latest Fear and Greed Index (silent mode)...');
    
    // Fetch data from CoinMarketCap API
    const response = fetchFromAPI();
    
    if (!response || !response.data) {
      console.error('No data received from API');
      return;
    }
    
    // Get the data object
    const latestData = response.data;
    
    // Write to sheet
    writeToSheet(latestData);
    
    // Check if email alert should be sent
    if (ENABLE_EMAIL_ALERTS) {
      checkAndSendAlert(latestData);
    }
    
    console.log('Update completed successfully');
    
  } catch (error) {
    console.error('Error in silent fetch:', error);
    // Optionally send error notification email
    sendErrorNotification(error);
  }
}

/**
 * Check thresholds and send email alert if needed
 */
function checkAndSendAlert(data) {
  const value = data.value;
  const classification = data.value_classification;
  
  // Use the update_time from the API response
  const updateTime = new Date(data.update_time);
  const formattedDateTime = Utilities.formatDate(updateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // Check if value crosses thresholds
  if (value < FEAR_THRESHOLD || value > GREED_THRESHOLD) {
    sendAlertEmail(value, classification, formattedDateTime);
  }
}

/**
 * Send email alert
 */
function sendAlertEmail(value, classification, dateTime) {
  try {
    // Determine alert type
    let alertType = '';
    let alertEmoji = '';
    let recommendation = '';
    
    if (value < FEAR_THRESHOLD) {
      alertType = 'FEAR ALERT';
      alertEmoji = '';
      recommendation = 'Market sentiment is fearful. This might be a buying opportunity ("Be greedy when others are fearful").';
    } else if (value > GREED_THRESHOLD) {
      alertType = 'GREED ALERT';
      alertEmoji = '';
      recommendation = 'Market sentiment is greedy. Consider taking profits or being cautious ("Be fearful when others are greedy").';
    }
    
    // Create email subject - keep it simple for Gmail
    const subject = `${alertEmoji} ${alertType}: Fear & Greed Index at ${value} (${classification})`;
    
    // Create email body with HTML formatting - Gmail-compatible
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #d1dda5 0%, #4ba28a 100%); color: white; padding: 20px; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; text-align: center;">${alertEmoji} ${alertType} ${alertEmoji}</h1>
        </div>
        
        <div style="background: #f7f7f7; padding: 20px; border: 1px solid #ddd; border-top: none;">
          <h2 style="color: #333; margin-top: 0;">Fear &amp; Greed Index Update</h2>
          
          <div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>Current Value:</strong></td>
                <td style="padding: 8px; border-bottom: 1px solid #eee; text-align: right; font-size: 24px; color: ${value < FEAR_THRESHOLD ? '#ff4d4d' : '#009900'};">
                  <strong>${value}</strong>
                </td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>Classification:</strong></td>
                <td style="padding: 8px; border-bottom: 1px solid #eee; text-align: right;">
                  <span style="background: ${value < FEAR_THRESHOLD ? '#ff4d4d' : '#009900'}; color: white; padding: 4px 8px; border-radius: 4px;">
                    ${classification}
                  </span>
                </td>
              </tr>
              <tr>
                <td style="padding: 8px;"><strong>Date &amp; Time:</strong></td>
                <td style="padding: 8px; text-align: right;">${dateTime}</td>
              </tr>
            </table>
          </div>
          
          <div style="background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; border-radius: 4px; margin-bottom: 15px;">
            <h3 style="margin-top: 0; color: #856404;"> Recommendation</h3>
            <p style="margin-bottom: 0; color: #856404;">${recommendation}</p>
          </div>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
          
          <p style="color: #666; font-size: 12px; text-align: center; margin-bottom: 0;">
            This alert was triggered because the Fear &amp; Greed Index crossed your threshold settings:<br>
            Fear Alert: &lt; ${FEAR_THRESHOLD} | Greed Alert: &gt; ${GREED_THRESHOLD}<br>
            <a href="https://coinmarketcap.com/charts/fear-and-greed-index/" style="color: #667eea;">View Coin Market Cap Chart</a>
          </p>
        </div>
      </div>
    `;
    
    // Plain text version
    const textBody = `
${alertEmoji} ${alertType}

Fear & Greed Index Update:
- Current Value: ${value}
- Classification: ${classification}
- Date & Time: ${dateTime}

${recommendation}

This alert was triggered because the index crossed your threshold settings:
Fear Alert: < ${FEAR_THRESHOLD} | Greed Alert: > ${GREED_THRESHOLD}

View Spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
    `;
    
    // Send email with proper encoding
    GmailApp.sendEmail(
      EMAIL_RECIPIENTS,
      subject,
      textBody,
      {
        htmlBody: htmlBody,
        name: 'Fear & Greed Alert System'
      }
    );
    
    console.log(`Alert email sent to ${EMAIL_RECIPIENTS}`);
    
  } catch (error) {
    console.error('Error sending email:', error);
  }
}

/**
 * Send error notification email
 */
function sendErrorNotification(error) {
  try {
    const subject = ' Fear & Greed Index - Error Notification';
    const body = `
An error occurred while fetching the Fear & Greed Index:

Error: ${error.toString()}
Time: ${new Date().toString()}

Please check the Google Apps Script logs for more details.
    `;
    
    GmailApp.sendEmail(EMAIL_RECIPIENTS, subject, body);
  } catch (emailError) {
    console.error('Failed to send error notification:', emailError);
  }
}

/**
 * Fetch data from the CoinMarketCap API
 */
function fetchFromAPI() {
  const options = {
    'method': 'GET',
    'headers': {
      'X-CMC_PRO_API_KEY': API_KEY,
      'Accept': 'application/json'
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`API error: ${responseCode}`);
    }
    
    return JSON.parse(response.getContentText());
    
  } catch (error) {
    throw new Error(`Failed to fetch data: ${error.message}`);
  }
}

/**
 * Write the latest data to Google Sheet
 */
function writeToSheet(data) {
  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create the sheet
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set headers
  const headers = [['Date & Time', 'Value', 'Classification']];
  
  // Use the update_time from the API response
  const updateTime = new Date(data.update_time);
  const formattedDateTime = Utilities.formatDate(updateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // Prepare data row
  const dataRow = [[
    formattedDateTime,
    data.value,
    data.value_classification
  ]];
  
  // Write headers
  sheet.getRange(1, 1, 1, 3).setValues(headers);
  
  // Write data
  sheet.getRange(2, 1, 1, 3).setValues(dataRow);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1e3a5f');
  headerRange.setFontColor('#ffffff');
  
  // Add color formatting to the value cell based on classification
  const valueCell = sheet.getRange(2, 2);
  const classification = data.value_classification;
  
  // Apply color based on classification
  if (classification === 'Extreme Fear') {
    valueCell.setBackground('#ff4d4d').setFontColor('#ffffff');
  } else if (classification === 'Fear') {
    valueCell.setBackground('#ff9933').setFontColor('#000000');
  } else if (classification === 'Neutral') {
    valueCell.setBackground('#ffcc00').setFontColor('#000000');
  } else if (classification === 'Greed') {
    valueCell.setBackground('#66cc66').setFontColor('#000000');
  } else if (classification === 'Extreme Greed') {
    valueCell.setBackground('#009900').setFontColor('#ffffff');
  }
  
  // Auto-resize columns
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.autoResizeColumn(3);
  
  // Center align value and classification
  sheet.getRange(2, 2, 1, 2).setHorizontalAlignment('center');
}

/**
 * Test email alert function
 */
function testEmailAlert() {
  // Send a test email with sample data
  sendAlertEmail(35, 'Fear', new Date().toString());
  SpreadsheetApp.getUi().alert('Test email sent! Check your inbox.');
}

/**
 * Create a custom menu when the sheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Fear & Greed')
    .addItem('🔄 Update Latest Value', 'fetchLatestFearGreedIndex')
    .addItem('📧 Test Email Alert', 'testEmailAlert')
    .addSeparator()
    .addItem('⏰ Setup Hourly Check with Alerts', 'setupHourlyTriggerWithAlerts')
    .addItem('🛑 Stop Auto-Updates', 'removeTriggers')
    .addSeparator()
    .addItem('⚙️ Configure Email Settings', 'showEmailSettings')
    .addToUi();
}

/**
 * Show current email settings
 */
function showEmailSettings() {
  const message = `
Current Email Alert Settings:

📧 Recipients: ${EMAIL_RECIPIENTS}
📉 Fear Threshold: < ${FEAR_THRESHOLD}
📈 Greed Threshold: > ${GREED_THRESHOLD}
${ENABLE_EMAIL_ALERTS ? '✅ Alerts: ENABLED' : '❌ Alerts: DISABLED'}

To change settings, edit the configuration at the top of the script.
  `;
  
  SpreadsheetApp.getUi().alert('Email Alert Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Setup hourly trigger with email alerts
 */
function setupHourlyTriggerWithAlerts() {
  // Check if email is configured
  if (EMAIL_RECIPIENTS === 'your-email@example.com') {
    SpreadsheetApp.getUi().alert(
      'Setup Required',
      'Please update EMAIL_RECIPIENTS in the script with your email address before setting up alerts.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Remove existing triggers
  removeTriggers();
  
  // Create new hourly trigger using silent function
  ScriptApp.newTrigger('fetchLatestFearGreedIndexSilent')
    .timeBased()
    .everyHours(1)
    .create();
    
  SpreadsheetApp.getUi().alert(
    'Success!',
    `Hourly monitoring activated!\n\nThe system will check every hour and send email alerts to:\n${EMAIL_RECIPIENTS}\n\nAlerts trigger when:\n• Fear Alert: Value < ${FEAR_THRESHOLD}\n• Greed Alert: Value > ${GREED_THRESHOLD}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Remove all triggers
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  if (triggers.length > 0) {
    SpreadsheetApp.getUi().alert('Auto-updates and alerts stopped.');
  }
}
