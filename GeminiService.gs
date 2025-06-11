// File: GeminiService.gs
// Description: Handles all interactions with the Google Gemini API for
// AI-powered parsing of email content to extract job application details.

// --- GEMINI API PARSING LOGIC ---
function callGemini_forApplicationDetails(emailSubject, emailBody, apiKey) {
  if (!apiKey) {
    Logger.log("[INFO] GEMINI_PARSE: API Key not provided. Skipping Gemini call.");
    return null;
  }
  if ((!emailSubject || emailSubject.trim() === "") && (!emailBody || emailBody.trim() === "")) {
    Logger.log("[WARN] GEMINI_PARSE: Both email subject and body are empty. Skipping Gemini call.");
    return null;
  }

  // Using gemini-1.5-flash-latest as it's good for this kind of task and generally available
  const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
  // Fallback option if Flash gives issues:
  // const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro:generateContent?key=" + apiKey;
  
  Logger.log(`[DEBUG] GEMINI_PARSE: Using API Endpoint: ${API_ENDPOINT.split('key=')[0] + "key=..."}`);

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : ""; // Max 12k chars for body snippet

  // ---- EXPANDED PROMPT ----
  const prompt = `You are a highly specialized AI assistant expert in parsing job application-related emails for a tracking system. Your sole purpose is to analyze the provided email Subject and Body, and extract three key pieces of information: "company_name", "job_title", and "status". You MUST return this information ONLY as a single, valid JSON object, with no surrounding text, explanations, apologies, or markdown.

CRITICAL INSTRUCTIONS - READ AND FOLLOW CAREFULLY:

**PRIORITY 1: Determine Relevance - IS THIS A JOB APPLICATION UPDATE FOR THE RECIPIENT?**
- Your FIRST task is to assess if the email DIRECTLY relates to a job application previously submitted by the recipient, or an update to such an application.
- **IF THE EMAIL IS NOT APPLICATION-RELATED:** This includes general newsletters, marketing or promotional emails, sales pitches, webinar invitations, event announcements, account security alerts, password resets, bills/invoices, platform notifications not tied to a specific submitted application (e.g., "new jobs you might like"), or spam.
    - In such cases, IMMEDIATELY set ALL three fields ("company_name", "job_title", "status") to the exact string "${MANUAL_REVIEW_NEEDED}".
    - Do NOT attempt to extract any information from these irrelevant emails.
    - Your output for these MUST be: {"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${MANUAL_REVIEW_NEEDED}"}

**PRIORITY 2: If Application-Related, Proceed with Extraction:**

1.  "company_name":
    *   **Goal**: Extract the full, official name of the HIRING COMPANY to which the user applied.
    *   **ATS Handling**: Emails often originate from Applicant Tracking Systems (ATS) like Greenhouse (notifications@greenhouse.io), Lever (no-reply@hire.lever.co), Workday, Taleo, iCIMS, Ashby, SmartRecruiters, etc. The sender domain may be the ATS. You MUST identify the actual hiring company mentioned WITHIN the email subject or body. Look for phrases like "Your application to [Hiring Company]", "Careers at [Hiring Company]", "Update from [Hiring Company]", or the company name near the job title.
    *   **Do NOT extract**: The name of the ATS (e.g., "Greenhouse", "Lever"), the name of the job board (e.g., "LinkedIn", "Indeed", "Wellfound" - unless the job board IS the direct hiring company), or generic terms.
    *   **Ambiguity**: If the hiring company name is genuinely unclear from an application context, or only an ATS name is present without the actual company, use "${MANUAL_REVIEW_NEEDED}".
    *   **Accuracy**: Prefer full legal names if available (e.g., "Acme Corporation" over "Acme").

2.  "job_title":
    *   **Goal**: Extract the SPECIFIC job title THE USER APPLIED FOR, as mentioned in THIS email. The title is often explicitly stated after phrases like "your application for...", "application to the position of...", "the ... role", or directly alongside the company name in application submission/viewed confirmations.
    *   **LinkedIn Emails ("Application Sent To..." / "Application Viewed By...")**: These emails (often from sender "LinkedIn") frequently state the company name AND the job title the user applied for directly in the main body or a prominent header within the email content. Scrutinize these carefully for both. Example: "Your application for **Senior Product Manager** was sent to **Innovate Corp**." or "A recruiter from **Innovate Corp** viewed your application for **Senior Product Manager**." Extract "Senior Product Manager".
    *   **ATS Confirmation Emails (e.g., from Greenhouse, Lever)**: These emails confirming receipt of an application (e.g., "We've received your application to [Company]") often DO NOT restate the specific job title within the body of *that specific confirmation email*. If the job title IS NOT restated, you MUST use "${MANUAL_REVIEW_NEEDED}" for the job_title. Do not assume it from the subject line unless the subject clearly states "Your application for [Job Title] at [Company]".
    *   **General Updates/Rejections**: Some updates or rejections may or may not restate the title. If the title of the specific application is not clearly present in THIS email, use "${MANUAL_REVIEW_NEEDED}".
    *   **Strict Rule**: Do NOT infer a job title from company career pages, other listed jobs, or generic phrases like "various roles" unless that phrase directly follows "your application for". Only extract what is stated for THIS specific application event in THIS email. If in doubt, or if only a very generic descriptor like "a role" is used without specifics, prefer "${MANUAL_REVIEW_NEEDED}".

3.  "status":
    *   **Goal**: Determine the current status of the application based on the content of THIS email.
    *   **Strictly Adhere to List**: You MUST choose a status ONLY from the following exact list. Do not invent new statuses or use variations:
        *   "${DEFAULT_STATUS}" (Maps to: Application submitted, application sent, successfully applied, application received - first confirmation)
        *   "${REJECTED_STATUS}" (Maps to: Not moving forward, unfortunately, decided not to proceed, position filled by other candidates, regret to inform)
        *   "${OFFER_STATUS}" (Maps to: Offer of employment, pleased to offer, job offer)
        *   "${INTERVIEW_STATUS}" (Maps to: Invitation to interview, schedule an interview, interview request, like to speak with you)
        *   "${ASSESSMENT_STATUS}" (Maps to: Online assessment, coding challenge, technical test, skills test, take-home assignment)
        *   "${APPLICATION_VIEWED_STATUS}" (Maps to: Application was viewed by recruiter/company, your profile was viewed for the role)
        *   "Update/Other" (Maps to: General updates like "still reviewing applications," "we're delayed," "thanks for your patience," status is mentioned but unclear which of the above it fits best.)
    *   **Exclusion**: "${ACCEPTED_STATUS}" is typically set manually by the user after they accept an offer; do not use it.
    *   **Last Resort**: If the email is clearly job-application-related for the recipient, but the status is absolutely ambiguous and doesn't fit "Update/Other" (very rare), then as a final fallback, use "${MANUAL_REVIEW_NEEDED}" for the status.

**Output Requirements**:
*   **ONLY JSON**: Your entire response must be a single, valid JSON object.
*   **NO Extra Text**: No explanations, greetings, apologies, summaries, or markdown formatting (like \`\`\`json\`\`\`).
*   **Structure**: {"company_name": "...", "job_title": "...", "status": "..."}
*   **Placeholder Usage**: Adhere strictly to using "${MANUAL_REVIEW_NEEDED}" when information is absent or criteria are not met, as instructed for each field.

--- EXAMPLES START ---
Example 1 (LinkedIn "Application Sent To Company - Title Clearly Stated"):
Subject: Francis, your application was sent to MycoWorks
Body: LinkedIn. Your application was sent to MycoWorks. MycoWorks - Emeryville, CA (On-Site). Data Architect/Analyst. Applied on May 16, 2025.
Output:
{"company_name": "MycoWorks","job_title": "Data Architect/Analyst","status": "${DEFAULT_STATUS}"}

Example 2 (Indeed "Application Submitted", title present):
Subject: Indeed Application: Senior Software Engineer
Body: indeed. Application submitted. Senior Software Engineer. Innovatech Solutions - Anytown, USA. The following items were sent to Innovatech Solutions.
Output:
{"company_name": "Innovatech Solutions","job_title": "Senior Software Engineer","status": "${DEFAULT_STATUS}"}

Example 3 (Rejection from ATS, title might be in subject, but not confirmed in this email body):
Subject: Update on your application for Product Manager at MegaEnterprises
Body: From: no-reply@greenhouse.io. Dear Applicant, Thank you for your interest in MegaEnterprises. After careful consideration, we have decided to move forward with other candidates for this position.
Output:
{"company_name": "MegaEnterprises","job_title": "Product Manager","status": "${REJECTED_STATUS}"} 
(Self-correction: Title "Product Manager" taken from subject if directly linked to "your application". If subject was generic like "Application Update", job_title would be ${MANUAL_REVIEW_NEEDED})

Example 4 (Interview Invitation via ATS, title present):
Subject: Invitation to Interview: Data Analyst at Beta Innovations (via Lever)
Body: We were impressed with your application for the Data Analyst role and would like to invite you to an interview...
Output:
{"company_name": "Beta Innovations","job_title": "Data Analyst","status": "${INTERVIEW_STATUS}"}

Example 5 (ATS Email - Application Received, NO specific title in THIS email body):
Subject: Thank you for applying to Handshake!
Body: no-reply@greenhouse.io. Hi Francis, Thank you for your interest in Handshake! We have received your application and will be reviewing your background shortly... Handshake Recruiting.
Output:
{"company_name": "Handshake","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${DEFAULT_STATUS}"}

Example 6 (Unrelated Marketing):
Subject: Join our webinar on Future Tech!
Body: Hi User, Don't miss out on our exclusive webinar...
Output:
{"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${MANUAL_REVIEW_NEEDED}"}

Example 7 (LinkedIn "Application Viewed By..." - Title Clearly Stated):
Subject: Your application was viewed by Gotham Technology Group
Body: LinkedIn. Great job getting noticed by the hiring team at Gotham Technology Group. Gotham Technology Group - New York, United States. Business Analyst/Product Manager. Applied on May 14.
Output:
{"company_name": "Gotham Technology Group","job_title": "Business Analyst/Product Manager","status": "${APPLICATION_VIEWED_STATUS}"}

Example 8 (Wellfound "Application Submitted" - Often has title):
Subject: Application to LILT successfully submitted
Body: wellfound. Your application to LILT for the position of Lead Product Manager has been submitted! View your application. LILT.
Output:
{"company_name": "LILT","job_title": "Lead Product Manager","status": "${DEFAULT_STATUS}"}

Example 9 (Email indicating general interest/no specific role or company clear):
Subject: An interesting opportunity
Body: Hi Francis, Your profile on LinkedIn matches an opening we have. Would you be open to a quick chat? Regards, Recruiter.
Output:
{"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "Update/Other"}
--- EXAMPLES END ---

--- START OF EMAIL TO PROCESS ---
Subject: ${emailSubject}
Body:
${bodySnippet}
--- END OF EMAIL TO PROCESS ---
Output JSON:
`; // End of prompt template literal
  // ---- END OF EXPANDED PROMPT ----

  const payload = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": {
      "temperature": 0.2, 
      "maxOutputTokens": 512, 
      "topP": 0.95, 
      "topK": 40
    },
    "safetySettings": [ 
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const options = {'method':'post', 'contentType':'application/json', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};

  if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE: Calling API for subj: "${emailSubject.substring(0,100)}". Prompt len (approx): ${prompt.length}`);
  let response; let attempt = 0; const maxAttempts = 2;

  while(attempt < maxAttempts){
    attempt++;
    try {
      response = UrlFetchApp.fetch(API_ENDPOINT, options);
      const responseCode = response.getResponseCode(); const responseBody = response.getContentText();
      if(DEBUG_MODE) Logger.log(`[DEBUG] GEMINI_PARSE (Attempt ${attempt}): RC: ${responseCode}. Body(start): ${responseBody.substring(0,200)}`);

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
          let extractedJsonString = jsonResponse.candidates[0].content.parts[0].text.trim();
          // Clean potential markdown code block formatting from the response
          if (extractedJsonString.startsWith("```json")) extractedJsonString = extractedJsonString.substring(7).trim();
          if (extractedJsonString.startsWith("```")) extractedJsonString = extractedJsonString.substring(3).trim();
          if (extractedJsonString.endsWith("```")) extractedJsonString = extractedJsonString.substring(0, extractedJsonString.length - 3).trim();
          
          if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE: Cleaned JSON from API: ${extractedJsonString}`);
          try {
            const extractedData = JSON.parse(extractedJsonString);
            // Basic validation that all expected keys are present
            if (typeof extractedData.company_name !== 'undefined' && 
                typeof extractedData.job_title !== 'undefined' && 
                typeof extractedData.status !== 'undefined') {
              Logger.log(`[INFO] GEMINI_PARSE: Success. C:"${extractedData.company_name}", T:"${extractedData.job_title}", S:"${extractedData.status}"`);
              return {
                  company: extractedData.company_name || MANUAL_REVIEW_NEEDED, 
                  title: extractedData.job_title || MANUAL_REVIEW_NEEDED, 
                  status: extractedData.status || MANUAL_REVIEW_NEEDED
              };
            } else {
              Logger.log(`[WARN] GEMINI_PARSE: JSON from Gemini missing one or more expected fields. Output: ${extractedJsonString}`);
              return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED}; // Fallback
            }
          } catch (e) {
            Logger.log(`[ERROR] GEMINI_PARSE: Error parsing JSON string from Gemini: ${e.toString()}\nOffending String: >>>${extractedJsonString}<<<`);
            return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED}; // Fallback
          }
        } else if (jsonResponse.promptFeedback?.blockReason) {
          Logger.log(`[ERROR] GEMINI_PARSE: Prompt blocked by API. Reason: ${jsonResponse.promptFeedback.blockReason}. Details: ${JSON.stringify(jsonResponse.promptFeedback.safetyRatings)}`);
          return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:`Blocked: ${jsonResponse.promptFeedback.blockReason}`}; // Include block reason in status for debugging
        } else {
          Logger.log(`[ERROR] GEMINI_PARSE: API response structure unexpected (no candidates/text part). Full Body (first 500 chars): ${responseBody.substring(0,500)}`);
          return null; 
        }
      } else if (responseCode === 429) { // Rate limit
        Logger.log(`[WARN] GEMINI_PARSE: Rate limit (HTTP 429). Attempt ${attempt}/${maxAttempts}. Waiting...`);
        if (attempt < maxAttempts) { Utilities.sleep(5000 + Math.floor(Math.random() * 5000)); continue; }
        else { Logger.log(`[ERROR] GEMINI_PARSE: Max retry attempts reached for rate limit.`); return null; }
      } else { // Other API errors (400, 404 model not found, 500, etc.)
        Logger.log(`[ERROR] GEMINI_PARSE: API call returned HTTP error. Code: ${responseCode}. Body (first 500 chars): ${responseBody.substring(0,500)}`);
        // Specific check for 404 model error, in case it switches during operation
        if (responseCode === 404 && responseBody.includes("is not found for API version")) {
            Logger.log(`[FATAL] GEMINI_MODEL_ERROR: The model specified (${API_ENDPOINT.split('/models/')[1].split(':')[0]}) may no longer be valid or available. Check model name and API version.`)
        }
        return null; // Indicates failure to parse
      }
    } catch (e) { // Catch network errors or other exceptions during UrlFetchApp.fetch
      Logger.log(`[ERROR] GEMINI_PARSE: Exception during API call (Attempt ${attempt}): ${e.toString()}\nStack: ${e.stack}`);
      if (attempt < maxAttempts) { Utilities.sleep(3000); continue; } // Basic retry wait
      return null;
    }
  }
  Logger.log(`[ERROR] GEMINI_PARSE: Failed to get a valid response from Gemini API after ${maxAttempts} attempts.`);
  return null; // Fallback if all attempts fail
}

/**
 * Calls the Gemini API with a prompt specifically designed to extract multiple job leads from email content.
 * @param {string} emailBody The plain text body of the email.
 * @param {string} apiKey The Gemini API key.
 * @return {{success: boolean, data: Object|null, error: string|null}} Result object including success status,
 *                                                                  API response data (if successful), or an error message.
 */

function callGemini_forJobLeads(emailBody, apiKey) {
    if (typeof emailBody !== 'string') {
        Logger.log(`[GEMINI_LEADS CRITICAL ERR] callGemini_forJobLeads: emailBody is not a string. Type: ${typeof emailBody}`);
        return { success: false, data: null, error: `emailBody is not a string.` };
    }
    // Logger.log(`[GEMINI_LEADS API CALL] Body length: ${emailBody.length}. Start: "${emailBody.substring(0, 70)}..."`);

    const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
    // Fallback model if needed:
    // const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro:generateContent?key=" + apiKey;

    // MOCK RESPONSE LOGIC (for testing without a real API key or to simulate responses)
    // Ensure this logic is bypassed if a valid apiKey is provided.
    if (!apiKey || apiKey === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || apiKey.trim() === '') {
        Logger.log("[GEMINI_LEADS WARN STUB] API Key is placeholder, empty, or null. Using MOCK response for job leads.");
        if (emailBody.toLowerCase().includes("multiple job listings inside") || emailBody.toLowerCase().includes("software engineer at google")) {
            return {
                success: true,
                data: {
                    candidates: [{
                        content: {
                            parts: [{
                                text: JSON.stringify([
                                    { "jobTitle": "Software Engineer (Mock)", "company": "Tech Alpha (Mock)", "location": "Remote", "linkToJobPosting": "https://example.com/job/alpha" },
                                    { "jobTitle": "Product Manager (Mock)", "company": "Innovate Beta (Mock)", "location": "New York, NY", "linkToJobPosting": "https://example.com/job/beta" }
                                ])
                            }]
                        }
                    }]
                },
                error: null
            };
        }
        // Default mock for other cases if API key is missing
        return {
            success: true,
            data: {
                candidates: [{
                    content: {
                        parts: [{
                            text: JSON.stringify([{ "jobTitle": "N/A (Mock Single)", "company": "Some Corp (Mock)", "location": "Remote", "linkToJobPosting": "N/A" }])
                        }]
                    }
                }]
            },
            error: null
        };
    }
    // END MOCK RESPONSE LOGIC

    const promptText = `From the following email content, identify each distinct job posting.
For EACH job posting found, extract the following details:
- Job Title
- Company
- Location (city, state, or remote)
- Link (a direct URL to the job application or description, if available)

If a field for a specific job is not found or not applicable, use the string "N/A" as its value.

Format your entire response as a single, valid JSON array where each element of the array is a JSON object representing one job posting.
Each JSON object should have the keys: "jobTitle", "company", "location", "linkToJobPosting".
If no job postings are found, return an empty JSON array: [].
Do not include any text before or after the JSON array. Ensure the JSON is strictly valid.

Example of a single job object:
{
  "jobTitle": "Software Engineer",
  "company": "Tech Corp",
  "location": "Remote",
  "linkToJobPosting": "https://example.com/job/123"
}

Email Content:
---
${emailBody.substring(0, 28000)} 
---
JSON Array Output:`;

    const payload = {
        contents: [{ parts: [{ "text": promptText }] }],
        generationConfig: {
            temperature: 0.2,
            maxOutputTokens: 8192, // Max for gemini-1.5-flash
        },
        safetySettings: [
            { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
            { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
            { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
            { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
        ]
    };

    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true // Important to handle errors manually
    };

    let attempt = 0;
    const maxAttempts = 2; // Simple retry mechanism for transient issues like rate limits

    while (attempt < maxAttempts) {
        attempt++;
        try {
            // Logger.log(`[GEMINI_LEADS API REQUEST Attempt ${attempt}] URL: ${API_ENDPOINT.split('key=')[0]}key=...`);
            const response = UrlFetchApp.fetch(API_ENDPOINT, options);
            const responseCode = response.getResponseCode();
            const responseBody = response.getContentText();

            if (responseCode === 200) {
                // Logger.log(`[GEMINI_LEADS API SUCCESS ${responseCode}] Raw response (first 300 chars): ${responseBody.substring(0, 300)}...`);
                try {
                    return { success: true, data: JSON.parse(responseBody), error: null };
                } catch (jsonParseError) {
                    Logger.log(`[GEMINI_LEADS API ERROR] Failed to parse Gemini JSON response: ${jsonParseError}. Raw body: ${responseBody}`);
                    return { success: false, data: null, error: `Failed to parse API JSON response: ${jsonParseError}. Response: ${responseBody}` };
                }
            } else if (responseCode === 429 && attempt < maxAttempts) { // Rate limit with retries left
                Logger.log(`[GEMINI_LEADS API WARN ${responseCode}] Rate limit hit on attempt ${attempt}. Waiting before retry...`);
                Utilities.sleep(3000 + Math.random() * 2000); // Wait 3-5 seconds
                continue; // Go to next attempt
            } else { // Other API errors (400, 500, or 429 on last attempt)
                Logger.log(`[GEMINI_LEADS API ERROR ${responseCode}] Full error response: ${responseBody}`);
                if (responseCode === 400 && responseBody.toLowerCase().includes("api key not valid")) {
                    Logger.log(`[GEMINI_LEADS FATAL] API Key is not valid. Please check the key stored in UserProperties via key name: "${GEMINI_API_KEY_PROPERTY}".`);
                } else if (responseCode === 404 && responseBody.toLowerCase().includes("is not found for api version")) {
                    Logger.log(`[GEMINI_LEADS FATAL] The model specified in API_ENDPOINT may no longer be valid or available. Check model name and API version.`);
                }
                return { success: false, data: null, error: `API Error ${responseCode}: ${responseBody}` };
            }
        } catch (e) { // Catch network errors or other exceptions during UrlFetchApp.fetch
            Logger.log(`[GEMINI_LEADS API CATCH ERROR Attempt ${attempt}] Failed to call Gemini API: ${e.toString()}`);
            if (attempt < maxAttempts) {
                Utilities.sleep(2000); // Basic wait before retry on fetch error
                continue;
            }
            return { success: false, data: null, error: `UrlFetchApp Fetch Error after ${maxAttempts} attempts: ${e.toString()}` };
        }
    }
    // Should not be reached if loop logic is correct, but as a fallback:
    return { success: false, data: null, error: `Exceeded max retry attempts for Gemini API call.` };
}

/**
 * Parses the raw JSON data from the Gemini API response (for job leads) into an array of job objects.
 * @param {Object} apiResponseData The 'data' part of the successful response from callGemini_forJobLeads.
 * @return {Array<Object>} An array of job lead objects, or an empty array if parsing fails or no jobs are found.
 */
function parseGeminiResponse_forJobLeads(apiResponseData) {
    // Logger.log(`[GEMINI_LEADS PARSE] Raw API Data from Gemini (first 300 chars): ${JSON.stringify(apiResponseData).substring(0, 300)}...`);
    let jobListings = [];
    try {
        let jsonStringFromLLM = "";
        if (apiResponseData &&
            apiResponseData.candidates &&
            apiResponseData.candidates.length > 0 &&
            apiResponseData.candidates[0].content &&
            apiResponseData.candidates[0].content.parts &&
            apiResponseData.candidates[0].content.parts.length > 0 &&
            typeof apiResponseData.candidates[0].content.parts[0].text === 'string') {

            jsonStringFromLLM = apiResponseData.candidates[0].content.parts[0].text.trim();
            
            // Clean potential markdown code block formatting
            if (jsonStringFromLLM.startsWith("```json")) {
                jsonStringFromLLM = jsonStringFromLLM.substring(7).trim();
            } else if (jsonStringFromLLM.startsWith("```")) {
                jsonStringFromLLM = jsonStringFromLLM.substring(3).trim();
            }
            if (jsonStringFromLLM.endsWith("```")) {
                jsonStringFromLLM = jsonStringFromLLM.substring(0, jsonStringFromLLM.length - 3).trim();
            }
            
            // Logger.log(`[GEMINI_LEADS PARSE] Cleaned JSON String (first 500 chars): ${jsonStringFromLLM.substring(0, 500)}...`);
            // For full string debugging: Logger.log(`[GEMINI_LEADS PARSE] FULL Cleaned JSON String from LLM: ${jsonStringFromLLM}`);
        } else {
            Logger.log(`[GEMINI_LEADS PARSE WARN] No parsable content string found in Gemini response structure for leads.`);
            if (apiResponseData && apiResponseData.promptFeedback && apiResponseData.promptFeedback.blockReason) {
                Logger.log(`[GEMINI_LEADS PARSE WARN] Prompt Feedback Block Reason: ${apiResponseData.promptFeedback.blockReason}. Safety Ratings: ${JSON.stringify(apiResponseData.promptFeedback.safetyRatings)}`);
            }
            return jobListings; // Return empty array
        }

        try {
            const parsedData = JSON.parse(jsonStringFromLLM);
            if (Array.isArray(parsedData)) {
                parsedData.forEach(job => {
                    if (job && typeof job === 'object' && (job.jobTitle || job.company)) { // Basic validation for a job object
                        jobListings.push({
                            jobTitle: job.jobTitle || "N/A",
                            company: job.company || "N/A",
                            location: job.location || "N/A",
                            linkToJobPosting: job.linkToJobPosting || "N/A"
                        });
                    } else {
                        Logger.log(`[GEMINI_LEADS PARSE WARN] Skipped an invalid item in the job listings array: ${JSON.stringify(job)}`);
                    }
                });
                // Logger.log(`[GEMINI_LEADS PARSE] Successfully parsed ${jobListings.length} job objects from JSON string.`);
            } else if (typeof parsedData === 'object' && parsedData !== null && (parsedData.jobTitle || parsedData.company)) {
                // Handle case where LLM might return a single object instead of an array with one object
                jobListings.push({
                    jobTitle: parsedData.jobTitle || "N/A",
                    company: parsedData.company || "N/A",
                    location: parsedData.location || "N/A",
                    linkToJobPosting: parsedData.linkToJobPosting || "N/A"
                });
                Logger.log(`[GEMINI_LEADS PARSE WARN] LLM output was a single object, not an array. Parsed as one job.`);
            } else {
                Logger.log(`[GEMINI_LEADS PARSE WARN] LLM output was not a JSON array or a parsable single job object. Output (first 200 chars): ${jsonStringFromLLM.substring(0, 200)}`);
            }
        } catch (jsonError) {
            Logger.log(`[GEMINI_LEADS PARSE ERROR] Failed to parse JSON string from LLM: ${jsonError}. String (first 500 chars): ${jsonStringFromLLM.substring(0, 500)}`);
        }
        return jobListings;
    } catch (e) {
        Logger.log(`[GEMINI_LEADS PARSE ERROR] Outer error during parsing leads response: ${e.toString()}. API Response Data (first 300 chars): ${JSON.stringify(apiResponseData).substring(0, 300)}`);
        return jobListings; // Return empty array on error
    }
}
