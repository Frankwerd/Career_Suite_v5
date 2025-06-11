// File: ParsingUtils.gs
// Description: Contains functions dedicated to parsing email content (subject, body, sender)
// using regular expressions and keyword matching to extract job application details.

// --- REGEX PARSING LOGIC (FALLBACK) ---
function parseCompanyFromDomain(sender) {
  const emailMatch = sender.match(/<([^>]+)>/); if (!emailMatch || !emailMatch[1]) return null;
  const emailAddress = emailMatch[1]; const domainParts = emailAddress.split('@'); if (domainParts.length !== 2) return null;
  let domain = domainParts[1].toLowerCase();
  if (IGNORED_DOMAINS.has(domain) && !domain.includes('wellfound.com') /*Exception for wellfound which can be an ATS domain*/ ) { return null; }
  domain = domain.replace(/^(?:careers|jobs|recruiting|apply|hr|talent|notification|notifications|team|hello|no-reply|noreply)[.-]?/i, '');
  domain = domain.replace(/\.(com|org|net|io|co|ai|dev|xyz|tech|ca|uk|de|fr|app|eu|us|info|biz|work|agency|careers|招聘|group|global|inc|llc|ltd|corp|gmbh)$/i, ''); // More comprehensive TLDs
  domain = domain.replace(/[^a-z0-9]+/gi, ' '); // Replace non-alphanumeric with space
  domain = domain.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); // Capitalize
  return domain.trim() || null;
}

function parseCompanyFromSenderName(sender) {
  const nameMatch = sender.match(/^"?(.*?)"?\s*</);
  let name = nameMatch ? nameMatch[1].trim() : sender.split('<')[0].trim();
  if (!name || name.includes('@') || name.length < 2) return null; // Basic validation

  // Remove common ATS/platform noise
  name = name.replace(/\|\s*(?:greenhouse|lever|wellfound|workday|ashby|icims|smartrecruiters|taleo|bamboohr|recruiterbox|jazzhr|workable|breezyhr|notion)\b/i, '');
  name = name.replace(/\s*(?:via Wellfound|via LinkedIn|via Indeed|from Greenhouse|from Lever|Careers at|Hiring at)\b/gi, '');
  // Remove common generic terms often appended
  name = name.replace(/\s*(?:Careers|Recruiting|Recruitment|Hiring Team|Hiring|Talent Acquisition|Talent|HR|Team|Notifications?|Jobs?|Updates?|Apply|Notification|Hello|No-?Reply|Support|Info|Admin|Department|Notifications)\b/gi, '');
  // Remove trailing legal entities, punctuation, and trim
  name = name.replace(/[|,_.\s]+(?:Inc\.?|LLC\.?|Ltd\.?|Corp\.?|GmbH|Solutions|Services|Group|Global|Technologies|Labs|Studio|Ventures)?$/i, '').trim();
  name = name.replace(/^(?:The|A)\s+/i, '').trim(); // Remove leading "The", "A"
  
  if (name.length > 1 && !/^(?:noreply|no-reply|jobs|careers|support|info|admin|hr|talent|recruiting|team|hello)$/i.test(name.toLowerCase())) {
    return name;
  }
  return null;
}

function extractCompanyAndTitle(message, platform, emailSubject, plainBody) {
  let company = MANUAL_REVIEW_NEEDED; let title = MANUAL_REVIEW_NEEDED;
  const sender = message.getFrom();
  if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Fallback C/T for subj: "${emailSubject.substring(0,100)}"`);
  
  let tempCompanyFromDomain = parseCompanyFromDomain(sender);
  let tempCompanyFromName = parseCompanyFromSenderName(sender);
  if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: From Sender -> Name="${tempCompanyFromName}", Domain="${tempCompanyFromDomain}"`);

  // Platform-specific logic (example for Wellfound)
  if (platform === "Wellfound" && plainBody) {
    let wfCoSub = emailSubject.match(/update from (.*?)(?: \|| at |$)/i) || emailSubject.match(/application to (.*?)(?: successfully| at |$)/i) || emailSubject.match(/New introduction from (.*?)(?: for |$)/i);
    if (wfCoSub && wfCoSub[1]) company = wfCoSub[1].trim();
    // More specific Wellfound body parsing can be added here if needed
    if (title === MANUAL_REVIEW_NEEDED && plainBody && sender.toLowerCase().includes("team@hi.wellfound.com")) {
        const markerPhrase = "if there's a match, we will make an email introduction.";
        const markerIndex = plainBody.toLowerCase().indexOf(markerPhrase);
        if (markerIndex !== -1) {
            const relevantText = plainBody.substring(markerIndex + markerPhrase.length);
            const titleMatch = relevantText.match(/^\s*\*\s*([A-Za-z\s.,:&'\/-]+?)(?:\s*\(| at | \n|$)/m);
            if (titleMatch && titleMatch[1]) title = titleMatch[1].trim();
        }
    }
  }

  // Regex patterns for subject line parsing (ordered by specificity or commonness)
  const complexPatterns = [
    { r: /Application for(?: the)?\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 }, // Title at Company
    { r: /Invite(?:.*?)(?:to|for)(?: an)? interview(?:.*?)\sfor\s+(?:the\s)?(.+?)(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 }, // Interview for Title at Company
    { r: /Your application for(?: the)?\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /Regarding your application for\s+(.+?)(?:\s-\s(.*?))?(?:\s@\s(.*?))?$/i, tI: 1, cI: 3, cI2: 2}, // Greenhouse subjects
    { r: /^(?:Update on|Your Application to|Thank you for applying to)\s+([^-:|–—]+?)(?:\s*-\s*([^-:|–—]+))?$/i, cI: 1, tI: 2 }, // Lever style: Company - Title OR Title - Company
    { r: /applying to\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /interest in the\s+(.+?)\s+role(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 },
    { r: /update on your\s+(.+?)\s+app(?:lication)?(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 }
  ];

  for (const pI of complexPatterns) {
    let m = emailSubject.match(pI.r);
    if (m) {
      if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Matched subject pattern: ${pI.r}`);
      let extractedTitle = pI.tI > 0 && m[pI.tI] ? m[pI.tI].trim() : null;
      let extractedCompany = pI.cI > 0 && m[pI.cI] ? m[pI.cI].trim() : null;
      if (!extractedCompany && pI.cI2 > 0 && m[pI.cI2]) extractedCompany = m[pI.cI2].trim();

      // Check if one part looks more like a company and the other a title, for Lever-style ambiguous subjects
      if (pI.cI === 1 && pI.tI === 2 && extractedCompany && extractedTitle) { // e.g. "Lever: Company - Title" or "Lever: Title - Company"
          // A simple heuristic: if one contains "Engineer", "Manager", "Analyst", etc. it's likely the title
          if (/\b(engineer|manager|analyst|developer|specialist|lead|director|coordinator|architect|consultant|designer|recruiter|associate|intern)\b/i.test(extractedCompany) &&
             !/\b(engineer|manager|analyst|developer|specialist|lead|director|coordinator|architect|consultant|designer|recruiter|associate|intern)\b/i.test(extractedTitle)) {
              // Swap them
              [extractedCompany, extractedTitle] = [extractedTitle, extractedCompany];
              if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Swapped Company/Title based on keywords. C: ${extractedCompany}, T: ${extractedTitle}`);
          }
      }
      
      if (extractedTitle && (title === MANUAL_REVIEW_NEEDED || title === DEFAULT_STATUS)) title = extractedTitle;
      if (extractedCompany && (company === MANUAL_REVIEW_NEEDED || company === DEFAULT_PLATFORM)) company = extractedCompany;
      
      if (company !== MANUAL_REVIEW_NEEDED && title !== MANUAL_REVIEW_NEEDED && company !== DEFAULT_PLATFORM && title !== DEFAULT_STATUS) break;
    }
  }

  if (company === MANUAL_REVIEW_NEEDED && tempCompanyFromName) company = tempCompanyFromName; // Use parsed sender name if subject failed for company

  // Body Scan Fallback (if still needed)
  if ((company === MANUAL_REVIEW_NEEDED || title === MANUAL_REVIEW_NEEDED || company === DEFAULT_PLATFORM || title === DEFAULT_STATUS) && plainBody) {
    const bodyCleaned = plainBody.substring(0, 1000).replace(/<[^>]+>/g, ' '); // Process first 1000 chars
    if (company === MANUAL_REVIEW_NEEDED || company === DEFAULT_PLATFORM) {
      let bodyCompanyMatch = bodyCleaned.match(/(?:applying to|application with|interview with|position at|role at|opportunity at|Thank you for your interest in working at)\s+([A-Z][A-Za-z\s.&'-]+(?:LLC|Inc\.?|Ltd\.?|Corp\.?|GmbH|Group|Solutions|Technologies)?)(?:[.,\s\n\(]|$)/i);
      if (bodyCompanyMatch && bodyCompanyMatch[1]) company = bodyCompanyMatch[1].trim();
    }
    if (title === MANUAL_REVIEW_NEEDED || title === DEFAULT_STATUS) {
      let bodyTitleMatch = bodyCleaned.match(/(?:application for the|position of|role of|applying for the|interview for the|title:)\s+([A-Za-z][A-Za-z0-9\s.,:&'\/\(\)-]+?)(?:\s\(| at | with |[\s.,\n\(]|$)/i);
      if (bodyTitleMatch && bodyTitleMatch[1]) title = bodyTitleMatch[1].trim();
    }
  }

  if (company === MANUAL_REVIEW_NEEDED && tempCompanyFromDomain) company = tempCompanyFromDomain; // Last resort for company

  const cleanE = (entity, isTitle = false) => {
    if (!entity || entity === MANUAL_REVIEW_NEEDED || entity === DEFAULT_STATUS || entity === DEFAULT_PLATFORM || entity.toLowerCase() === "n/a") return MANUAL_REVIEW_NEEDED;
    let cl = entity.split(/[\n\r#(]| - /)[0]; // Take first line, remove text after # or some " - " patterns
    cl = cl.replace(/ (?:inc|llc|ltd|corp|gmbh)[\.,]?$/i, '').replace(/[,"']?$/, '');
    cl = cl.replace(/^(?:The|A)\s+/i, '');
    cl = cl.replace(/\s+/g, ' ').trim();
    if (isTitle) {
        cl = cl.replace(/JR\d+\s*[-–—]?\s*/i, '');
        cl = cl.replace(/\(Senior\)/i, 'Senior'); // Hoist (Senior)
        cl = cl.replace(/\(.*?(?:remote|hybrid|onsite|contract|part-time|full-time|intern|co-op|stipend|urgent|hiring|opening|various locations).*?\)/gi, '');
        cl = cl.replace(/[-–—:]\s*(?:remote|hybrid|onsite|contract|part-time|full-time|intern|co-op|various locations)\s*$/gi, '');
        cl = cl.replace(/^[-\s#*]+|[,\s]+$/g, ''); // Clean leading/trailing punctuation and spaces
    }
    cl = cl.replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'").replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"').replace(/&/gi, '&').replace(/ /gi, ' ');
    cl = cl.trim();
    return cl.length < 2 ? MANUAL_REVIEW_NEEDED : cl; // If cleaning results in too short, mark as N/A
  };
    
  company = cleanE(company);
  title = cleanE(title, true);

  if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Final Fallback Result -> Company:"${company}", Title:"${title}"`);
  return {company: company, title: title};
}

function parseBodyForStatus(plainBody) {
  if (!plainBody || plainBody.length < 10) { if (DEBUG_MODE) Logger.log("[DEBUG] RGX_STATUS: Body too short/missing."); return null; }
  let bL = plainBody.toLowerCase().replace(/[.,!?;:()\[\]{}'"“”‘’\-–—]/g, ' ').replace(/\s+/g, ' ').trim();
  if (OFFER_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: OFFER.`); return OFFER_STATUS; }
  if (INTERVIEW_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: INTERVIEW.`); return INTERVIEW_STATUS; }
  if (ASSESSMENT_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: ASSESSMENT.`); return ASSESSMENT_STATUS; }
  if (APPLICATION_VIEWED_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: APP_VIEWED.`); return APPLICATION_VIEWED_STATUS; }
  if (REJECTION_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: REJECTION.`); return REJECTED_STATUS; }
  if (DEBUG_MODE) Logger.log("[DEBUG] RGX_STATUS: No specific keywords found by regex.");
  return null;
}
