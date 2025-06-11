// File: Config.gs
// Description: Contains all global configuration constants for the "Master Job Manager" project,
// covering both the Job Application Tracker and the Job Leads Tracker modules.
// Modifying values here will change the behavior of the script.

// --- General Configuration ---
const DEBUG_MODE = true; // Set to false for production to reduce logging

// --- Gemini AI Configuration ---
const GEMINI_API_KEY_PROPERTY = 'GEMINI_API_KEY'; // UserProperty key for storing the Gemini API Key (used by both modules)

// --- Spreadsheet Configuration (Main Workbook) ---
// This is the primary spreadsheet where all data (applications, leads, dashboard) will reside.
const FIXED_SPREADSHEET_ID = ""; // IMPORTANT: Set this to YOUR actual spreadsheet ID, or leave blank to find/create by name
const TARGET_SPREADSHEET_FILENAME = "Automated Job Application Tracker Data"; // Used if FIXED_SPREADSHEET_ID is blank

// --- Tab Names within the Main Spreadsheet ---
// For Job Application Tracker Module
const APP_TRACKER_SHEET_TAB_NAME = "Applications";      // Data sheet for actual job applications
const DASHBOARD_TAB_NAME = "Dashboard";                 // Dashboard sheet for application analytics
const HELPER_SHEET_NAME = "DashboardHelperData";        // Hidden helper sheet for dashboard charts

// For Job Leads Tracker Module
const LEADS_SHEET_TAB_NAME = "Potential Job Leads";     // New tab for incoming job leads

// --- Gmail Label Configuration (Hierarchical) ---
const MASTER_GMAIL_LABEL_PARENT = "Master Job Manager"; // The new top-level parent

// Labels for the Main Job Application Tracker module
const TRACKER_GMAIL_LABEL_PARENT = MASTER_GMAIL_LABEL_PARENT + "/Job Application Tracker";
const TRACKER_GMAIL_LABEL_TO_PROCESS = TRACKER_GMAIL_LABEL_PARENT + "/To Process";
const TRACKER_GMAIL_LABEL_PROCESSED = TRACKER_GMAIL_LABEL_PARENT + "/Processed";
const TRACKER_GMAIL_LABEL_MANUAL_REVIEW = TRACKER_GMAIL_LABEL_PARENT + "/Manual Review Needed";

// New Filter Query for Job Application Tracker module (to catch application updates)
const TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES = 'subject:("your application" OR "application to" OR "application for" OR "application update" OR "thank you for applying" OR "thanks for applying" OR "thank you for your interest" OR "received your application")';

// Labels for the Job Leads Tracker module
const LEADS_GMAIL_LABEL_PARENT = MASTER_GMAIL_LABEL_PARENT + "/Job Application Potential";
const LEADS_GMAIL_LABEL_NEEDS_PROCESS = LEADS_GMAIL_LABEL_PARENT + "/NeedsProcess";
const LEADS_GMAIL_LABEL_DONE_PROCESS = LEADS_GMAIL_LABEL_PARENT + "/DoneProcess";
const LEADS_GMAIL_FILTER_QUERY = 'subject:("job alert") OR subject:(jobalert)'; // Gmail filter query for leads

// --- Job Application Tracker: Column Indices (1-based) for "Applications" Sheet ---
const PROCESSED_TIMESTAMP_COL = 1;
const EMAIL_DATE_COL = 2;
const PLATFORM_COL = 3;
const COMPANY_COL = 4;
const JOB_TITLE_COL = 5;
const STATUS_COL = 6;
const PEAK_STATUS_COL = 7;
const LAST_UPDATE_DATE_COL = 8;
const EMAIL_SUBJECT_COL = 9;
const EMAIL_LINK_COL = 10;
const EMAIL_ID_COL = 11;
const TOTAL_COLUMNS_IN_APP_SHEET = 11; // Total columns in the "Applications" sheet

// --- Job Application Tracker: Status Values & Hierarchy ---
const DEFAULT_STATUS = "Applied";
const REJECTED_STATUS = "Rejected";
const OFFER_STATUS = "Offer Received";
const ACCEPTED_STATUS = "Offer Accepted";
const INTERVIEW_STATUS = "Interview Scheduled";
const ASSESSMENT_STATUS = "Assessment/Screening";
const APPLICATION_VIEWED_STATUS = "Application Viewed";
const MANUAL_REVIEW_NEEDED = "N/A - Manual Review"; // Used if parsing fails or for manual checks
const DEFAULT_PLATFORM = "Other"; // Default platform if not detected

const STATUS_HIERARCHY = {
  [MANUAL_REVIEW_NEEDED]: -1,
  "Update/Other": 0,
  [DEFAULT_STATUS]: 1,
  [APPLICATION_VIEWED_STATUS]: 2,
  [ASSESSMENT_STATUS]: 3,
  [INTERVIEW_STATUS]: 4,
  [OFFER_STATUS]: 5, // Offer and Rejected can be at the same level if offer is not the final "win" state before acceptance
  [REJECTED_STATUS]: 5,
  [ACCEPTED_STATUS]: 6
};

// --- Job Application Tracker: Config for Auto-Reject Stale Applications ---
const WEEKS_THRESHOLD = 7; // Number of weeks after which a non-finalized application is considered stale
const FINAL_STATUSES_FOR_STALE_CHECK = new Set([REJECTED_STATUS, ACCEPTED_STATUS, "Withdrawn"]); // Statuses exempt from stale check

// --- Job Application Tracker: Keywords for Email Parsing (Regex Fallback) ---
const REJECTION_KEYWORDS = ["unfortunately", "regret to inform", "not moving forward", "decided not to proceed", "other candidates", "filled the position", "thank you for your time but"];
const OFFER_KEYWORDS = ["pleased to offer", "offer of employment", "job offer", "formally offer you the position"];
const INTERVIEW_KEYWORDS = ["invitation to interview", "schedule an interview", "interview request", "like to speak with you", "next steps involve an interview", "interview availability"];
const ASSESSMENT_KEYWORDS = ["assessment", "coding challenge", "online test", "technical screen", "next step is a skill assessment", "take a short test"];
const APPLICATION_VIEWED_KEYWORDS = ["application was viewed", "your application was viewed by", "recruiter viewed your application", "company viewed your application", "viewed your profile for the role"];
const PLATFORM_DOMAIN_KEYWORDS = { "linkedin.com": "LinkedIn", "indeed.com": "Indeed", "wellfound.com": "Wellfound", "angel.co": "Wellfound" };
const IGNORED_DOMAINS = new Set(['greenhouse.io', 'lever.co', 'myworkday.com', 'icims.com', 'ashbyhq.com', 'smartrecruiters.com', 'bamboohr.com', 'taleo.net', 'gmail.com', 'google.com', 'example.com']);

// --- Job Leads Tracker: UserProperty KEYS for storing specific Label IDs ---
// These are used by the Leads Tracker setup to create Gmail filters, as filters require label IDs.
// While label names could be used to fetch IDs on the fly, the original script stored them.
const LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID = 'leadsGmailNeedsProcessLabelId';
const LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID = 'leadsGmailDoneProcessLabelId';
// Storing label names in UserProperties is optional if deriving from constants above is preferred.
// const LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_NAME = 'leadsGmailNeedsProcessLabelName';
// const LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_NAME = 'leadsGmailDoneProcessLabelName';

// --- Column Headers for "Potential Job Leads" Sheet (Used by Leads Tracker Module) ---
// This defines the expected structure for the leads sheet. The actual mapping is done dynamically in the leads script.
const LEADS_SHEET_HEADERS = [
    "Date Added", "Job Title", "Company", "Location", "Source Email Subject",
    "Link to Job Posting", "Status", "Source Email ID", "Processed Timestamp", "Notes"
];
