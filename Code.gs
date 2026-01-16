    /* ============================================
    KOL REQUEST FORM - GOOGLE APPS SCRIPT
    ============================================
    
    Structure:
    1. Configuration
    2. Web App Entry Points
    3. Main Form Handler
    4. Data Insertion Functions
    5. Email Notification Functions
    6. Utility Functions
    
    ============================================ */


    /* ============================================
    1. CONFIGURATION
    ============================================ */

    const CONFIG = {
    // Sheet Names
    SHEETS: {
        MAIN: 'Main_KOL_Requests',
        EDITORIAL: 'Editorial_Participation_Details',
        MEDIA: 'Media_Partnership_Details',
        PAID: 'Paid_Partnership_Details',
        DELIVERABLES: 'KOL_Deliverables',
        POSTING: 'Posting_Requirements',
        EVENT: 'Event_Activity_Details'
    },
    
    // Email Notification Settings
    EMAIL: {
        ENABLED: true,
        // Company email - used as both sender and recipient
        // Note: In Google Apps Script, the sender is always the account running the script
        // This email receives all form submission notifications
        COMPANY_EMAIL: 'boborol.marcelangelo.beloy@gmail.com'
    },
    
    // Microsoft Teams Notification Settings (via Power Automate Workflow)
    TEAMS: {
        ENABLED: true,
        // To get webhook URL:
        // 1. In Teams channel, click "..." > "Workflows"
        // 2. Search for "Send webhook alerts to test" or similar webhook template
        // 3. Click "Add workflow" and configure it for your channel
        // 4. Copy the webhook URL and paste it below
        WORKFLOW_URL: 'https://defaulte8e1fa0b7d9d418783a28a8010ef06.b4.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/42711929e9e3426c8e66b8b746d79287/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tTMb4is7W1k2keTopnmDxQb-ib21eExUqp8QwmPwZkY'
    },
    
    // ID Prefixes
    ID_PREFIX: {
        REQUEST: 'REQ',
        EDITORIAL: 'ED',
        MEDIA: 'MD',
        PAID: 'PD',
        DELIVERABLE: 'DL',
        POSTING: 'PR',
        EVENT: 'EV'
    }
    };


    /* ============================================
    2. WEB APP ENTRY POINTS
    ============================================ */

    /**
    * Serves the HTML form when the web app is accessed
    */
    function doGet() {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('KOL Request Form')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    /**
    * Includes HTML files for templating
    * @param {string} filename - Name of the HTML file to include
    */
    function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
    }


    /* ============================================
    3. MAIN FORM HANDLER
    ============================================ */

    /**
    * Main form submission handler
    * @param {Object} formData - Form data from the client
    * @returns {Object} Success/error response with request ID
    */
    function submitForm(formData) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const requestId = generateRequestId(ss, formData.typeOfRequest);
        
        // 1. Insert main request
        insertMainRequest(ss, requestId, formData);
        
        // 2. Insert type-specific details
        insertTypeSpecificDetails(ss, requestId, formData);
        
        // 3. Insert deliverables (if any)
        if (hasDeliverables(formData)) {
        insertKOLDeliverables(ss, requestId, formData);
        }
        
        // 4. Insert posting requirements (if applicable)
        if (hasPostingRequirements(formData)) {
        insertPostingRequirements(ss, requestId, formData);
        }
        
        // 5. Insert event details (if applicable)
        if (hasEventDetails(formData)) {
        insertEventActivityDetails(ss, requestId, formData);
        }
        
        // 6. Send email notification
        if (CONFIG.EMAIL.ENABLED) {
        sendEmailNotification(requestId, formData);
        }
        
        // 7. Send Teams notification
        if (CONFIG.TEAMS.ENABLED) {
        sendTeamsNotification(requestId, formData);
        }
        
        return { success: true, requestId: requestId };
        
    } catch (error) {
        Logger.log('Error in submitForm: ' + error.toString());
        return { success: false, error: error.toString() };
    }
    }

    /**
    * Routes to the correct type-specific insert function
    */
    function insertTypeSpecificDetails(ss, requestId, formData) {
    const typeHandlers = {
        'Editorial Participation': insertEditorialDetails,
        'Media Partnership': insertMediaPartnershipDetails,
        'Paid Partnership': insertPaidPartnershipDetails
    };
    
    const handler = typeHandlers[formData.typeOfRequest];
    if (handler) {
        handler(ss, requestId, formData);
    }
    }


    /* ============================================
    4. DATA INSERTION FUNCTIONS
    ============================================ */

    /**
    * Inserts main request data
    */
    function insertMainRequest(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.MAIN);
    
    sheet.appendRow([
        requestId,
        data.requestorName,
        data.requestorEmail,
        data.businessUnit,
        data.teamBrand,
        data.summitTitle || '',
        data.typeOfRequest,
        new Date()
    ]);
    }

    /**
    * Inserts Editorial Participation details
    */
    function insertEditorialDetails(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.EDITORIAL);
    const editorialId = generateId(CONFIG.ID_PREFIX.EDITORIAL);
    
    const editorialTypes = arrayToString(data.editorialTypes) || data.editorialType || '';
    const detailedInclusions = arrayToString(data.detailedInclusions) || '';
    
    sheet.appendRow([
        editorialId,
        requestId,
        editorialTypes,
        detailedInclusions,
        data.otherInclusion || '',
        data.numberOfKOLs || '',
        data.kolDescription || ''
    ]);
    }

    /**
    * Inserts Media Partnership details
    */
    function insertMediaPartnershipDetails(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.MEDIA);
    const mediaId = generateId(CONFIG.ID_PREFIX.MEDIA);
    
    const summitMediaBrands = arrayToString(data.summitMediaBrands) || data.summitMediaBrand || '';
    const kolAmbassadorTypes = arrayToString(data.kolAmbassadorTypes) || data.kolAmbassadorType || '';
    
    sheet.appendRow([
        mediaId,
        requestId,
        summitMediaBrands,
        kolAmbassadorTypes,
        data.numberOfKOLs || '',
        data.kolDescription || ''
    ]);
    }

    /**
    * Inserts Paid Partnership details
    */
    function insertPaidPartnershipDetails(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.PAID);
    const paidId = generateId(CONFIG.ID_PREFIX.PAID);
    
    const paidTypes = arrayToString(data.paidTypes) || '';
    const paidDetailedInclusions = arrayToString(data.paidDetailedInclusions) || '';
    
    sheet.appendRow([
        paidId,
        requestId,
        data.pitchGo || '',
        paidTypes,
        paidDetailedInclusions,
        data.paidOtherInclusion || '',
        data.numberOfKOLs || '',
        data.kolDescription || ''
    ]);
    }

    /**
    * Inserts KOL Deliverables
    */
    function insertKOLDeliverables(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.DELIVERABLES);
    const deliverableId = generateId(CONFIG.ID_PREFIX.DELIVERABLE);
    const d = data.deliverables || {};
    
    sheet.appendRow([
        deliverableId,
        requestId,
        // Social Posting - with quantities
        qtyToValue(d.igReel),
        qtyToValue(d.tiktokVideo),
        qtyToValue(d.igCarousel),
        qtyToValue(d.tiktokCarousel),
        qtyToValue(d.igStories),
        // Cross-posting
        data.crossPosting || 'No',
        // Event - boolean checkboxes
        boolToYes(d.eventAttendance),
        boolToYes(d.eventParticipation),
        boolToYes(d.eventSpeaker),
        boolToYes(d.eventHosting),
        boolToYes(d.eventPerformer),
        // Talent - boolean checkboxes
        boolToYes(d.videoTalent),
        boolToYes(d.voTalent),
        boolToYes(d.videoTalentModel),
        boolToYes(d.photoshootTalent),
        boolToYes(d.resourcePerson),
        // Other
        data.otherDeliverable || ''
    ]);
    }

    /**
    * Converts quantity to display value
    * Returns the number if > 0, empty string otherwise
    */
    function qtyToValue(value) {
    const qty = parseInt(value) || 0;
    return qty > 0 ? qty.toString() : '';
    }

    /**
    * Inserts Posting Requirements
    */
    function insertPostingRequirements(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.POSTING);
    const postingId = generateId(CONFIG.ID_PREFIX.POSTING);
    
    sheet.appendRow([
        postingId,
        requestId,
        data.targetLiveDate || '',
        data.mandatories || '',
        data.references || ''
    ]);
    }

    /**
    * Inserts Event/Activity Details
    */
    function insertEventActivityDetails(ss, requestId, data) {
    const sheet = getSheet(ss, CONFIG.SHEETS.EVENT);
    const eventId = generateId(CONFIG.ID_PREFIX.EVENT);
    
    sheet.appendRow([
        eventId,
        requestId,
        data.eventActivityName || '',
        data.eventDate || '',
        data.eventTime || '',
        data.eventVenueAddress || '',
        data.eventNotes || ''
    ]);
    }


    /* ============================================
    5. EMAIL NOTIFICATION FUNCTIONS
    ============================================ */

    /**
    * Sends email notification on form submission
    * - Sends full notification to company email (with spreadsheet link)
    * - Sends receipt copy to requestor email (without spreadsheet link)
    * 
    * @param {string} requestId - The request ID
    * @param {Object} formData - The form data object
    */
    function sendEmailNotification(requestId, formData) {
    try {
        Logger.log('sendEmailNotification called with requestId: ' + requestId);
        
        // Validate inputs
        if (!formData) {
        Logger.log('Error: formData is undefined. Use testEmailNotification() to test.');
        return;
        }
        if (!requestId) {
        Logger.log('Error: requestId is undefined.');
        return;
        }
        
        const companyEmail = CONFIG.EMAIL.COMPANY_EMAIL;
        const requestorEmail = formData.requestorEmail;
        const typeOfRequest = formData.typeOfRequest || 'KOL Request';
        const subject = 'New KOL Request: ' + requestId + ' - ' + typeOfRequest;
        const requestorSubject = 'KOL Request Submitted: ' + requestId + ' - ' + typeOfRequest;
        
        // 1. Send email to company (with spreadsheet link)
        Logger.log('Building company email template...');
        const companyHtmlBody = buildEmailTemplate(requestId, formData, true);
        
        Logger.log('Sending email to company: ' + companyEmail);
        MailApp.sendEmail({
        to: companyEmail,
        subject: subject,
        htmlBody: companyHtmlBody
        });
        Logger.log('Company email sent successfully to: ' + companyEmail);
        
        // 2. Send receipt email to requestor (without spreadsheet link)
        if (requestorEmail) {
        Logger.log('Building requestor receipt email template...');
        const requestorHtmlBody = buildEmailTemplate(requestId, formData, false);
        
        Logger.log('Sending receipt email to requestor: ' + requestorEmail);
        MailApp.sendEmail({
            to: requestorEmail,
            subject: requestorSubject,
            htmlBody: requestorHtmlBody
        });
        Logger.log('Receipt email sent successfully to: ' + requestorEmail);
        } else {
        Logger.log('No requestor email provided, skipping receipt email.');
        }
        
    } catch (error) {
        Logger.log('Email notification error: ' + error.toString());
        Logger.log('Error stack: ' + error.stack);
    }
    }

    /**
    * Builds HTML email template with all form details
    * Matches the format of the review page (Step 4)
    * 
    * @param {string} requestId - The request ID
    * @param {Object} formData - The form data object
    * @param {boolean} includeSpreadsheetLink - Whether to include the spreadsheet link (true for company, false for requestor)
    */
    function buildEmailTemplate(requestId, formData, includeSpreadsheetLink) {
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    
    // Build section title based on request type
    const detailsSectionTitle = getDetailsSectionTitle(formData.typeOfRequest);
    
    // Build type-specific details
    const typeSpecificHtml = buildTypeSpecificDetailsHtml(formData);
    
    // Build deliverables section
    const deliverablesSection = buildDeliverablesSection(formData);
    
    // Build posting requirements section
    const postingSection = buildPostingSection(formData);
    
    // Build event details section
    const eventSection = buildEventSection(formData);
    
    // Customize header and footer based on recipient type
    const isCompanyEmail = includeSpreadsheetLink;
    const headerTitle = isCompanyEmail ? 'New KOL Request Submitted' : 'Your KOL Request Has Been Submitted';
    const headerSubtitle = isCompanyEmail 
        ? 'A new request has been submitted and is awaiting review.' 
        : 'Thank you for submitting your KOL request. Below is a copy for your records.';
    const footerText = isCompanyEmail
        ? 'This is an automated notification from the KOL Request Form.'
        : 'This is a confirmation receipt of your KOL request submission. Our team will review your request and get back to you soon.';
    
    // Build button section (only for company email)
    const buttonSection = includeSpreadsheetLink 
        ? `<a href="${spreadsheetUrl}" class="button" style="display: inline-block; background-color: #4C88FF; color: #ffffff !important; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: 500; margin-top: 8px;">View in Spreadsheet</a>`
        : `<div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 16px; margin-top: 8px;">
            <p style="margin: 0; color: #166534; font-size: 14px;">
            <strong>✓ Request Received</strong><br>
            <span style="color: #15803d;">Your request has been successfully submitted. Please save this email and your Request ID for future reference.</span>
            </p>
        </div>`;
    
    return `
        <!DOCTYPE html>
        <html>
        <head>
        <style>
            body { font-family: Arial, sans-serif; line-height: 1.6; color: #333333; margin: 0; padding: 0; background-color: #f5f5f5; }
            .container { max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5; }
            .header { background: linear-gradient(135deg, #4C88FF 0%, #2E5CB8 100%); color: #ffffff; padding: 28px 24px; border-radius: 8px 8px 0 0; }
            .header h1 { margin: 0; font-size: 22px; font-weight: 600; color: #ffffff; }
            .header p { margin: 8px 0 0; opacity: 0.9; font-size: 14px; color: #ffffff; }
            .request-id { display: inline-block; background: rgba(0, 0, 0, 0.2); padding: 6px 14px; border-radius: 4px; font-size: 13px; margin-top: 12px; font-weight: 600; color: #ffffff; border: 1px solid rgba(255, 255, 255, 0.3); }
            .content { background: #ffffff; border: 1px solid #e5e5e5; border-top: none; padding: 24px; border-radius: 0 0 8px 8px; }
            
            /* Review Section Styling - matches form review page */
            .review-section { 
            background: #fafafa; 
            border-radius: 6px; 
            padding: 16px; 
            margin-bottom: 16px; 
            border: 1px solid #e5e5e5; 
            border-left: 3px solid #4C88FF; 
            }
            .review-section h3 { 
            font-size: 11px; 
            font-weight: 600; 
            color: #2E5CB8; 
            text-transform: uppercase; 
            letter-spacing: 0.5px; 
            margin: 0 0 12px; 
            padding: 0; 
            }
            .review-item { 
            padding: 10px 0; 
            border-bottom: 1px solid #e5e5e5; 
            }
            .review-item:last-child { 
            border-bottom: none; 
            padding-bottom: 0; 
            }
            .review-label { 
            display: block;
            color: #737373; 
            font-size: 11px; 
            text-transform: uppercase; 
            letter-spacing: 0.3px; 
            margin-bottom: 4px;
            }
            .review-value { 
            display: block;
            color: #171717; 
            font-weight: 500; 
            font-size: 14px;
            word-break: break-word;
            }
            .review-value:empty::before,
            .empty-value::before { 
            content: '—'; 
            color: #d4d4d4; 
            font-weight: 400; 
            }
            
            /* Badge styling */
            .badge { display: inline-block; background: #e6f2ff; color: #4C88FF; padding: 4px 12px; border-radius: 4px; font-size: 13px; font-weight: 500; }
            .badge-outline { display: inline-block; background: #f5f5f5; color: #525252; padding: 3px 8px; border-radius: 4px; font-size: 12px; margin: 2px 4px 2px 0; }
            
            /* Button styling - inline styles also added to element for email client compatibility */
            .button { display: inline-block; background-color: #4C88FF !important; color: #ffffff !important; padding: 12px 24px; text-decoration: none !important; border-radius: 6px; font-weight: 500; margin-top: 8px; }
            
            /* Footer */
            .footer { text-align: center; margin-top: 24px; font-size: 12px; color: #a3a3a3; background-color: #f5f5f5; }
            .footer p { margin: 4px 0; color: #a3a3a3; }
        </style>
        </head>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333; margin: 0; padding: 0; background-color: #f5f5f5;">
        <div class="container" style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f5f5f5;">
            <div class="header" style="background: linear-gradient(135deg, #4C88FF 0%, #2E5CB8 100%); background-color: #4C88FF; color: #ffffff; padding: 28px 24px; border-radius: 8px 8px 0 0;">
            <h1 style="margin: 0; font-size: 22px; font-weight: 600; color: #ffffff;">${headerTitle}</h1>
            <p style="margin: 8px 0 0; font-size: 14px; color: #ffffff; opacity: 0.9;">${headerSubtitle}</p>
            <p style="margin: 12px 0 0; font-size: 13px; color: #ffffff; opacity: 0.85;">Request ID: <strong>${requestId}</strong></p>
            </div>
            <div class="content" style="background: #ffffff; border: 1px solid #e5e5e5; border-top: none; padding: 24px; border-radius: 0 0 8px 8px;">
            
            <!-- Basic Information Section -->
            <div class="review-section" style="background: #fafafa; border-radius: 6px; padding: 16px; margin-bottom: 16px; border: 1px solid #e5e5e5; border-left: 3px solid #4C88FF;">
                <h3 style="font-size: 11px; font-weight: 600; color: #2E5CB8; text-transform: uppercase; letter-spacing: 0.5px; margin: 0 0 12px; padding: 0;">Basic Information</h3>
                <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
                <span class="review-label">Requestor Name</span>
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Requestor Name</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.requestorName || '—'}</span>
                </div>
                <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Requestor Email</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.requestorEmail || '—'}</span>
                </div>
                <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Business Unit</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.businessUnit || '—'}</span>
                </div>
                <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Participating Brands</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.teamBrand || '—'}</span>
                </div>
                <div class="review-item" style="padding: 10px 0; border-bottom: none;">
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Type of Request</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;"><span class="badge" style="display: inline-block; background: #e6f2ff; color: #4C88FF; padding: 4px 12px; border-radius: 4px; font-size: 13px; font-weight: 500;">${formData.typeOfRequest || '—'}</span></span>
                </div>
            </div>
            
            <!-- Request Details Section (type-specific) -->
            <div class="review-section" style="background: #fafafa; border-radius: 6px; padding: 16px; margin-bottom: 16px; border: 1px solid #e5e5e5; border-left: 3px solid #4C88FF;">
                <h3 style="font-size: 11px; font-weight: 600; color: #2E5CB8; text-transform: uppercase; letter-spacing: 0.5px; margin: 0 0 12px; padding: 0;">${detailsSectionTitle}</h3>
                ${typeSpecificHtml}
                <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Number of KOLs</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.numberOfKOLs || '—'}</span>
                </div>
                <div class="review-item" style="padding: 10px 0; border-bottom: none;">
                <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">KOL Description</span>
                <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.kolDescription || '—'}</span>
                </div>
            </div>
            
            ${deliverablesSection}
            
            ${postingSection}
            
            ${eventSection}
            
            ${buttonSection}
            </div>
            <div class="footer" style="text-align: center; margin-top: 24px; font-size: 12px; color: #a3a3a3; background-color: #f5f5f5;">
            <p style="margin: 4px 0; color: #a3a3a3;">${footerText}</p>
            <p style="margin: 4px 0; color: #a3a3a3;">Summit Media • Clout+KOL Request System</p>
            </div>
        </div>
        </body>
        </html>
    `;
    }

    /**
    * Gets the section title based on request type
    */
    function getDetailsSectionTitle(typeOfRequest) {
    switch (typeOfRequest) {
        case 'Editorial Participation':
        return 'Editorial Participation Details';
        case 'Media Partnership':
        return 'Media Partnership Details';
        case 'Paid Partnership':
        return 'Paid Partnership Details';
        default:
        return 'Request Details';
    }
    }

    /**
    * Builds HTML for type-specific details matching review page format
    */
    function buildTypeSpecificDetailsHtml(formData) {
    let html = '';
    
    if (formData.typeOfRequest === 'Editorial Participation') {
        const types = arrayToString(formData.editorialTypes) || formData.editorialType || '';
        const detailedInclusions = arrayToString(formData.detailedInclusions) || '';
        
        if (types) {
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">KOL Ambassadors</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${types}</span>
            </div>
        `;
        }
        if (detailedInclusions) {
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Detailed KOL Ambassadors</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${detailedInclusions}</span>
            </div>
        `;
        }
        if (formData.otherInclusion) {
        const formatted = formData.otherInclusion.split('\n').filter(line => line.trim()).map(line => '• ' + line.trim()).join('<br>');
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Other Inclusions</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formatted}</span>
            </div>
        `;
        }
    }
    
    if (formData.typeOfRequest === 'Media Partnership') {
        const kolTypes = arrayToString(formData.kolAmbassadorTypes) || formData.kolAmbassadorType || '';
        
        if (kolTypes) {
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">KOL Ambassadors</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${kolTypes}</span>
            </div>
        `;
        }
    }
    
    if (formData.typeOfRequest === 'Paid Partnership') {
        if (formData.pitchGo) {
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Pitch/GO</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.pitchGo}</span>
            </div>
        `;
        }
        const paidTypes = arrayToString(formData.paidTypes) || formData.paidType || '';
        const paidDetailedInclusions = arrayToString(formData.paidDetailedInclusions) || '';
        
        if (paidTypes) {
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">KOL Ambassadors</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${paidTypes}</span>
            </div>
        `;
        }
        if (paidDetailedInclusions) {
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Detailed KOL Ambassadors</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${paidDetailedInclusions}</span>
            </div>
        `;
        }
        if (formData.paidOtherInclusion) {
        const formatted = formData.paidOtherInclusion.split('\n').filter(line => line.trim()).map(line => '• ' + line.trim()).join('<br>');
        html += `
            <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Other Inclusions</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formatted}</span>
            </div>
        `;
        }
    }
    
    return html;
    }

    /**
    * Builds the deliverables section HTML matching review page format
    */
    function buildDeliverablesSection(formData) {
    const d = formData.deliverables || {};
    const items = [];
    
    // Social Posting - with quantities
    if (d.igReel) items.push(`IG Reel (${d.igReel})`);
    if (d.tiktokVideo) items.push(`TikTok Video (${d.tiktokVideo})`);
    if (d.igCarousel) items.push(`IG Carousel (${d.igCarousel})`);
    if (d.tiktokCarousel) items.push(`TikTok Carousel (${d.tiktokCarousel})`);
    if (d.igStories) items.push(`IG Stories (${d.igStories})`);
    
    // Event - boolean checkboxes
    if (d.eventAttendance) items.push('Event Attendance');
    if (d.eventParticipation) items.push('Event Participation');
    if (d.eventSpeaker) items.push('Event Speaker');
    if (d.eventHosting) items.push('Event Hosting');
    if (d.eventPerformer) items.push('Event Performer');
    
    // Talent - boolean checkboxes
    if (d.videoTalent) items.push('Video Talent');
    if (d.voTalent) items.push('VO Talent');
    if (d.videoTalentModel) items.push('Video Talent/Model');
    if (d.photoshootTalent) items.push('Photoshoot Talent');
    if (d.resourcePerson) items.push('Resource Person');
    
    // Check if cross-posting is enabled
    const crossPostingEnabled = formData.crossPosting === 'Yes';
    
    // Check if there are any deliverables or other deliverable text
    if (items.length === 0 && !formData.otherDeliverable) return '';
    
    let html = `
<div class="review-section" style="background: #fafafa; border-radius: 6px; padding: 16px; margin-bottom: 16px; border: 1px solid #e5e5e5; border-left: 3px solid #4C88FF;">
      <h3 style="font-size: 11px; font-weight: 600; color: #2E5CB8; text-transform: uppercase; letter-spacing: 0.5px; margin: 0 0 12px; padding: 0;">KOL Deliverables</h3>`;
    
    if (items.length > 0) {
        const crossPostingBadge = crossPostingEnabled ? ' <span style="display: inline-block; padding: 2px 8px; background-color: #4C88FF; color: white; font-size: 10px; font-weight: 600; border-radius: 12px; margin-left: 8px; text-transform: uppercase;">+ Cross-post</span>' : '';
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Selected Deliverables</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${items.join(', ')}${crossPostingBadge}</span>
        </div>`;
    }
    
    // Add cross-posting instructions if applicable
    if (crossPostingEnabled && formData.crossPostingInstructions) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: ${formData.otherDeliverable ? '1px solid #e5e5e5' : 'none'};">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Cross-Posting Instructions</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.crossPostingInstructions}</span>
        </div>`;
    }
    
    if (formData.otherDeliverable) {
        const formatted = formData.otherDeliverable.split('\n').filter(line => line.trim()).map(line => '• ' + line.trim()).join('<br>');
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: none;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Other Deliverables</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formatted}</span>
        </div>`;
    }
    
    html += `
        </div>`;
    
    return html;
    }

    /**
    * Builds the posting requirements section HTML matching review page format
    */
    function buildPostingSection(formData) {
    const hasPostingReqs = formData.targetLiveDate || formData.mandatories || formData.references;
    
    if (!hasPostingReqs) return '';
    
    let html = `
<div class="review-section" style="background: #fafafa; border-radius: 6px; padding: 16px; margin-bottom: 16px; border: 1px solid #e5e5e5; border-left: 3px solid #4C88FF;">
      <h3 style="font-size: 11px; font-weight: 600; color: #2E5CB8; text-transform: uppercase; letter-spacing: 0.5px; margin: 0 0 12px; padding: 0;">Social Posting Details</h3>`;
    
    if (formData.targetLiveDate) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Target Live Date</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.targetLiveDate}</span>
        </div>`;
    }
    
    if (formData.mandatories) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Mandatories</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.mandatories}</span>
        </div>`;
    }
    
    if (formData.references) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: none;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">References</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.references}</span>
        </div>`;
    }
    
    html += `
        </div>`;
    
    return html;
    }

    /**
    * Builds the event details section HTML matching review page format
    */
    function buildEventSection(formData) {
    if (!formData.eventActivityName) return '';
    
    let html = `
        <div class="review-section" style="background: #fafafa; border-radius: 6px; padding: 16px; margin-bottom: 16px; border: 1px solid #e5e5e5; border-left: 3px solid #4C88FF;">
        <h3 style="font-size: 11px; font-weight: 600; color: #2E5CB8; text-transform: uppercase; letter-spacing: 0.5px; margin: 0 0 12px; padding: 0;">Event Attendance / Talent Requirements Details</h3>
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Event/Activity Name</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.eventActivityName}</span>
        </div>`;
    
    if (formData.eventDate) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Event Date</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.eventDate}</span>
        </div>`;
    }
    
    if (formData.eventTime) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Event Time</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.eventTime}</span>
        </div>`;
    }
    
    if (formData.eventVenueAddress) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: 1px solid #e5e5e5;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Event Venue/Address</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.eventVenueAddress}</span>
        </div>`;
    }
    
    if (formData.eventNotes) {
        html += `
        <div class="review-item" style="padding: 10px 0; border-bottom: none;">
            <span class="review-label" style="display: block; color: #737373; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px;">Event Notes</span>
            <span class="review-value" style="display: block; color: #171717; font-weight: 500; font-size: 14px;">${formData.eventNotes}</span>
        </div>`;
    }
    
    html += `
        </div>`;
    
    return html;
    }

    /**
    * Test function - Run this to test email notification
    * Go to Apps Script Editor > Run > testEmailNotification
    */
    function testEmailNotification() {
    const testData = {
        requestorName: 'Test User',
        requestorEmail: 'test@summitmedia.com.ph',
        businessUnit: 'Digital',
        teamBrand: 'PEP, Preview',
        summitTitle: 'Test Summit Title',
        typeOfRequest: 'Paid Partnership',
        pitchGo: 'GO',
        paidTypes: ['Influencer Marketing', 'Content Creator'],
        paidDetailedInclusions: ['Beauty & Skincare', 'Lifestyle'],
        summitMediaBrands: ['PEP', 'Preview'],
        clientName: 'Test Client',
        brandProduct: 'Test Product',
        numberOfKOLs: '3',
        kolDescription: 'Looking for lifestyle influencers with 50k+ followers, authentic engagement, and experience in beauty/skincare content.',
        deliverables: {
        igReel: true,
        tiktokVideo: true,
        eventAttendance: true,
        photoshootTalent: true
        },
        otherDeliverable: 'Custom video for brand launch\nBehind-the-scenes content',
        targetLiveDate: '2024-02-15',
        mandatories: 'Must tag @brand and use hashtag #TestCampaign',
        references: 'https://example.com/reference',
        eventActivityName: 'Product Launch Event',
        eventDate: '2024-02-20',
        eventTime: '2:00 PM - 5:00 PM',
        eventVenueAddress: 'Summit Media HQ, BGC, Taguig',
        eventNotes: 'Dress code: Smart casual. Parking available.'
    };
    
    const testRequestId = 'TEST' + new Date().getTime();
    
    Logger.log('Sending test email to: ' + CONFIG.EMAIL.COMPANY_EMAIL);
    sendEmailNotification(testRequestId, testData);
    Logger.log('Test completed! Check the email inbox.');
    }

    /**
    * RUN THIS FUNCTION FIRST TO AUTHORIZE EMAIL!
    * This function forces the authorization dialog to appear.
    * After running once and authorizing, you can use testEmailNotification.
    */
    function authorizeEmailAccess() {
    // This will trigger authorization dialog
    MailApp.getRemainingDailyQuota();
    Logger.log('Authorization successful! Remaining daily quota: ' + MailApp.getRemainingDailyQuota());
    Logger.log('You can now run testEmailNotification');
    }

    /**
    * RUN THIS FUNCTION TO AUTHORIZE EXTERNAL REQUESTS (Teams notifications)
    * This function forces the authorization dialog for UrlFetchApp.
    */
    function authorizeExternalRequests() {
    // This triggers the authorization dialog for external URL requests
    const response = UrlFetchApp.fetch('https://httpbin.org/get', { muteHttpExceptions: true });
    Logger.log('Authorization successful! Response code: ' + response.getResponseCode());
    Logger.log('You can now run testTeamsNotification');
    }


    /* ============================================
    MICROSOFT TEAMS NOTIFICATION FUNCTIONS
    (Using Power Automate Workflows)
    ============================================ */

    /**
    * Sends Teams notification via Power Automate Workflow
    * @param {string} requestId - The request ID
    * @param {Object} formData - The form data object
    */
    function sendTeamsNotification(requestId, formData) {
    try {
        Logger.log('sendTeamsNotification called with requestId: ' + requestId);
        
        if (!CONFIG.TEAMS.WORKFLOW_URL || CONFIG.TEAMS.WORKFLOW_URL === 'YOUR_POWER_AUTOMATE_WORKFLOW_URL_HERE') {
        Logger.log('Teams Workflow URL not configured. Skipping Teams notification.');
        return;
        }
        
        if (!formData || !requestId) {
        Logger.log('Error: formData or requestId is undefined.');
        return;
        }
        
        const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
        const payload = buildTeamsWorkflowPayload(requestId, formData, spreadsheetUrl);
        
        const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(CONFIG.TEAMS.WORKFLOW_URL, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode >= 200 && responseCode < 300) {
        Logger.log('Teams notification sent successfully');
        } else {
        Logger.log('Teams notification failed. Response code: ' + responseCode);
        Logger.log('Response: ' + response.getContentText());
        }
    } catch (error) {
        Logger.log('Teams notification error: ' + error.toString());
    }
    }

    /**
    * Builds payload for Power Automate Workflow
    * Uses Adaptive Card format matching the email template structure
    */
    function buildTeamsWorkflowPayload(requestId, formData, spreadsheetUrl) {
    const typeOfRequest = formData.typeOfRequest || 'KOL Request';
    const detailsSectionTitle = getDetailsSectionTitle(typeOfRequest);
    
    // Build Basic Information facts
    const basicInfoFacts = [
        { title: 'Requestor Name', value: formData.requestorName || '—' },
        { title: 'Requestor Email', value: formData.requestorEmail || '—' },
        { title: 'Business Unit', value: formData.businessUnit || '—' },
        { title: 'Participating Brands', value: formData.teamBrand || '—' },
        { title: 'Summit Title', value: formData.summitTitle || '—' },
        { title: 'Type of Request', value: typeOfRequest }
    ];
    
    // Build type-specific details facts
    const detailsFacts = buildTeamsDetailsFacts(formData);
    detailsFacts.push({ title: 'Number of KOLs', value: formData.numberOfKOLs || '—' });
    detailsFacts.push({ title: 'KOL Description', value: formData.kolDescription || '—' });
    
    // Build deliverables facts
    const deliverablesFacts = buildTeamsDeliverablesFacts(formData);
    
    // Build posting requirements facts
    const postingFacts = buildTeamsPostingFacts(formData);
    
    // Build event details facts
    const eventFacts = buildTeamsEventFacts(formData);
    
    // Build the card body
    const cardBody = [
        // Header
        {
        "type": "Container",
        "style": "emphasis",
        "bleed": true,
        "backgroundColor": "#4C88FF",
        "items": [
            {
            "type": "TextBlock",
            "text": "📋 New KOL Request Submitted",
            "weight": "Bolder",
            "size": "Large",
            "color": "Light",
            "wrap": true
            },
            {
            "type": "TextBlock",
            "text": "A new request has been submitted and is awaiting review.",
            "spacing": "Small",
            "isSubtle": true,
            "wrap": true,
            "color": "Light"
            },
            {
            "type": "TextBlock",
            "text": "Request ID: " + requestId,
            "spacing": "Small",
            "weight": "Bolder",
            "color": "Warning",
            "wrap": true
            }
        ]
        },
        // Basic Information Section
        {
        "type": "Container",
        "spacing": "Medium",
        "items": [
            {
            "type": "TextBlock",
            "text": "BASIC INFORMATION",
            "weight": "Bolder",
            "size": "Small",
            "color": "Accent",
            "spacing": "Medium"
            },
            {
            "type": "FactSet",
            "facts": basicInfoFacts
            }
        ]
        },
        // Request Details Section
        {
        "type": "Container",
        "spacing": "Medium",
        "items": [
            {
            "type": "TextBlock",
            "text": detailsSectionTitle.toUpperCase(),
            "weight": "Bolder",
            "size": "Small",
            "color": "#4C88FF",
            "spacing": "Medium"
            },
            {
            "type": "FactSet",
            "facts": detailsFacts
            }
        ]
        }
    ];
    
    // Add Deliverables Section if exists
    if (deliverablesFacts.length > 0) {
        cardBody.push({
        "type": "Container",
        "spacing": "Medium",
        "items": [
            {
            "type": "TextBlock",
            "text": "KOL DELIVERABLES",
            "weight": "Bolder",
            "size": "Small",
            "color": "#4C88FF",
            "spacing": "Medium"
            },
            {
            "type": "FactSet",
            "facts": deliverablesFacts
            }
        ]
        });
    }
    
    // Add Posting Requirements Section if exists
    if (postingFacts.length > 0) {
        cardBody.push({
        "type": "Container",
        "spacing": "Medium",
        "items": [
            {
            "type": "TextBlock",
            "text": "SOCIAL POSTING DETAILS",
            "weight": "Bolder",
            "size": "Small",
            "color": "#4C88FF",
            "spacing": "Medium"
            },
            {
            "type": "FactSet",
            "facts": postingFacts
            }
        ]
        });
    }
    
    // Add Event Details Section if exists
    if (eventFacts.length > 0) {
        cardBody.push({
        "type": "Container",
        "spacing": "Medium",
        "items": [
            {
            "type": "TextBlock",
            "text": "EVENT ATTENDANCE / TALENT REQUIREMENTS DETAILS",
            "weight": "Bolder",
            "size": "Small",
            "color": "#4C88FF",
            "spacing": "Medium"
            },
            {
            "type": "FactSet",
            "facts": eventFacts
            }
        ]
        });
    }
    
    // Power Automate Workflow expects an Adaptive Card attachment
    return {
        type: "message",
        attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        contentUrl: null,
        content: {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.4",
            "body": cardBody,
            "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "View in Spreadsheet",
                "url": spreadsheetUrl,
                "style": "positive"
            }
            ]
        }
        }]
    };
    }

    /**
    * Builds type-specific details facts for Teams card
    */
    function buildTeamsDetailsFacts(formData) {
    const facts = [];
    
    if (formData.typeOfRequest === 'Editorial Participation') {
        const types = arrayToString(formData.editorialTypes) || formData.editorialType || '';
        const detailedInclusions = arrayToString(formData.detailedInclusions) || '';
        
        if (types) facts.push({ title: 'KOL Ambassadors', value: types });
        if (detailedInclusions) facts.push({ title: 'Detailed KOL Ambassadors', value: detailedInclusions });
        if (formData.otherInclusion) facts.push({ title: 'Other Inclusions', value: formData.otherInclusion });
    }
    
    if (formData.typeOfRequest === 'Media Partnership') {
        const brands = arrayToString(formData.summitMediaBrands) || formData.summitMediaBrand || '';
        const kolTypes = arrayToString(formData.kolAmbassadorTypes) || formData.kolAmbassadorType || '';
        
        if (brands) facts.push({ title: 'Summit Media Brand', value: brands });
        if (kolTypes) facts.push({ title: 'KOL Ambassadors', value: kolTypes });
    }
    
    if (formData.typeOfRequest === 'Paid Partnership') {
        if (formData.pitchGo) facts.push({ title: 'Pitch/GO', value: formData.pitchGo });
        
        const paidTypes = arrayToString(formData.paidTypes) || formData.paidType || '';
        const paidDetailedInclusions = arrayToString(formData.paidDetailedInclusions) || '';
        
        if (paidTypes) facts.push({ title: 'KOL Ambassadors', value: paidTypes });
        if (paidDetailedInclusions) facts.push({ title: 'Detailed KOL Ambassadors', value: paidDetailedInclusions });
        if (formData.paidOtherInclusion) facts.push({ title: 'Other Inclusions', value: formData.paidOtherInclusion });
    }
    
    return facts;
    }

    /**
    * Builds deliverables facts for Teams card
    */
    function buildTeamsDeliverablesFacts(formData) {
    const facts = [];
    const deliverables = getDeliverablesArray(formData);
    const crossPostingEnabled = formData.crossPosting === 'Yes';
    
    if (deliverables.length > 0) {
        // Filter out "Other:" items for the main list
        const mainDeliverables = deliverables.filter(d => !d.startsWith('Other:'));
        if (mainDeliverables.length > 0) {
        let value = mainDeliverables.join(', ');
        if (crossPostingEnabled) {
            value += ' [+ Cross-post]';
        }
        facts.push({ title: 'Selected Deliverables', value: value });
        }
    }
    
    // Add cross-posting instructions if applicable
    if (crossPostingEnabled && formData.crossPostingInstructions) {
        facts.push({ title: 'Cross-Posting Instructions', value: formData.crossPostingInstructions });
    }
    
    if (formData.otherDeliverable) {
        facts.push({ title: 'Other Deliverables', value: formData.otherDeliverable });
    }
    
    return facts;
    }

    /**
    * Builds posting requirements facts for Teams card
    */
    function buildTeamsPostingFacts(formData) {
    const facts = [];
    
    if (formData.targetLiveDate) facts.push({ title: 'Target Live Date', value: formData.targetLiveDate });
    if (formData.mandatories) facts.push({ title: 'Mandatories', value: formData.mandatories });
    if (formData.references) facts.push({ title: 'References', value: formData.references });
    
    return facts;
    }

    /**
    * Builds event details facts for Teams card
    */
    function buildTeamsEventFacts(formData) {
    const facts = [];
    
    if (formData.eventActivityName) {
        facts.push({ title: 'Event/Activity Name', value: formData.eventActivityName });
        if (formData.eventDate) facts.push({ title: 'Event Date', value: formData.eventDate });
        if (formData.eventTime) facts.push({ title: 'Event Time', value: formData.eventTime });
        if (formData.eventVenueAddress) facts.push({ title: 'Event Venue/Address', value: formData.eventVenueAddress });
        if (formData.eventNotes) facts.push({ title: 'Event Notes', value: formData.eventNotes });
    }
    
    return facts;
    }

    /**
    * Gets array of selected deliverables with quantities
    */
    function getDeliverablesArray(formData) {
    const d = formData.deliverables || {};
    const items = [];
    
    // Social deliverables - include quantities
    if (d.igReel) items.push('IG Reel (' + d.igReel + ')');
    if (d.tiktokVideo) items.push('TikTok Video (' + d.tiktokVideo + ')');
    if (d.igCarousel) items.push('IG Carousel (' + d.igCarousel + ')');
    if (d.tiktokCarousel) items.push('TikTok Carousel (' + d.tiktokCarousel + ')');
    if (d.igStories) items.push('IG Stories (' + d.igStories + ')');
    
    // Event deliverables - boolean
    if (d.eventAttendance) items.push('Event Attendance');
    if (d.eventParticipation) items.push('Event Participation');
    if (d.eventSpeaker) items.push('Event Speaker');
    if (d.eventHosting) items.push('Event Hosting');
    if (d.eventPerformer) items.push('Event Performer');
    
    // Talent deliverables - boolean
    if (d.videoTalent) items.push('Video Talent');
    if (d.voTalent) items.push('VO Talent');
    if (d.videoTalentModel) items.push('Video Talent/Model');
    if (d.photoshootTalent) items.push('Photoshoot Talent');
    if (d.resourcePerson) items.push('Resource Person');
    
    // Other deliverable text
    if (formData.otherDeliverable) items.push('Other: ' + formData.otherDeliverable);
    
    return items;
    }

    /**
    * Test function for Teams notification
    * Run this to test Power Automate Workflow
    */
    function testTeamsNotification() {
    const testData = {
        requestorName: 'Test User',
        requestorEmail: 'test@summitmedia.com.ph',
        businessUnit: 'Digital',
        teamBrand: 'PEP',
        typeOfRequest: 'Paid Partnership',
        pitchGo: 'GO',
        summitMediaBrands: ['PEP', 'Preview'],
        clientName: 'Test Client',
        brandProduct: 'Test Product',
        numberOfKOLs: '3',
        kolDescription: 'Looking for lifestyle influencers with 50k+ followers',
        deliverables: {
        igReel: true,
        tiktokVideo: true,
        eventAttendance: true
        },
        targetLiveDate: '2024-02-15',
        mandatories: 'Must tag @brand'
    };
    
    const testRequestId = 'TEST' + new Date().getTime();
    
    Logger.log('Sending test Teams notification via Power Automate...');
    sendTeamsNotification(testRequestId, testData);
    Logger.log('Test completed! Check your Teams channel.');
    }


    /* ============================================
    6. UTILITY FUNCTIONS
    ============================================ */

    /**
    * Gets a sheet by name, throws error if not found
    */
    function getSheet(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error('Sheet "' + sheetName + '" not found');
    }
    return sheet;
    }

    /**
    * Generates a unique ID with prefix
    */
    function generateId(prefix) {
    return prefix + new Date().getTime();
    }

    /**
    * Generates the main request ID in format XYY-###
    * X = First letter of request type (P=Paid, E=Editorial, M=Media, O=Others)
    * YY = Last two digits of current year
    * ### = Sequential submission number for this year (padded to 3 digits)
    * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
    * @param {string} typeOfRequest - The type of request
    */
    function generateRequestId(ss, typeOfRequest) {
    // Get the first letter of the request type
    let typePrefix = 'O'; // Default to Others
    if (typeOfRequest) {
        const type = typeOfRequest.toLowerCase();
        if (type.includes('paid')) {
        typePrefix = 'P';
        } else if (type.includes('editorial')) {
        typePrefix = 'E';
        } else if (type.includes('media')) {
        typePrefix = 'M';
        } else {
        typePrefix = 'O';
        }
    }
    
    // Get current year (last 2 digits)
    const now = new Date();
    const year = now.getFullYear().toString().slice(-2);
    
    // Get sequential number by counting existing rows in main sheet
    const mainSheet = getSheet(ss, CONFIG.SHEETS.MAIN);
    const lastRow = mainSheet.getLastRow();
    const sequenceNum = lastRow; // Row 1 is header, so lastRow equals the count of submissions
    
    // Format as 3 digits with leading zeros
    const paddedSequence = String(sequenceNum).padStart(3, '0');
    
    return `${typePrefix}${year}-${paddedSequence}`;
    }

    /**
    * Converts array to comma-separated string
    */
    function arrayToString(arr) {
    if (Array.isArray(arr) && arr.length > 0) {
        return arr.join(', ');
    }
    return '';
    }

    /**
    * Converts boolean/truthy value to 'Yes' or empty string
    */
    function boolToYes(value) {
    return value ? 'Yes' : '';
    }

    /**
    * Checks if form has deliverables
    */
    function hasDeliverables(formData) {
    return formData.deliverables && Object.keys(formData.deliverables).length > 0;
    }

    /**
    * Checks if form has posting requirements
    */
    function hasPostingRequirements(formData) {
    return formData.targetLiveDate || formData.mandatories || formData.references;
    }

    /**
    * Checks if form has event details
    */
    function hasEventDetails(formData) {
    return formData.eventActivityName;
    }


    /* ============================================
    7. SHEET SETUP FUNCTIONS
    ============================================ */

    /**
    * RUN THIS FUNCTION ONCE TO CREATE ALL REQUIRED SHEETS
    * This will create all necessary sheets with proper headers
    * Run from: Apps Script Editor > Run > setupAllSheets
    */
    function setupAllSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log('Setting up KOL Request Form sheets...');
    
    // Define all sheets with their headers
    const sheetsConfig = {
        [CONFIG.SHEETS.MAIN]: [
        'Request ID',
        'Requestor Name',
        'Requestor Email',
        'Business Unit',
        'Participating Brands',
        'Summit Title',
        'Type of Request',
        'Submission Date'
        ],
        
        [CONFIG.SHEETS.EDITORIAL]: [
        'Editorial ID',
        'Request ID',
        'KOL Ambassadors',
        'Detailed KOL Ambassadors',
        'Other Inclusions',
        'Number of KOLs',
        'KOL Description'
        ],
        
        [CONFIG.SHEETS.MEDIA]: [
        'Media ID',
        'Request ID',
        'Summit Media Brands',
        'KOL Ambassadors',
        'Number of KOLs',
        'KOL Description'
        ],
        
        [CONFIG.SHEETS.PAID]: [
        'Paid ID',
        'Request ID',
        'Pitch/GO',
        'KOL Ambassadors',
        'Detailed KOL Ambassadors',
        'Other Inclusions',
        'Number of KOLs',
        'KOL Description'
        ],
        
        [CONFIG.SHEETS.DELIVERABLES]: [
        'Deliverable ID',
        'Request ID',
        // Social Posting (with quantities)
        'IG Reel (Qty)',
        'TikTok Video (Qty)',
        'IG Carousel (Qty)',
        'TikTok Carousel (Qty)',
        'IG Stories (Qty)',
        // Cross-posting
        'Cross-posting',
        // Event
        'Event Attendance',
        'Event Participation',
        'Event Speaker',
        'Event Hosting',
        'Event Performer',
        // Talent
        'Video Talent',
        'VO Talent',
        'Video Talent/Model',
        'Photoshoot Talent',
        'Resource Person',
        // Other
        'Other Deliverables'
        ],
        
        [CONFIG.SHEETS.POSTING]: [
        'Posting ID',
        'Request ID',
        'Target Live Date',
        'Mandatories',
        'References'
        ],
        
        [CONFIG.SHEETS.EVENT]: [
        'Event ID',
        'Request ID',
        'Event/Activity Name',
        'Event Date',
        'Event Time',
        'Event Venue/Address',
        'Event Notes'
        ]
    };
    
    // Create/update each sheet
    Object.entries(sheetsConfig).forEach(([sheetName, headers]) => {
        let sheet = ss.getSheetByName(sheetName);
        
        if (!sheet) {
        // Create new sheet
        sheet = ss.insertSheet(sheetName);
        Logger.log('Created sheet: ' + sheetName);
        } else {
        Logger.log('Sheet exists: ' + sheetName);
        }
        
        // Set headers in row 1
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setValues([headers]);
        
        // Format header row
        headerRange.setBackground('#4C88FF');
        headerRange.setFontColor('#FFFFFF');
        headerRange.setFontWeight('bold');
        headerRange.setHorizontalAlignment('center');
        
        // Freeze header row
        sheet.setFrozenRows(1);
        
        // Auto-resize columns
        for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
        }
        
        // Set minimum column width
        for (let i = 1; i <= headers.length; i++) {
        if (sheet.getColumnWidth(i) < 100) {
            sheet.setColumnWidth(i, 100);
        }
        }
    });
    
    // Reorder sheets (Main first)
    const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
    if (mainSheet) {
        ss.setActiveSheet(mainSheet);
        ss.moveActiveSheet(1);
    }
    
    Logger.log('Sheet setup complete!');
    Logger.log('Sheets created/updated:');
    Object.keys(sheetsConfig).forEach(name => Logger.log('  - ' + name));
    }

    /**
    * Clears all data from sheets (keeps headers)
    * Use with caution! This deletes all submission data.
    * Run from: Apps Script Editor > Run > clearAllData
    */
    function clearAllData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetNames = Object.values(CONFIG.SHEETS);
    
    sheetNames.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && sheet.getLastRow() > 1) {
        // Delete all rows except header
        sheet.deleteRows(2, sheet.getLastRow() - 1);
        Logger.log('Cleared data from: ' + sheetName);
        }
    });
    
    Logger.log('All data cleared (headers preserved).');
    }

    /**
    * Adds sample test data to sheets for testing
    * Run from: Apps Script Editor > Run > addSampleData
    */
    function addSampleData() {
    const testData = {
        requestorName: 'Maria Santos',
        requestorEmail: 'maria.santos@summitmedia.com.ph',
        businessUnit: 'Digital Marketing',
        teamBrand: 'PEP, Preview, Cosmo',
        participatingBrands: ['PEP', 'Preview', 'Cosmo'],
        summitTitle: 'Beauty Editor',
        typeOfRequest: 'Paid Partnership',
        pitchGo: 'GO',
        paidTypes: ['Influencer', 'Content Creator'],
        paidDetailedInclusions: ['Beauty & Skincare', 'Lifestyle'],
        paidOtherInclusion: 'Custom brand video\nBTS content for socials',
        numberOfKOLs: '5',
        kolDescription: 'Looking for beauty and lifestyle influencers with 50k-200k followers, high engagement rate (3%+), and experience in skincare content.',
        deliverables: {
        igReel: 2,
        tiktokVideo: 3,
        igCarousel: 1,
        igStories: 5,
        eventAttendance: true,
        photoshootTalent: true
        },
        crossPosting: 'Yes',
        otherDeliverable: 'Brand ambassador announcement post\nProduct unboxing video',
        targetLiveDate: '2025-01-15',
        mandatories: 'Must tag @brand, use #BeautyLaunch2025, include product shot in first 3 seconds',
        references: 'https://instagram.com/p/example1\nhttps://tiktok.com/@example/video/123',
        eventActivityName: 'Beauty Product Launch Event',
        eventDate: '2025-01-20',
        eventTime: '2:00 PM - 6:00 PM',
        eventVenueAddress: 'Summit Media HQ\n6th Floor, Robinsons Cybergate\nPioneer St., Mandaluyong City',
        eventNotes: 'Dress code: Smart casual, all-white preferred\nParking available at basement\nMedia kit will be provided'
    };
    
    Logger.log('Adding sample data...');
    const result = submitForm(testData);
    
    if (result.success) {
        Logger.log('Sample data added successfully!');
        Logger.log('Request ID: ' + result.requestId);
    } else {
        Logger.log('Error adding sample data: ' + result.error);
    }
    }