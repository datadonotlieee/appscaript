# KOL Request Form - Architecture & Module Documentation

## Overview
This is a **Google Apps Script web application** that serves as a multi-step KOL (Key Opinion Leader) request form for Summit Media. The application uses a modular architecture where the backend (Google Apps Script) manages data and server-side logic, while the frontend (HTML/CSS/JavaScript) handles the user interface.

---

## Module Hierarchy & Call Flow

### 📊 System Architecture Diagram

```
ENTRY POINT
    │
    ├── Code.gs (Backend - Google Apps Script)
    │   ├── Configuration (CONFIG object)
    │   ├── doGet() ──────────────────┐
    │   ├── doPost()                  │
    │   ├── submitForm()              │ Called by Scripts.html via
    │   ├── insertData()              │ google.script.run
    │   ├── sendEmailNotification()   │
    │   └── sendTeamsNotification()   │
    │                                  │
    └─────────────────────────────────┘
                                      │
                Query/Response via google.script.run (async)
                                      │
                    ┌─────────────────┘
                    │
                    ▼
        Index.html (Main Template)
            │
            ├── Static HTML Structure
            ├── Header & Progress Bar
            ├── Form Container
            │   ├── Step 1 (via include)
            │   ├── Step 2 (via include)
            │   ├── Step 3 (via include)
            │   └── Step 4 (via include)
            └── Navigation Buttons
                │
                ├── Styles.html (CSS)
                │   └── Brand styling & responsive design
                │
                ├── Step1.html
                │   └── Requestor Details form fields
                │
                ├── Step2.html
                │   ├── Editorial Participation fields
                │   ├── Media Partnership fields
                │   ├── Paid Partnership fields
                │   └── Others fields
                │
                ├── Step3.html
                │   ├── Post Engagement section
                │   ├── Event Requirements section
                │   ├── Talent Requirements section
                │   └── Others section
                │
                ├── Step4.html
                │   └── Review & Confirmation layout
                │
                └── Scripts.html (Frontend JavaScript)
                    ├── Navigation Logic
                    │   ├── showStep()
                    │   ├── nextStep()
                    │   └── previousStep()
                    ├── Form Handling
                    │   ├── collectFormData()
                    │   ├── validateForm()
                    │   └── submitForm()
                    ├── Dynamic UI Logic
                    │   ├── setupTypeOfRequestListener()
                    │   ├── setupDeliverablesListener()
                    │   ├── setupBusinessUnitListener()
                    │   └── setupTimePickerSync()
                    └── Communication Layer
                        └── google.script.run.submitForm(formData)
                            ├── Calls Code.gs::submitForm()
                            └── Triggers data insertion & notifications
```

---

## File Structure & Responsibilities

### Backend Module: `Code.gs` (67 KB)
**Responsibility**: Server-side logic, data persistence, external notifications

#### Key Functions:
```
1. CONFIGURATION
   └── CONFIG object (sheets, email, teams webhook)

2. WEB APP ENTRY POINTS
   ├── doGet()
   │   └── Renders Index.html with all included modules
   └── doPost()
       └── Alternative entry point if needed

3. FORM SUBMISSION HANDLER
   ├── submitForm(formData)
   │   └── Returns confirmation with requestId

4. DATA INSERTION FUNCTIONS
   ├── insertMainRequest()
   ├── insertEditorialDetails()
   ├── insertMediaDetails()
   ├── insertPaidDetails()
   ├── insertDeliverablesInfo()
   ├── insertPostingRequirements()
   └── insertEventActivityDetails()

5. NOTIFICATION FUNCTIONS
   ├── sendEmailNotification(requestData)
   └── sendTeamsNotification(requestData)

6. UTILITY FUNCTIONS
   ├── generateRequestId()
   ├── formatDate()
   └── etc.
```

**Data Persistence**: Writes to 7 Google Sheets
- `Main_KOL_Requests`
- `Editorial_Participation_Details`
- `Media_Partnership_Details`
- `Paid_Partnership_Details`
- `KOL_Deliverables`
- `Posting_Requirements`
- `Event_Activity_Details`

---

### Frontend Modules

#### `Index.html` (Main Container)
**Responsibility**: Main template structure and module inclusion

Uses the `<?!= include('moduleName') ?>` directive to embed:
- Styles.html (CSS)
- Step1.html - Step4.html (Form content)
- Scripts.html (JavaScript)

**Structure**:
```html
<!DOCTYPE html>
<head>
  <!-- Meta tags and styles -->
  <?!= include('Styles'); ?>
</head>
<body>
  <!-- Toast notifications container -->
  <!-- Loading spinner -->
  <div class="container">
    <!-- Header with logo -->
    <!-- Progress bar (4 steps) -->
    <form>
      <div id="step1"><?!= include('Step1'); ?></div>
      <div id="step2"><?!= include('Step2'); ?></div>
      <div id="step3"><?!= include('Step3'); ?></div>
      <div id="step4"><?!= include('Step4'); ?></div>
      <!-- Navigation buttons -->
    </form>
    <!-- Success message template -->
  </div>
  <?!= include('Scripts'); ?>
</body>
</html>
```

---

#### `Styles.html` (CSS Module)
**Responsibility**: All visual styling and responsive design

**Scope**:
- CSS variables for brand colors (Summit Media palette)
- Base styling (typography, layout)
- Component styling (forms, buttons, progress bar)
- Interactive states (focus, hover, disabled)
- Responsive breakpoints
- Special components (time picker, toast notifications, loading spinner)

**Key CSS Classes**:
- `.container` - Main form wrapper
- `.form-step` - Each step container
- `.form-group` - Form field wrapper
- `.dynamic-fields` - Conditional content
- `.progress-bar` - Step indicator
- `.btn` - Button styles (primary, secondary, success)
- `.toast-container` - Notifications

---

#### `Step1.html` - Requestor Details
**Responsibility**: Collect initial requestor information

**Form Fields**:
- Requestor Name (text, required)
- Requestor Email (email, required)
- Business Unit (select, required) + conditional "Other" field
- Participating Brands (checkboxes, required) - 9 brands available
- Type of Request (select, required) + conditional "Other" field

**Data Model**:
```javascript
{
  requestorName: string,
  requestorEmail: string,
  businessUnit: string,
  businessUnitOther?: string,
  participatingBrands: array<string>,
  typeOfRequest: string,
  typeOfRequestOther?: string
}
```

---

#### `Step2.html` - Request Details
**Responsibility**: Request-specific details based on type selection

**Conditional Sections** (shown based on `typeOfRequest`):

1. **Editorial Participation** (`#editorialFields`)
   - KOL Ambassadors selection (6 options)
   - Detailed KOL Ambassadors (conditional)
   - Other Inclusions (conditional)
   - Number of KOLs (stepper control)
   - KOL Description (textarea)

2. **Media Partnership** (`#mediaFields`)
   - KOL Ambassadors selection (6 options)
   - Number of KOLs (stepper control)
   - KOL Description (textarea)

3. **Paid Partnership** (`#paidFields`)
   - Pitch/GO status (select)
   - KOL Ambassadors selection (6 options)
   - Detailed KOL Ambassadors (conditional)
   - Other Inclusions (conditional)
   - Number of KOLs (stepper control)
   - KOL Description (textarea)

4. **Others** (`#othersFields`)
   - Number of KOLs (stepper control)
   - KOL Description (textarea)

**Dynamic Behavior**:
- Only one section visible at a time (controlled by Scripts.html)
- Event listener: `setupTypeOfRequestListener()` triggers field visibility
- Number stepper buttons with +/- controls

---

#### `Step3.html` - KOL Deliverables & Requirements
**Responsibility**: Deliverables selection and detailed requirements

**Main Sections**:

1. **Cross-posting Toggle** (conditional)
   - Shows when applicable
   - Cross-posting Instructions textarea (conditional)

2. **Post Engagement** (`#postEngagementSection`)
   - Checkboxes for social posting options
   - **Social Posting Details** (conditional when selected):
     - Target Live Date (date picker)
     - Mandatories (textarea)
     - References URL (text field)

3. **Event Requirements** (`#eventRequirementsSection`)
   - Checkboxes for event options

4. **Talent Requirements** (`#talentRequirementsSection`)
   - Checkboxes for talent requirement options
   - **Event Details** (shown when event/talent selected):
     - Event/Activity Name (text)
     - Date (date picker)
     - Time (custom time picker)
       - Start time
       - End time
     - Address/Venue (textarea)
     - Notes (textarea)

5. **Others** (`#otherDeliverablesSection`)
   - "Others, please specify" checkbox
   - Conditional textarea for details

**Key Feature**: Custom time picker component
- Hour/Minute/AM-PM selection
- Scroll-based picker UI
- OK/CANCEL buttons

---

#### `Step4.html` - Review & Confirmation
**Responsibility**: Display summary for final approval

**Review Sections**:
1. **Basic Information**
   - Requestor Name
   - Requestor Email
   - Business Unit
   - Participating Brands
   - Type of Request

2. **Request Details** (dynamic)
   - Content varies by request type

3. **KOL Deliverables** (conditional)
   - Details of selected deliverables

4. **Social Posting Details** (conditional)
   - If posting requirements were selected

5. **Event Attendance/Talent Details** (conditional)
   - If event requirements were selected

**Data Flow**: All data populated from form via `populateReviewStep()` in Scripts.html

---

#### `Scripts.html` - Frontend Logic (104 KB)
**Responsibility**: All client-side interactivity and backend communication

### Major Function Groups:

#### 1. **Navigation Logic**
```javascript
function setupNavigation()
function showStep(stepNumber)
function previousStep()
function nextStep()
```
- Manages step visibility
- Validates current step before proceeding
- Updates progress bar indicators

#### 2. **Form Data Collection**
```javascript
function getFormData()
function collectFieldData(requestType)
function populateReviewStep()
```
- Gathers all form inputs
- Extracts selected checkboxes, radio buttons, text fields
- Structures data for backend

#### 3. **Form Validation**
```javascript
function validateStep(stepNumber)
function setupInputValidation()
```
- Checks required fields
- Validates email format
- Ensures at least one selection where needed
- Shows error messages for invalid inputs

#### 4. **Dynamic Field Management**
```javascript
function setupTypeOfRequestListener()
function setupDeliverablesListener()
function setupBusinessUnitListener()
```
- Shows/hides conditional fields
- Toggles form sections based on selections
- Populates dynamic checkboxes

#### 5. **Special Components**
```javascript
function setupTimePickerSync()
function setupNumberSteppers()
function setupScrollIntoView()
```
- Time picker initialization and value sync
- Number input stepper buttons
- Auto-scroll to form fields on focus

#### 6. **Form Submission**
```javascript
function submitForm(e)
function handleSubmissionResponse(response)
function showSuccessMessage(requestId)
```
- Collects all form data
- Shows loading spinner
- Calls `google.script.run.submitForm(formData)`
- Handles backend response
- Displays success message with Request ID

#### 7. **Backend Communication**
```javascript
google.script.run
  .withSuccessHandler(handleSubmissionResponse)
  .withFailureHandler(handleError)
  .submitForm(formData)
```
- Async call to Code.gs::submitForm()
- Success handler processes response
- Failure handler shows error notifications

#### 8. **Notifications**
```javascript
function showToast(message, type)
function showLoadingSpinner(show)
```
- Toast messages for feedback
- Loading spinner during submission
- Error alerts

---

## Call Flow Diagram: Form Submission Process

```
User fills entire form (Steps 1-4)
         │
         ▼
User clicks "Submit Request" (Step 4)
         │
         ▼
Scripts.html: submitForm() event handler
         │
         ├─ getFormData() → collects all inputs
         │
         ├─ validateStep(4) → validates review page
         │
         ├─ showLoadingSpinner(true)
         │
         └─ google.script.run.submitForm(formData)
                        │
                        │ (ASYNC - crosses boundary)
                        │
                        ▼
        Code.gs: submitForm(formData)
                        │
                        ├─ generateRequestId()
                        │
                        ├─ insertMainRequest() → writes to Main_KOL_Requests
                        │
                        ├─ insertEditorialDetails() [if applicable]
                        │
                        ├─ insertMediaDetails() [if applicable]
                        │
                        ├─ insertPaidDetails() [if applicable]
                        │
                        ├─ insertDeliverablesInfo() → KOL_Deliverables
                        │
                        ├─ insertPostingRequirements() [if applicable]
                        │
                        ├─ insertEventActivityDetails() [if applicable]
                        │
                        ├─ sendEmailNotification(requestData)
                        │   └─ Sends to CONFIG.EMAIL.COMPANY_EMAIL
                        │
                        ├─ sendTeamsNotification(requestData) [if enabled]
                        │   └─ Posts to Teams webhook
                        │
                        └─ return { success: true, requestId: "KOL-123456" }
                        │
                        └─ (ASYNC RESPONSE)
                        │
                        ▼
        Scripts.html: handleSubmissionResponse(response)
                        │
                        ├─ showLoadingSpinner(false)
                        │
                        ├─ displaySuccessMessage(response.requestId)
                        │
                        └─ showToast("Request submitted successfully!")
```

---

## Data Flow Between Modules

### Step 1 → Step 2
- **Trigger**: `setupTypeOfRequestListener()` detects change in `typeOfRequest` select
- **Action**: Shows/hides appropriate fields in Step 2 (`#editorialFields`, `#mediaFields`, `#paidFields`, `#othersFields`)
- **Data Passed**: Selected request type

### Step 3 - Dynamic Field Visibility
- **Trigger**: `setupDeliverablesListener()` detects checkbox changes
- **Action**: Shows/hides sections:
  - Post Engagement → Social Posting Details
  - Event/Talent Requirements → Event Details Section
- **Data Passed**: Selected deliverables

### Final Submission
- **Trigger**: Submit button click on Step 4
- **Process**:
  1. Collect all form data via `getFormData()`
  2. Send to backend via `google.script.run.submitForm(formData)`
  3. Backend processes and stores in 7 Google Sheets
  4. Backend sends email and Teams notification
  5. Frontend receives response with Request ID
  6. Display success message

---

## Module Dependencies

```
Code.gs (Independent - no dependencies)
    │
    ├─ relies on: Google Sheets API, Gmail API, HTTP for Teams

Index.html (Container - depends on all included modules)
    │
    ├─ includes: Styles.html
    ├─ includes: Step1.html
    ├─ includes: Step2.html
    ├─ includes: Step3.html
    ├─ includes: Step4.html
    └─ includes: Scripts.html

Styles.html (Independent CSS)
    └─ no dependencies

Step1.html (Depends on: Scripts.html)
    └─ Uses: setupBusinessUnitListener(), setupTypeOfRequestListener()

Step2.html (Depends on: Scripts.html)
    └─ Uses: setupTypeOfRequestListener()

Step3.html (Depends on: Scripts.html)
    └─ Uses: setupDeliverablesListener(), setupTimePickerSync()

Step4.html (Depends on: Scripts.html)
    └─ Uses: populateReviewStep()

Scripts.html (Depends on: Code.gs)
    └─ Calls: google.script.run.submitForm()
```

---

## Configuration & Settings

**In Code.gs**:
```javascript
const CONFIG = {
  SHEETS: {
    MAIN: 'Main_KOL_Requests',
    EDITORIAL: 'Editorial_Participation_Details',
    MEDIA: 'Media_Partnership_Details',
    PAID: 'Paid_Partnership_Details',
    DELIVERABLES: 'KOL_Deliverables',
    POSTING: 'Posting_Requirements',
    EVENT: 'Event_Activity_Details'
  },
  EMAIL: {
    ENABLED: true,
    COMPANY_EMAIL: 'boborol.marcelangelo.beloy@gmail.com'
  },
  TEAMS: {
    ENABLED: true,
    WEBHOOK_URL: '[Power Automate Webhook URL]'
  }
}
```

---

## Deployment & Access

- **Type**: Google Apps Script Web App
- **Access Level**: ANYONE_ANONYMOUS (publicly accessible)
- **Runtime**: V8
- **Timezone**: Asia/Singapore
- **OAuth Scopes Required**:
  - `https://www.googleapis.com/auth/spreadsheets` (data storage)
  - `https://www.googleapis.com/auth/script.send_mail` (email notifications)
  - `https://www.googleapis.com/auth/script.external_request` (Teams webhook)

---

## Summary

This application uses a **modular architecture** where:
- **Backend (Code.gs)** handles business logic, data persistence, and notifications
- **Frontend (Index.html + included modules)** provides the UI framework
- **Styling (Styles.html)** keeps presentation separate from logic
- **Content (Step1-4.html)** organizes form sections
- **Interaction (Scripts.html)** manages all client-side behavior and server communication

The modules communicate through:
1. **Server-to-Client**: `doGet()` renders complete HTML
2. **Client-to-Server**: `google.script.run.submitForm(formData)`
3. **Within Frontend**: Event listeners and DOM manipulation trigger conditional field visibility

This separation of concerns makes the codebase maintainable and scalable.
