# HTML

Reusable HTML email templates for Logic Apps, Power Automate, Azure Functions, or other notification workflows.

## Templates

| Template | Status | Purpose | Expected Inputs | Example Use |
| --- | --- | --- | --- | --- |
| `ApprovalEmail.html` | Current template | Approval-style email for a new user or account request submission. | Request metadata such as submitter, requestor, manager, job title, department, and other onboarding fields. | Use in an approval workflow that sends account request details to an administrator or manager. |
| `File-Upload-Email-Notification.html` | Current template | Notification for a successful file upload and index completion event. | File name, upload result, report metadata, and any workflow-specific status values. | Use after an automation pipeline finishes validating and indexing a newly uploaded file. |
| `Malicious-File-Upload-Email-Notification.html` | Current template | Alert template for a file upload detected as malicious. | File metadata, detection result, workflow identifiers, and remediation or escalation details. | Use when a malware scan or validation stage flags an uploaded file for security review. |
| `SFTP-Modification-Notification.html` | Current template | Notification for changes detected in an SFTP repository. | Subject, repository change details, modified file information, and workflow variables. | Use in a scheduled or event-driven monitor that reports new or changed files in SFTP. |
| `SharePoint-Upload-Notification.html` | Current template | Notification for a new file uploaded to SharePoint. | SharePoint file metadata, uploader details, site or library references, and report values. | Use in a SharePoint-triggered workflow that notifies a downstream processing team. |

## Typical Workflow

1. Open the template that matches the workflow event.
2. Replace placeholder text or Logic App expressions with your runtime fields.
3. Send the HTML output through your mail action or notification service.
4. Test the rendered email in Outlook and on mobile clients.

## Input Notes

- Several templates contain literal placeholder text such as `ENTER LOGIC APP FORM INFORMATION`.
- Some templates already include Logic App style expressions such as `@{variables('subject')}`.
- Styling is embedded directly in each file so the templates can be used without external CSS assets.

## Caution

- Email clients render HTML inconsistently, so test before production use.
- If you standardize branding, update colors, headings, and placeholder text across all templates together.