def create_reset_pw_content(knt, url, code):
    return f'''Dear {knt},

We received a request to reset your EIW Website account password. Click the link below to reset your password:

{url}

Additionally, use the following code to complete the reset process:

Reset Code: {code}

If you did not request this, please ignore this email.

Best regards,
EIW Team
'''


def create_reminder_content(department_name, tool_link, sharepoint_link):
    return f'''Dear Promotion Team of {department_name},

The tool has been registered, but no documents have been uploaded. Please upload the necessary documents to the corresponding SharePoint.

Tool Link: {tool_link}
SharePoint Link for Documents: {sharepoint_link}

We will delete the tool in two weeks  from the registration date if the documents are not uploaded to the provided SharePoint.

Thank you.

Best regards,
EIW Team
''', "Subject: 【Remind】【EIW Tool】Upload Documents to SharePoint"


def create_request_sharepoint_content(current_sharepoint, link_request_sharepoint, link_sharepoint_management):
    return f'''Dear Admin EIW,

This SharePoint has reached its storage limit: {current_sharepoint}. Please request a new SharePoint site for this department.

Link to request SharePoint: {link_request_sharepoint}

Once the new SharePoint is ready, kindly update the information on the SharePoint management page: {link_sharepoint_management}

Thank you.

Best regards,
EIW Team
''', "Subject: 【Remind】【EIW Tool】Request for New SharePoint Site"
