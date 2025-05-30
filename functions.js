
function openTemplate(event) {
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["support@yourcompany.com"],
    subject: "New Support Request: [Issue Type] - [Short Description]",
    htmlBody: `
      <p>Hello Support Team,</p>
      <p>Please find the details of the support request below:</p>
      <ul>
        <li><b>Requestor Name:</b> [Auto-filled or manually entered]</li>
        <li><b>Department:</b> [Dropdown or manual entry]</li>
        <li><b>Issue Type:</b> [Hardware / Software / Access / Other]</li>
        <li><b>Priority:</b> [Low / Medium / High / Critical]</li>
        <li><b>Description of Issue:</b><br/>[User enters detailed description here]</li>
        <li><b>Attachments:</b> [Optional]</li>
      </ul>
      <p>Thank you,<br/>[User’s Name]</p>
    `
  });
  event.completed();
}
