const express = require('express');
const bodyParser = require('body-parser');
const { exec } = require('child_process');
const app = express();
const port = 3000;
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.get('/', (req, res) => {
   res.sendFile(__dirname + '/index.html');
});
app.post('/addMailboxPermission', (req, res) => {
   const { mailbox, user, permission } = req.body;
   // PowerShell command to add mailbox permission
   const command = `
       Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop;
       Import-Module ExchangeOnlineManagement;
       $UserCredential = Get-Credential;
       Connect-ExchangeOnline -Credential $UserCredential;
       Add-MailboxPermission -Identity "${mailbox}" -User "${user}" -AccessRights "${permission}";
   `;
   // Execute the PowerShell command
   exec(`pwsh -Command "${command}"`, (error, stdout, stderr) => {
       if (error) {
           console.error(`Error: ${error.message}`);
           res.status(500).json({ message: 'An error occurred while executing the PowerShell command.' });
           return;
       }
       if (stderr) {
           console.error(`stderr: ${stderr}`);
           res.status(500).json({ message: 'An error occurred while executing the PowerShell command.' });
           return;
       }
       console.log(`stdout: ${stdout}`);
       res.json({ message: 'Mailbox permission added successfully.' });
   });
});
app.listen(port, () => {
   console.log(`Server is running on http://localhost:${port}`);
});
