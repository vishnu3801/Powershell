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
   // PowerShell command to import the Exchange Online module and connect
   const connectCommand = `Install-module Exchnageonlinemanagement; Import-Module ExchangeOnlineManagement; Connect-ExchangeOnline  -ShowProgress $true`;
   // PowerShell command to add mailbox permission
   const addPermissionCommand = `Add-MailboxPermission -Identity ${mailbox} -User ${user} -AccessRights ${permission}`;
   // Execute the PowerShell command to connect to Exchange Online
   exec(connectCommand, (connectError, connectStdout, connectStderr) => {
       if (connectError) {
           console.error(`Error: ${connectError.message}`);
           res.status(500).json({ message: 'An error occurred while connecting to Exchange Online.' });
           return;
       }
       if (connectStderr) {
           console.error(`stderr: ${connectStderr}`);
           res.status(500).json({ message: 'An error occurred while connecting to Exchange Online.' });
           return;
       }
       // Execute the PowerShell command to add mailbox permission
       exec(addPermissionCommand, (addError, addStdout, addStderr) => {
           if (addError) {
               console.error(`Error: ${addError.message}`);
               res.status(500).json({ message: 'An error occurred while adding mailbox permission.' });
               return;
           }
           if (addStderr) {
               console.error(`stderr: ${addStderr}`);
               res.status(500).json({ message: 'An error occurred while adding mailbox permission.' });
               return;
           }
           console.log(`stdout: ${addStdout}`);
           res.json({ message: 'Mailbox permission added successfully.' });
       });
   });
});
app.listen(port, () => {
   console.log(`Server is running on http://localhost:${port}`);
});
