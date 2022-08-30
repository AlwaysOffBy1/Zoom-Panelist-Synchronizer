# **Zoom Panelist Synchronizer**

#### The problem:
Zoom allows their users to uplaod a CSV of panelists for a webinar, but 
1. If you have a panelist on both your CSV and within Zoom, Zoom wont upload the CSV
2. If you delete the user on Zoom to upload the CSV their invitation link will change

#### The solution:
With **Zoom Panelist Synchronizer** you use the `Invitations.xlsm`  file to create a full list of people invited to the webinar. Then,
1. if you have a panelist on both the `Invitations.xlsm` file and within Zoom, **Zoom Panelist Sychronizer** will simply ignore them. *(planned to resend their original invitation email)*
2. if a user does not appear on the `Invitations.xlsm` spreadsheet, they are removed from the webinar

##How to edit:

1. Download and open the code in Visual Studio
2. In the Solution Explorer, double click  `Zoom Panelist Synchronizer.sln` 
3. Begin editing!

##How to run:

1. Download the code
2. Find the file `App.config` 
3. Put in your credentials
4. Navigate to the /bin/Release/net6.0/Zoom Panelist Sychronizer.exe and run!