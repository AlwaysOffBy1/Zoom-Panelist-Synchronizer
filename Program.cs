using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Collections.Specialized;
using System.Linq;

Console.WriteLine("Hello, World!");
Console.WriteLine("Beginning...");

//Retrieve connection info from App.config
string? clientID = ConfigurationManager.AppSettings.Get("ClientID");
string? clientSecret = ConfigurationManager.AppSettings.Get("ClientSecret"); 
string? accountId = ConfigurationManager.AppSettings.Get("AccountID"); 
string? userID = ConfigurationManager.AppSettings.Get("UserID");
long webinarID = 0;


//Check if connection info is null
bool emptyField = false;
string msg = "The application cannot proceed because the following fields in App.config are null\n\n";
if (clientID is null) { msg += "Client ID\n"; emptyField = true; }
if (clientSecret is null) { msg += "Client Secret\n"; emptyField = true; }
if (accountId is null) { msg += "Account ID\n"; emptyField = true; }
if (userID is null) { msg += "User ID (Email address for account)\n"; emptyField = true; }
if (emptyField) ExitWithError(msg); 
try
{
    webinarID = (long)Convert.ToDouble(ConfigurationManager.AppSettings.Get("WebinarID")?.Replace(" ", ""));
}
catch (FormatException)
{
    ExitWithError("The Webinar ID needs to be a number.");
}
catch (ArgumentNullException)
{
    ExitWithError("The Webinar ID cannot be null.");
}


//Connect to zoom using ZoomNet
var connectionInfo = new ZoomNet.OAuthConnectionInfo(clientID, clientSecret, accountId,
    (_, newAccessToken) =>
    {

    });

var zoomClient = new ZoomNet.ZoomClient(connectionInfo);

//Getting Panelists from Zoom
Task<ZoomNet.Models.Panelist[]> panelistTask = zoomClient.Webinars.GetPanelistsAsync(webinarID);
try
{
    ZoomNet.Models.Panelist[]? panelists = await panelistTask;
    string str = "";
    foreach (ZoomNet.Models.Panelist panelist in panelists)
    {
        str += panelist.Email + ", ";
    }
    Console.WriteLine($"From Zoom, fetching panelist {str}\n");
}
catch (ZoomNet.Utilities.ZoomException ex) { Console.WriteLine(ex.ToString()); Console.Read(); Environment.Exit(0); }


//get member and guest lists from excel sheet.
string exePath = AppDomain.CurrentDomain.BaseDirectory;
string projPath = Path.GetFullPath(Path.Combine(exePath, @"..\..\..\"));
if(File.Exists(projPath + @"Invitations_edit.xlsm"))
{
    Console.WriteLine("File found! Opening...");
    Excel.Application xlApp = new Excel.Application() { Visible=true, WindowState=Excel.XlWindowState.xlMinimized};
    
    Excel.Workbook workbook = xlApp.Workbooks.Open(projPath + @"Invitations_edit.xlsm");
    Excel.Worksheet instructionSheet = xlApp.Worksheets[1];
    Console.WriteLine("File opened! Importing data from Excel...");
    foreach(Excel.Worksheet ws in workbook.Worksheets)
    {
        if(ws.Name == "Instructions")
        {
            instructionSheet = ws;
            break;
        }
    }
    //start cell end cell for range, add to array
    List<(string Email, string FullName, string VirtualBackgroundId)> panelistsToDelete = new List<(string Email, string FullName, string VirtualBackgroundId)>();
    List<(string Email, string FullName, string VirtualBackgroundId)> panelistsToInsert = new List<(string Email, string FullName, string VirtualBackgroundId)>();
    for (int i = 2; i <= instructionSheet.UsedRange.Rows.Count; i++)
    {
        if (instructionSheet.Cells[i, 1].value is not null)
        {
            panelistsToDelete.Add((instructionSheet.Cells[i, 2].value, instructionSheet.Cells[i, 1].value, ""));
        }
        if(instructionSheet.Cells[i,3].value is not null)
        {
            panelistsToInsert.Add((instructionSheet.Cells[i, 4].value, instructionSheet.Cells[i, 3].value, ""));
        }
        Console.WriteLine($"Importing row {i}");
    }
    Console.WriteLine("Finished importing data from Excel! Beginning Sync...");

    //Excel import success. Now we need to sync!
    //Invitations.xlsm had done the work of making sure invitees are not on both the "delete" and "add" lists, so we can delete first then add!
    //TODO AddPanelistsAsync does not try to insert a panelist who is already present. This means Invitations.xlsm can send over duplicate invitees!
    await zoomClient.Webinars.AddPanelistsAsync(webinarID, panelistsToInsert);
    Console.WriteLine($"{panelistsToInsert} have been inserted");
    //no way to remove multiple panelists at once so im using a loop
    foreach((string Email, string FullName, string VirtualBackgroundId) p in panelistsToDelete)
    {
        await zoomClient.Webinars.RemovePanelistAsync(webinarID, p.Email);
        Console.WriteLine($"{p.Email} has been removed");
    }
    
    
}
else
{
    ExitWithError("You must add your data to import into the Invitations.xlsm worksheet before running this program");
}



Console.WriteLine("Goodbye, World!");
Console.Read();


void ExitWithError(string msg)
{
    Console.WriteLine(msg);
    Console.WriteLine("Press any key to exit");
    Console.Read();
    Environment.Exit(0);
}