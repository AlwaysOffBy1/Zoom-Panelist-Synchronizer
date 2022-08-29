using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Collections.Specialized;


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
    webinarID = (long)Convert.ToDouble(ConfigurationManager.AppSettings.Get("WebinarID").Replace(" ", ""));
}
catch (FormatException ex)
{
    ExitWithError("The Webinar ID needs to be a number.");
}
catch (ArgumentNullException ex)
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
/*
Task<ZoomNet.Models.Panelist[]> panelistTask = zoomClient.Webinars.GetPanelistsAsync(webinarID);
try
{
    ZoomNet.Models.Panelist[]? panelists = await panelistTask;
    string str = "";
    foreach (ZoomNet.Models.Panelist panelist in panelists)
    {
        str += panelist.Email + ", ";
    }
    Console.WriteLine(str);
}
catch (ZoomNet.Utilities.ZoomException ex) { Console.WriteLine(ex.ToString()); Console.Read(); Environment.Exit(0); }
*/

//get member and guest lists from excel sheet.
string exePath = AppDomain.CurrentDomain.BaseDirectory;
if(File.Exists(exePath + @"Invitations_edit.xlsm"))
{
    Console.WriteLine("File found!");
    Excel.Application xlApp = new Excel.Application();
    Excel.Workbook workbook = xlApp.Workbooks.Open(exePath + @"Invitations_edit.xlsm");
    Excel.Worksheet instructionSheet = xlApp.Worksheets[0];
    foreach(Excel.Worksheet ws in workbook.Worksheets)
    {
        if(ws.Name == "Instructions")
        {
            instructionSheet = ws;
            break;
        }
    }
    //start cell end cell for range, add to array
    string test = instructionSheet.Cells[2, 2].value;
}
else
{
    
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