using Excel = Microsoft.Office.Interop.Excel;
Console.WriteLine("Hello, World!");
Console.WriteLine("Beginning...");

var clientId = "KZwbNOJRTQ2_bM_lOBIkiQ";
var clientSecret = "6I4Dt3oYuVCrAv10l2OWailNs1D45ZVr";
var accountId = "1LfeiJSCSdW0BCLXHeRrOA";

var userID = "born4cheese@gmail.com";
long webinarID = 0; //Insert Webinar ID here.

var connectionInfo = new ZoomNet.OAuthConnectionInfo(clientId, clientSecret, accountId,
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
}

Console.WriteLine("Goodbye, World!");
Console.Read();