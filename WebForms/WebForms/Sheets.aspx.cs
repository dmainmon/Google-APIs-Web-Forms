using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
/*
 *  Author: Damon Borgnino
 *  Date: Nov 12, 2018
 * 
 * To use this you will need to use the NuGet package manager to download the Google.Apis packages.
 * Also, you need to create a service account from the Google dev console. When you create the service account,
 * download the .json file and put in the root directory of the site. Rename the file to client_secret.json
 * 
 * IMPORTANT: in the .json file is the email address associated with the service account, you need to share
 * the google sheet with this email address. Otherwise you will get an exception. * 
 * 
 * Special thanks to the VB code provided here: https://stackoverflow.com/questions/22911691/google-apis-auth-oauth-for-webforms
 * This VB example was the first Web Forms sample I was able to get working. I then converted the code to C#. I also added some 
 * functionality: Binding the google sheets list data into a datagrid that allows updating and deleting. Also, a feature to add
 * a data record. 
 * 
 * Thanks to the codes samples here: https://www.twilio.com/blog/2017/03/google-spreadsheets-and-net-core.html
 * I was able to add the create, update and delete record functions.
 * 
 * Note. If you get an exception, try selecting a smaller range (eg. A0-B3). Make sure there are no blank cells in the 
 * selected range.
 * 
 */
using System.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using System.IO;


public partial class Sheets : System.Web.UI.Page
{

    static string spreadsheetId = "your-spreadsheet-id";
    static string mySheet = "Inventory By Location";
    static string ApplicationName = "Inventory By Location";
    static string[] Scopes = { SheetsService.Scope.Drive };
    static Int32 mySheetID = 18; // only required to delete rows
    static bool deleteRows = true; // false will clear the row leaving the cells blank, true removes the row
                                   /*
                                        To find the sheet id, open the spreadsheet URL in a browser,  
                                         https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit#gid={sheetId}
                                        sheedId is the number at the end of the url

                                   */
    Int32 myPageIndex = 0;
    String mySort = "Row ASC";



    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            //  ReadSheet(); // simple reading of data to a string of formatted HTML to show on page
            MyBindSheetData(mySort); // reads data and binds to a datagrid, sorts in ASCending or DESCending order
        }
    }

    private SheetsService GetServiceAccount()
    {
        string credPath = Server.MapPath("/") + "client_secret.json";

        ServiceAccountCredential credential;

        using (var stream =
            new FileStream(credPath, FileMode.Open, FileAccess.Read))
        {

            credential = (ServiceAccountCredential)GoogleCredential.FromStream(stream).CreateScoped(Scopes).UnderlyingCredential;
        }

        SheetsService service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });


        return service;

    }

    /*
    private void ReadSheet()
    {
        int rowCnt = 3;

        SheetsService service = GetServiceAccount();
        // string range = "Inventory By Location!A" + rowCnt  + ":C50";
        string range = mySheet + "!A3:C"; // I start on A3 because the sheet has blank fields in A1-A2 (throws exception)
                                          //  string range = mySheet + "!A5:C100"; // alternate selection
        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        ValueRange response = request.Execute();

        IList<IList<object>> values = response.Values;

        System.Text.StringBuilder HTMLtable = new System.Text.StringBuilder();
        HTMLtable.AppendLine("<table style=\"width:500px; padding:5px;\">");
        HTMLtable.AppendLine("<tr><th>Row</th><th>Location</th><th>SKU</th><th>QTY</th></tr>");
        // read data until cells are null after last record which throws exception. If any cells in the sheet are blank, an exception is 
        // thrown and reading stops. Make sure the sheet has no blank fields in the data being read. A blank field will stop the reading. 
        // If the read does NOT return all your records, the next record in the list probably has a blank field.

        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                try
                {
                    HTMLtable.AppendLine("<tr><td>" + rowCnt++ + "</td><td>" + row[0] + "</td><td>" + row[1] + "</td><td>" + row[2] + "</td></tr>");
                }
                catch (Exception)
                {
                    break; //stop reading
                }
            }
        }
        else
        {
            HTMLtable.AppendLine("<tr><th>No data found.</th></tr>");
        }

        HTMLtable.AppendLine("</table>");
        MyMessage.Text = HTMLtable.ToString();

        service.Dispose();
    }
    */

    // Updates a google sheet cell (SheetField), with a new value, eg. UpdateEntry("A0","my value")
    private void UpdateEntry(string SheetField, string newVal)
    {
        SheetsService service = GetServiceAccount();
        var range = mySheet + "!" + SheetField;
        var valueRange = new ValueRange();

        var oblist = new List<object>() { newVal };
        valueRange.Values = new List<IList<object>> { oblist };

        var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        var appendReponse = updateRequest.Execute();

        service.Dispose();

    }

    // creates a new row on a google spreadsheet
    private void CreateEntry()
    {
        SheetsService service = GetServiceAccount();

        var range = mySheet + "!A:F";
        var valueRange = new ValueRange();

        var oblist = new List<object>() { LocationBox.Text, SKUbox.Text, QBox.Text, "inserted", "via", "C#" };
        valueRange.Values = new List<IList<object>> { oblist };

        var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        var appendReponse = appendRequest.Execute();

        service.Dispose();
    }

    // removes or clears a row depending on the deleteRows switch setting
    private void DeleteEntry(string row)
    {

        SheetsService service = GetServiceAccount();

        if (deleteRows)
        {
            /* this code deletes the record (row) completely */

            Request RequestBody = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = mySheetID,
                        Dimension = "ROWS",
                        StartIndex = Convert.ToInt32(row) - 1,
                        EndIndex = Convert.ToInt32(row)
                    }
                }
            };

            List<Request> RequestContainer = new List<Request>();
            RequestContainer.Add(RequestBody);

            BatchUpdateSpreadsheetRequest DeleteRequest = new BatchUpdateSpreadsheetRequest();
            DeleteRequest.Requests = RequestContainer;

            // service.Spreadsheets.BatchUpdate(DeleteRequest, spreadsheetId);
            SpreadsheetsResource.BatchUpdateRequest Deletion = new SpreadsheetsResource.BatchUpdateRequest(service, DeleteRequest, spreadsheetId);
            Deletion.Execute();

            /* end of delete code */
        }
        else
        {
            
            /* this code clears the entry */
            var range = mySheet + "!A" + row + ":F" + row;
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
            var deleteReponse = deleteRequest.Execute();
        }

        service.Dispose();
    }

    public void MyBindSheetData(string sortStr)
    {

        int rowCnt = 3;

        DataTable dt = new DataTable();

        // Define the columns of the table.
        dt.Columns.Add(new DataColumn("Row", typeof(Int32)));
        dt.Columns.Add(new DataColumn("Location", typeof(String)));
        dt.Columns.Add(new DataColumn("SKU", typeof(String)));
        dt.Columns.Add(new DataColumn("QTY", typeof(Double)));


        SheetsService service = GetServiceAccount();
        string range = mySheet + "!A" + rowCnt + ":C";
        // I start on A3 because the sheet has blank fields in A1-A2 (throws exception)
        //  string range = mySheet + "!A5:C100"; // alternate selection
        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        ValueRange response = request.Execute();

        IList<IList<object>> values = response.Values;

        // read data until cells are null after last record which throws exception. If any cells in the sheet are blank, an exception is 
        // thrown and reading stops. Make sure the sheet has no blank fields in the data being read. A blank field will stop the reading. 
        // If the read does NOT return all your records, the next record in the list probably has a blank field.

        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                try // try is NOT needed if you limit the range to cells with NO blank fields, A5:C100 instead of A5:C
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = rowCnt.ToString();
                    dr[1] = row[0].ToString();
                    dr[2] = row[1].ToString();
                    dr[3] = row[2].ToString();
                    dt.Rows.Add(dr);

                    rowCnt++;
                }
                catch (Exception)
                {
                    break; // exception thrown for blank cells or EOF
                }
            }
        }
        else
        {
            DataRow dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = "No data found.";
            dr[2] = "No data found.";
            dr[3] = 0.0;
            dt.Rows.Add(dr);
        }

        service.Dispose();
        
        DataView dv = new DataView(dt); // use DataView to sort records then bind to the view instead of the table
        dv.Sort = sortStr;
        MySheetGrid.DataSource = dv;
      //  MySheetGrid.DataSource = dt; // comment this line out if using the DataView for sorting
        MySheetGrid.DataBind();

        dt.Dispose();

    }

    public void MyDataGrid_ItemDataBound(object sender, DataGridItemEventArgs e)
    {

        if ((e.Item.ItemType == ListItemType.Header) || (e.Item.ItemType == ListItemType.Footer))
            return;
        else
        {

            TableCell cellItemDel = (TableCell)e.Item.Cells[5]; // 0 being the first cell of the table                      

            Button LBtnDel = (Button)cellItemDel.Controls[0]; //0 being the first control in the TableCell
            LBtnDel.Attributes.Add("onclick", "javascript: return confirm('Are you sure you want to delete this entry?')");

        }

    }

    public void MyDataGrid_Update(Object sender, DataGridCommandEventArgs e)
    {
        TableCell cellItemRow = (TableCell)e.Item.Cells[0];
        TableCell cellItemQTY = (TableCell)e.Item.Cells[3];
        TextBox QTYbox = (TextBox)cellItemQTY.Controls[1];

        UpdateEntry("C" + cellItemRow.Text, QTYbox.Text);

        MyMessage.Text = "<div style=\"color:red\">Updated Row " + cellItemRow.Text + " Value: " + QTYbox.Text + "</div>";

        MyBindSheetData(mySort);
    }

    public void MyDataGrid_Delete(Object sender, DataGridCommandEventArgs e)
    {
        TableCell cellItemRow = (TableCell)e.Item.Cells[0];
        DeleteEntry(cellItemRow.Text);
        MyMessage.Text = "<div style=\"color:red\">Deleted Row: " + cellItemRow.Text + "</div>";

        if (((DataGrid)sender).Items.Count <= 1)
            MySheetGrid.CurrentPageIndex = MySheetGrid.CurrentPageIndex - 1;

        MyBindSheetData(mySort);
     
    }

    public void MyData_Paged(Object sender, DataGridPageChangedEventArgs e)
    {
        MySheetGrid.CurrentPageIndex = e.NewPageIndex;
        myPageIndex = e.NewPageIndex;
        MyBindSheetData(mySort);
    }

    protected void MyDataGrid_Sort(Object sender, DataGridSortCommandEventArgs e)
    {
        MyBindSheetData(e.SortExpression);
        mySort = e.SortExpression;
       
    }

    protected void AddBtn_Click(object sender, EventArgs e)
    {
        CreateEntry();
        // ReadSheet();
        MyBindSheetData("Row DESC");
        //myPageIndex = 0;
    }
}
