# Google-APIs-Web-Forms
.NET Web Forms C# Page Using Google Sheets API

 To use this you will need to use the NuGet package manager to download the Google.Apis packages.
 Also, you need to create a service account from the Google dev console. When you create the service account,
 download the .json file and put in the root directory of the site. Rename the file to client_secret.json
 
 IMPORTANT: in the .json file is the email address associated with the service account, you need to share
 the google sheet with this email address. Otherwise you will get an exception. * 
 
 Special thanks to the VB code provided here: https://stackoverflow.com/questions/22911691/google-apis-auth-oauth-for-webforms
 This VB example was the first Web Forms sample I was able to get working. I then converted the code to C#. I also added some 
 functionality: Binding the google sheets list data into a datagrid that allows updating and deleting. Also, a feature to add
 a data record. 
 
 Thanks to the codes samples here: https://www.twilio.com/blog/2017/03/google-spreadsheets-and-net-core.html
 I was able to add the create, update and delete record functions.
 
 Note. If you get an exception, try selecting a smaller range (eg. A0-B3). Make sure there are no blank cells in the 
 selected range.
