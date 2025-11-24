<%@ Language=VBScript %>
<% 


option explicit 
Response.Expires = -1
Server.ScriptTimeout = 600
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001
%>
<!-- #include file="freeaspupload.asp" -->
<%


  ' ****************************************************
  ' Change the value of the variable below to the pathname
  ' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  ' ****************************************************

  Dim uploadsDirVar, StrRecNo, strJnam, strEML, strPWD, strFileNam

  If strPwd="" Then 
  
  strPwd=request.querystring("PWD")

  End If
  
  If strEML="" then
  strEML=request.querystring("EML")
End If 


If strRecNo = "" Then 

  strRecNo=request.querystring("RecNo")


  
  End If
  
  If strJnam="" Then
  
  strJnam=request.querystring("Jnam")
 

 
 End If
 




  uploadsDirVar = "D:\Hosting\7626218\html\uploads" 
  

  ' Note: this file uploadTester2.asp is just an example to demonstrate
  ' the capabilities of the freeASPUpload.asp class. There are no plans
  ' to add any new features to uploadTester2.asp itself. Feel free to add
  ' your own code. If you are building a content management system, you
  ' may also want to consider this script: http://www.webfilebrowser.com/

function OutputForm()
%>
    <form name="frmSend" method="POST" enctype="multipart/form-data" accept-charset="utf-8" action="uploadTester2.asp?RecNo=<%=strRecNo%>&Jnam=<%=strJNam%>&Pwd=<%=strPwd%>&EML=<%=strEML%>" onSubmit="return onSubmitForm();">
	
	<B>File names:</B><br>
    File 1: <input name="attach1" type="file" size=35><br>
    File 2: <input name="attach2" type="file" size=35><br>
    File 3: <input name="attach3" type="file" size=35><br>
    File 4: <input name="attach4" type="file" size=35><br>
    <br> 

    <input style="margin-top:4" type=submit value="Submit">
    </form>
<%
end function

function TestEnvironment()
    Dim fso, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester2.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions." & Response.write(Server.MapPath("uploadTester2.asp"))
        exit function
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester2.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        TestEnvironment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey


Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic, strtyp

db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"

connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword


Dim objDBConn, objRs
Dim strSQL




Set objDBConn = Server.CreateObject("ADODB.Connection")
  Set objRs = Server.CreateObject("ADODB.RecordSet")
objDBConn.Open connectstr













    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) Then
    

        SaveFiles = "<B>Files uploaded:</B> "
        for each fileKey in Upload.UploadedFiles.keys
            SaveFiles = SaveFiles &strRecNo&Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
       strFilenam=strRecNo&Upload.UploadedFiles(fileKey).FileName
 strFilenam = Replace(strFilenam, "'", "''")
strSQL = "Insert Into Uploads ("
strSQL = strSQL & " Job"
strSQL = strSQL & ",filenam"
strSQL = strSQL & ") Values ("
strSQL = strSQL & "'" & strRecNo & "',"
strSQL = strSQL & "'"&strFilenam&"')"

		objDBConn.Execute strSQL
		
		

		
		Next
		
    else
        SaveFiles = "No file selected for upload or the file name specified in the upload form does not correspond to a valid file in the system."
    end if
	objDBConn.Close
objRS.Close
Set objRS=nothing
Set objDBConn = Nothing
end Function

On Error Resume Next 
%>

<HTML>
<HEAD>
<TITLE>Finco Metals Upload your file...</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
BODY {background-color: white;font-family:arial; font-size:12}
</style>
<script>
function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.attach1.value == "" && formDOMObj.attach2.value == "" && formDOMObj.attach3.value == "" && formDOMObj.attach4.value == "" )
        alert("Please press the Browse button and pick a file.")
    else
        return true;
    return false;
}
</script>

</HEAD>

<BODY>


<br><br>
<div style="border-bottom: #A91905 2px solid;font-size:16">Upload your pictures for job <%=strRecNo%>-<%=strJNam%>.</div>
<%
Dim diagnostics
if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = TestEnvironment()
    if diagnostics<>"" then
        response.write "<div style=""margin-left:20; margin-top:30; margin-right:30; margin-bottom:30;"">"
        response.write diagnostics
        response.write "<p>After you correct this problem, reload the page."
        response.write "</div>"
    else
        response.write "<div style=""margin-left:150"">"
        OutputForm()
        response.write "</div>"
    end if
else
    response.write "<div style=""margin-left:150"">"
    OutputForm()
    response.write SaveFiles()
    response.write "<br><br></div>"%>

Upload successful!  You can add more files or...
<a href="details.asp?reco=<%=strRecno%>&pass=99999">GO BACK</a>

	<%

end if

%>

<!-- Please support this free script by having a link to freeaspupload.net either in this page or somewhere else in your site. -->
<div style="border-bottom: #A91905 2px solid;font-size:10">Powered by <A HREF="http://www.fincometals.com/" style="color:black">www.fincometals.com</A></div>

<br><br>

<!--- START OF HTML TO REMOVE - contains the script ratings submission -->


</table>
<!-- end of html to remove ------------------------->

</BODY>
</HTML>

