
<%

Dim objDBConn, objRs, objDBConn2
Dim strSQL
Dim strDate, strLname, strFName, strPnum, strEmail, strPass1, strPass2, strUserIP , strACOD, strCoNam              

strDate = now()
strZip=request.form("Zip")
strFName = Request.Form("fname")
strSta=request.form("Sta")
strCty=request.form("Cty")
strLname = Request.Form("lname")
strCun=Request.form("Cun")
strPnum = Request.Form("number")
strEmail = Request.Form("email")
strEmail2 = Request.Form("email2")
strPass1 = Request.Form("pass")
strPass2 = Request.Form("pass2")
strACOD=Request.Form("acod")
strAdr=Request.Form("Adr")
strAdr2=request.form("Adr2")
strUserIP = Request.ServerVariables("REMOTE_ADDR") '글쓴이의 아이피를 가져오는 센스
strCoNam=request.form("CoNam")




Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic

db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "FMemDB"


connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword





Set objDBConn = Server.CreateObject("ADODB.Connection")
Set objDBConn2 = Server.CreateObject("ADODB.Connection")
 





objDBConn.Open connectstr


strSQL="SELECT * FROM FMEMDB WHERE EML='"&strEmail&"'"


Set ors = objDBconn.execute (strSQL)

If strFName="" OR strLname="" Or strPnum="" Or strEmail="" Or strEMail2="" Or strPass1="" Or strPass2="" Then
ors.close

Set ors = nothing
Set objDBConn = nothing%>

<script language="javascript">
<!--
alert("ALL FIELDS MUST BE FILLED OUT COMPLETELY!");
location.href="customer.asp";
//-->
</script> 


<% Else If strPass1=strPass2 AND strEmail=strEmail2 AND ors.eof then
objdbconn2.open connectstr
strSQL = "Insert Into FMemDB ("
strSQL = strSQL & " FNAM"
strSQL = strSQL & ",LNAM"
strSQL = strSQL & ",PWD"
strSQL = strSQL & ",EML"
strSQL = strSQL & ",ACD"
strSQL = strSQL & ",PHN"
strSQL = strSQL & ",ADR"
strSQL = strSQL & ",CTY"
strSQL = strSQL & ",STA"
strSQL = strSQL & ",ZIP"
strSQL = strSQL & ",CUN"
strSQL = strSQL & ",REGIP"
strSQL = strSQL & ",ADR2"
strSQL = strSQL & ",CoNam"
strSQL = strSQL & ",Approved"
strSQL = strSQL & ") Values ("
strSQL = strSQL & "'" & strFName & "',"
strSQL = strSQL & "'" & strLName & "',"
strSQL = strSQL & "'" & strPass1 & "',"
strSQL = strSQL & "'" & strEmail & "',"
strSQL = strSQL & "'" & strACOD & "',"
strSQL = strSQL & "'" & strPnum &"',"
strSQL = strSQL & "'" & strAdr &"',"
strSQL = strSQL & "'" & strCty &"',"
strSQL = strSQL & "'" & strSta &"',"
strSQL = strSQL & "'" & strZip &"',"
strSQL = strSQL & "'" & StrCun &"',"
strSQL = strSQL & "'" & strUserIP &"',"
strSQL = strSQL & "'" & strAdr2 &"',"
strSQL = strSQL & "'" & strCoNam &"',"
strSQL = strSQL & "'NO')"
objDBConn2.Execute strSQL
objDBConn2.Close
ors.close

Set ors = nothing
Set objDBConn = nothing%>
<script language="javascript">
<!--
alert("Information saved!  Your information will be reviewed and your account will be created shortly!  Your email address is your username.");
location.href="customer.asp";
//-->
</script> 

<%
Else
ors.close

Set ors = nothing
Set objDBConn = nothing

%>
<script language="javascript">
<!--
alert("Passords and/or Email fields do not match OR That Email address has already been registered");
location.href="customer.asp";
//-->
</script> 
<%


End If
End if%>
