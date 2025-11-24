
<%

Dim objDBConnm, objRs
Dim strSQL
Dim strDate, strCnam, strJnam, strODate, strFSize, strMType, strSinfo, strComments, strDDate, strCam, strCnamnew        

strDate = now()
strCnam = Request.Form("Cnam")
strJnam = Request.Form("Jnam")
strOdate = Request.Form("ODate")
strDDate= Request.Form("DDate")
strFsize = Request.Form("FSize")
strMtype = Request.Form("MType")
strSinfo = Request.Form("Sinfo")
strComments = Request.Form("Comments")
strCnamNew=request.form("Cnamnew")

If strCnam="NEW" Then
strCnam=strCnamNew

Else

strCnam=request.form("cnam")

End If


If strJnam="" Then%>
<script language="javascript">
<!--
alert("Job Name can not be blank!");
location.href="jobupdater.asp?pass=99999";
//-->
</script> 


<%else

strJNam = Replace(strJNam, "'", "''")
strFSize = Replace(strFSize, "'", "''")
strMtype = Replace(strMtype, "'", "''")
strSinfo = Replace(strSinfo, "'", "''")
strComments = Replace(strComments, "'", "''")


strSQL = "Insert Into MemDB ("
strSQL = strSQL & " CNam"
strSQL = strSQL & ",JNam"
strSQL = strSQL & ",ODate"
strSQL = strSQL & ",DDate"

strSQL = strSQL & ",FSize"
strSQL = strSQL & ",MType"
strSQL = strSQL & ",SInfo"
strSQL = strSQL & ",Instructions"
strSQL = strSQL & ",Status"
strSQL = strSQL & ",CC"
strSQL = strSQL & ",Cam"
strSQL = strSQL & ",CR"
strSQL = strSQL & ") Values ("
strSQL = strSQL & "'" & strCnam & "',"
strSQL = strSQL & "'" & strJnam & "',"
strSQL = strSQL & "'" & strODate & "',"
strSQL = strSQL & "'" & strDDate & "',"
strSQL = strSQL & "'" & strFSize & "',"
strSQL = strSQL & "'" & strMtype & "',"
strSQL = strSQL & "'" & strSinfo &"'," 
strSQL = strSQL & "'" & strComments &"',"
strSQL = strSQL & "'OPEN',"
strSQL = strSQL & "'0PEN',"
strSQL = strSQL & "'0PEN',"
strSQL = strSQL & "'YES')"



Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic

db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "MemDB"


connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword





Set objDBConn = Server.CreateObject("ADODB.Connection")
  Set objRs = Server.CreateObject("ADODB.RecordSet")
objDBConn.Open connectstr






objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing

End if
%>
<script language="javascript">
<!--
alert("Information saved...please proceed browsing available jobs");
location.href="jobupdater.asp?pass=99999";
//-->
</script> 

