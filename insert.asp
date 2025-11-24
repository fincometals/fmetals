
<%

Dim objDBConn
Dim strSQL
Dim strJobnumber, strCID, strODate, strDDate, strJnam, st strMTrFSize,ype, strStones, strRemarks, strStatus, strTracking, strCnam, strUserIP


strJnam=request.form("Jnam")
strCnam=request.form("CNam")
strODate=request.form("ODate")
strDDate=request.form("DDate")
strFSize=request.form("FSize")
strMType=request.form("MType")
strStones=request.form("Sinfo")
strRemarks=request.form("Comments")
strStatus=request.form("SStatus")
strTracking=request.form("Tracking")




Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic

db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "InvDB"


connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr

 



strSQL = "Insert Into Orders ("
strSQL = strSQL & " ODate"
strSQL = strSQL & ",Jnam"
strSQL = strSQL & ",FSize"
strSQL = strSQL & ",MType"
strSQL = strSQL & ",Remarks"
strSQL = strSQL & ",Stones"
strSQL = strSQL & ",CNam"

strSQL = strSQL & ") Values ("
strSQL = strSQL & "'" & strOdate & "',"
strSQL = strSQL & "'" & strJnam & "',"
strSQL = strSQL & "'" & strFSize & "',"
strSQL = strSQL & "'" & strMType & "',"
strSQL = strSQL & "'" & strRemarks & "',"
strSQL = strSQL & "'" & strStones &"'," 
strSQL = strSQL & "'" & strCNam &"')"





objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0
Transitional//EN"> <HTML> <HEAD>
<TITLE> New Document </TITLE> </HEAD>

<script language=javascript>
function closemyself() {
 window.opener=self;
 window.close();
 //self.close();
}
</script>

<BODY background="vwConfig/Image1/$File/bg-logo.gif"
onLoad="setTimeout('closemyself()',10);" >

<table border="0">
 <tr>
<td>
<img src="vwConfig/Image1/$File/homepage-logo2.gif">
</td>
 </tr>
</table>
<p>  </p>
<table border="0" width="100%">
 <tr>
<td align="center"><font color=#330066
size="4"><strong>
Thank you for updating</strong></font>
</td>
 </tr>
 <tr>
<td align="center">
 
</td>
 </tr>
 <tr>
<td align="center"><font color=#330066>
(This window will close automatically)</font>
</td>
 </tr>
</table>

</BODY>
</HTML>
