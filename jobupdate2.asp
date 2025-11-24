
<%

Dim objDBConn
Dim strSQL
Dim strJobnumber, strCID, strODate, strDDate, strJnam, strFSize, strMType, strStones, strRemarks, strStatus, strTracking, strCnam, strpwd3

strpwd3="99999"

strJobnumber=request.form("JNum")
strCID=request.form("CID")
strODate=request.form("ODate")
strDDate=request.form("DDate")
strJNam=request.form("JNam")
strFSize=request.form("FSize")
strMType=request.form("MType")
strStones=request.form("Stones")
strRemarks=request.form("Remarks")
strStatus=request.form("Status")
strTracking=request.form("Tracking")
strCNam=request.form("CNam")

strUserIP = Request.ServerVariables("REMOTE_ADDR") '글쓴이의 아이피를 가져오는 센스
vail="N"



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










 strSQL = "UPDATE MemDB SET Status='OPEN' WHERE RecNo = '"&strJobNumber&"'" 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0
Transitional//EN"> <HTML> <HEAD>
<TITLE> New Document </TITLE> </HEAD>

<script language="javascript">
<!--
alert("Job Re-opened...");
location.href="jobupdater2.asp?pass=<%=strpwd3%>";
//-->
</script> 

<BODY  >



</BODY>
</HTML>
