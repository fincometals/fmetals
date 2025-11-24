
<%

Dim objDBConn
Dim strSQL
Dim strJobnumber, strCID, strODate, strDDate, strJnam, strFSize, strMType, strStones, strRemarks, strStatus, strTracking, strCnam, strCom, strToday, Strcomments, strCC, strCAM, strPWD


strPWD="99999"

strJobnumber=request.form("JNum")
strCID=request.form("CID")
strODate=request.form("ODate")
strDDate=request.form("DDate")
strJNam=request.form("JNam")
strFSize=request.form("FSize")
strMType=request.form("MType")
strStones=request.form("Sinfo")
strRemarks=request.form("Comments")
strStatus=request.form("Status")
strTracking=request.form("Tracking")
strCNam=request.form("CNam")
strCom=request.form("edit")
StrComments=request.form("pcom")
strCC=request.form("cc")


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


If strCom = "Close Job" then



connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr




strToday=now()





 strSQL = "UPDATE MemDB SET Status='"&strToday&"' WHERE RecNo = '"&strJobNumber&"'" 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing

End If

If strCom = "Close CAM" then



connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr



 strSQL = "UPDATE MemDB SET CAM='Complete' WHERE RecNo = '"&strJobNumber&"'" 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing

End If

If strCom = "ReOpen CAM" then



connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr



 strSQL = "UPDATE MemDB SET CAM='0PEN' WHERE RecNo = '"&strJobNumber&"'" 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing

End If





If strCom = "Update" then





connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr


strRemarks = strComments &"<br>" &Date&"***"&strRemarks

strStones = Replace(strStones, "'", "''")
strRemarks = Replace(strRemarks, "'", "''")
strFsize = Replace(strFsize, "'", "''")
strMtype = Replace(strMtype, "'", "''")



 strSQL = "UPDATE MemDb SET DDate='"&strDdate&"', FSize='"&strFsize&"',MType='"&strMType&"',Sinfo='"&strstones&"',Comments='"&strRemarks&"' WHERE RecNo='"&strJobnumber&"';" 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing




End If

If strCom = "CAD Complete" Then

connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr




 strSQL = "UPDATE MemDB SET CC='Complete' WHERE RecNo = '"&strJobNumber&"'" 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing

End If

If strCom = "ReOpen CAD" Then

connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

Set objDBConn = Server.CreateObject("ADODB.Connection")

objDBConn.Open connectstr


strSQL = "UPDATE MemDb SET CC='0PEN', CAM='0PEN' WHERE RecNo='"&strJobnumber&"';" 

 

objDBConn.Execute strSQL
objDBConn.Close
Set objDBConn = Nothing

End if



%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0
Transitional//EN"> <HTML> <HEAD>
<TITLE></TITLE> </HEAD>

<script language="javascript">
<!--
alert("Information saved...");
location.href="jobupdater.asp?pass=<%=strpwd%>&cur=<%=strjobnumber%>#<%=strjobnumber%>";
//-->
</script> 

<BODY  >



</BODY>
</HTML>
