
<%

Dim strSQL, strRecno, strChooser, strTimInst

strTimInst=request.form("TimInst")

strChooser=request.form("button")

strRecno=request.form("recno")

If strChooser="approve" Then

strSQL = "UPDATE MemDB SET CR='YES', Tim='no' where RecNo='"&strRecno&"'" 
End If

If strChooser="cancel" then
strSQL = "UPDATE MemDB SET Status='Declined' WHERE RecNo='"&strRECno&"';"
End If
If strChooser="TIM" Then

strTimInst = Replace(strTimInst, "'", "''")
strSQL = "UPDATE MemDB SET CR='YES', Tim='Tim', TimDone='no', TimInst='"&strTimInst&"' where RecNo='"&strRecno&"'" 

End If

 



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
Set objDBConn = nothing
%>
<script language="javascript">
<!--
alert("Information saved...please proceed browsing available jobs");
location.href="jobupdater.asp?pass=99999";
//-->
</script> 

