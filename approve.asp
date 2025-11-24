
<%

Dim strSQL, strRecno, strConam



strRecno=request.form("Rec")
strConam=request.form("Company")





 



Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server


db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"



connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword


 strSQL = "UPDATE FMemDB SET Conam='"&strCoNam&"', Approved='yes' WHERE RecNo='"&strRecno&"';" 


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

