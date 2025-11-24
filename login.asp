
 <% Option Explicit 
On Error Resume Next

 Dim strView

 strView=request.form("view")
 
 

Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic, strtyp

Dim strCanRec

strCanRec=request.form("cant")

db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"

connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword

dim strUname, strPwd, strPassChecker, strChooser, strCust, strPwd2, strCan



strCust=request.querystring("cust")


strChooser=request.form("button")

strUname=request.form("Uname")
strPwd=request.form("UPWD")

If strChooser="GO BACK" Or strChooser="edit" Or strChooser="REFRESH" or strChooser="Cancel " Or strChooser="view" Or strChooser="minimize" Then

strChooser="Log in"

End If 








Dim strDate, strCnam, strJnam, strODate, strFSize, strMType, strSinfo, strComments, strDDate, strCam, strChoo        

strDate = now()
strCnam = Request.Form("Cnam")
strJnam = Request.Form("Jnam")
strOdate = Request.Form("ODate")
strDDate= Request.Form("DDate")
strFsize = Request.Form("FSize")
strMtype = Request.Form("MType")
strSinfo = Request.Form("Sinfo")
strComments = Request.Form("Comments")
strChoo=request.querystring("choo")

Dim objDBConn, objRs
Dim strSQL, strUpRec

strUpRec=request.form("uprec")

If strChooser="" Then
strChoo="Log in"
strChooser=strChoo
End If


If strChooser="update" Then

Set objDBConn = Server.CreateObject("ADODB.Connection")
  Set objRs = Server.CreateObject("ADODB.RecordSet")
objDBConn.Open connectstr
strSinfo = Replace(strSinfo, "'", "''")
strFsize = Replace(strFsize, "'", "''")
strComments = Replace(strComments, "'", "''")
strMtype = Replace(strMtype, "'", "''")


strSQL = "UPDATE MemDB SET DDate='"&strDDate&"',FSize='"&strFSize&"',MType='"&strMType&"',SInfo='"&strSinfo&"',Instructions='"&strComments&"'  WHERE RECNO='"&strUpRec&"'" 

objDBConn.Execute strSQL
objDBConn.Close
objRS.Close
Set objRS=nothing
Set objDBConn = Nothing%>
<script language="javascript">
<!--
alert("Information saved...please proceed browsing available jobs");

//-->
</script> <%

strChooser="Log in"
End If

'start cancel pending
Dim objDBConn12, objRs12
Dim strSQL12, strbutton2
Dim strjnam1

strjnam1=request.form("jview")

strbutton2=request.form("button2")
If strchooser="cancel" Or strbutton2="view" Then 



Set objDBConn12 = Server.CreateObject("ADODB.Connection")
  Set objRs12 = Server.CreateObject("ADODB.RecordSet")
objDBConn12.Open connectstr

If strChooser="cancel" Then

strSQL12 = "UPDATE MemDB SET Status='User Cancelled' WHERE JNam='"&strCanRec&"';" 

objDBConn12.Execute strSQL12
objDBConn12.Close
objRS12.Close
Set objRS12=nothing
Set objDBConn12 = Nothing%>
<script language="javascript">
<!--
alert("Information saved...");

//-->
</script>
<%End If



If strbutton2="view" Then 
strSQL12 = "UPDATE MemDB SET Mread='yes' WHERE JNam='"&strJnam1&"';" 
objDBConn12.Execute strSQL12
objDBConn12.Close
objRS12.Close
Set objRS12=nothing
Set objDBConn12 = Nothing
End If

%>
 <%


strChooser="Log in"
End if



'end cancel pending


If strChooser = "ADD" Or strChooser="Add Instructions" Then
If strJnam="" Then%>
<script language="javascript">
<!--
alert("Job must have a name!  Your Job was not saved...please try again!");

//-->
</script> 


<%Else








Set objDBConn = Server.CreateObject("ADODB.Connection")
  Set objRs = Server.CreateObject("ADODB.RecordSet")
objDBConn.Open connectstr


If strChooser="ADD" Then



strSinfo = Replace(strSinfo, "'", "''")
strJnam = Replace(strJNam, "'", "''")
strFsize = Replace(strFsize, "'", "''")
strComments = Replace(strComments, "'", "''")
strMtype = Replace(strMtype, "'", "''")
strDDate = Replace(strDDate, "'", "''")




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
strSQL = strSQL & ",Change"
strSQL = strSQL & ") Values ("
strSQL = strSQL & "'" & strCnam & "',"
strSQL = strSQL & "'" & strJnam & "',"
strSQL = strSQL & "'" & strODate & "',"
strSQL = strSQL & "'" & strDDate & "',"
strSQL = strSQL & "'" & strFSize & "',"
strSQL = strSQL & "'" & strMtype & "',"
strSQL = strSQL & "'" & strSinfo &"',"
strSQL = strSQL & "'" & "<span style=style48>"&strDate&"***"& strComments &"</span>'," 
strSQL = strSQL & "'OPEN',"
strSQL = strSQL & "'0PEN',"
strSQL = strSQL & "'0PEN',"
strSQL = strSQL & "'NO',"
strSQL = strSQL & "'no')"

Else

Dim strInst, strInstOld, strNewInst, strCurin

strCurin=request.form("curin")
strInst=request.form("inst")
strInstOld=request.form("instold")

strNewInst=strInstOld&"<BR>"&"<span class=style48>"&strDate&"***"&strInst&"</span>"

strNewInst = Replace(strNewInst, "'", "''")

strSQL = "UPDATE MemDB SET Instructions='"&strNewInst&"',Change='yes' WHERE RECNO='"&strCurin&"'" 

End if





objDBConn.Execute strSQL
objDBConn.Close
objRS.Close
Set objRS=nothing
Set objDBConn = Nothing%>
<script language="javascript">
<!--
alert("Information saved...please proceed browsing available jobs");

//-->
</script> 
<%
End if




strChooser="Log in"

End If






If strCust <> "" Then

strChooser="Log in"

strUname=strCust
strPwd2=request.querystring("pwd2")




End If 





%>   
<html>   
<head>   


<title>JOB PANEL</title>   
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
body {
	background-image: url(http://www.fincometals.com/watches/gearb.jpg);
}
.style14 {
	font-family: Impact;
	font-size: 60px;
}
.style21 {
	font-family: "Calisto MT";
	color: #990033;
	font-size: 36px;
}
.style30 {
	font-size: 24px;
	color: #FFFFFF;
}
.style31 {
	color: #FF0000;
	font-weight: bold;
}
.style32 {color: #0033FF}
.style34 {font-size: 24px}
.style35 {color: #0000FF}
.style36 {color: #00FF00}
.style37 {
	font-family: Georgia, "Times New Roman", Times, serif
}
.style38 {font-size: 36px}
.style40 {
	color: #660000;
	font-size: 50px;
}
.style42 {font-family: Impact}
.style44 {color: #000000}
.style48 {font-size: 16px}
.style49 {font-size: 18px}
.style51 {font-size: 18}
.style52 {
	color: #660000;
	font-family: Impact;
}
.style54 {
	color: #660000;
	font-family: Impact;
	font-size: 18;
}
.style60 {color: #000033}
.style61 {font-family: Impact; color: #000033; }
.style62 {color: #660000}
.style66 {color: #000000; font-family: Impact; font-size: 16px; }
.style69 {
	font-size: 18px;
	font-family: "Times New Roman", Times, serif;
	font-weight: bold;
}
.style70 {font-family: "Times New Roman", Times, serif}
.style72 {color: #660000; font-weight: bold; }
.style73 {
	font-size: 45px;
	color: #009900;
	font-family: "Calisto MT";
}
.style74 {font-family: "Calisto MT"}
.style75 {font-family: "Calisto MT"; font-size: 36px; }
.style76 {font-weight: bold}
.style77 {
	color: #00CC00
}
.style79 {font-family: "Calisto MT"; font-size: 36px; font-weight: bold; }
-->
</style>
<script language=javascript>
document.onkeydown = function(){

if(window.event && window.event.keyCode == 116)
        { // Capture and remap F5
    window.event.keyCode = 505;
      }

if(window.event && window.event.keyCode == 505)
        { // New action for F5
    return false;
        // Must return false or the browser will refresh anyway
    }
}

</script>


</head>   
<body>   

<div align="center">
 
    <span class="style44">
    <%   
db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "InvDb"
connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword




Dim Conn7,strSQL7,objRec7, strmodel, strApproved, strConam
Set Conn7 = Server.Createobject("ADODB.Connection")   
Conn7.Open connectstr



If strChooser="Log in" then

strSQL7 = "SELECT * FROM FMemDB Where Eml='"& strUname&"'"

Set objRec7 = Server.CreateObject("ADODB.Recordset")   
objRec7.Open strSQL7, Conn7, 1,3   

strConam=objRec7.Fields("conam").Value


If objRec7.EOF Then


%>

    <script language="javascript">
<!--
alert("That username is not registered.");
location.href="customer.asp";
//-->
    </script> 

    <%

Else

strpasschecker=objRec7.Fields("Pwd").Value
strApproved=objRec7.Fields("Approved").Value

End If

If strPWD="" Then
strPWD=strpwd2
End If


If strpasschecker=strPWD AND strApproved="yes" Then%>
    </span>
  <p class="style42 style44 style73"><%=strCoNam%> Job Board</p>
  <table width="650" border="0" align="center">
  <tr>
    <td align="center"><span class="style76"><%=objRec7.Fields("ADR").Value%></span></td>
  </tr>
  <tr>
    <td align="center"><span class="style76"><%=objRec7.Fields("Cty").Value%>,&nbsp;<%=objRec7.Fields("Sta").Value%>&nbsp;<%=objRec7.Fields("Zip").Value%></span></td>
  </tr>
  <tr>
    <td align="center"><span class="style76">(<%=objRec7.Fields("ACD").Value%>)<%=objRec7.Fields("Phn").Value%></span></td>
    </tr>
    <tr>
    <td align="center"><span class="style76">Contact: <%=objRec7.Fields("FNam").Value%>&nbsp;<%=objRec7.Fields("LNam").Value%></span></td>
  </tr>
</table>

    <%
Dim strAJ

strAJ=request.form("AJ")

If strAJ="Add Job" then


%>
 <p class="style38 style74 style23 style14 style77"><strong>Add a Job</strong></p> 
 <span class="style66">First add your job, then upload pictures once listed in the &quot;pending jobs&quot; menu... </span>
 <table align="center" width="750" border="1" bgcolor="#666666">
    <tr>
      
               
              
   
     <td align="center" width="75">Job Name <span class="style16 style62">(required)</span></td>
    <td align="center" width="25">Order Date</td>
    <td align="center" width="25">Requested Due Date</td>
    <td align="center" width="15">Finger Size</td>
    <td align="center" width="25">Metal Type</td>
  </tr>
  <form action="login.asp" method="post" target="_self" id="form1">
  
  <tr>
 
    <input type="hidden" width="500" name="Cnam" value="<%=strConam%>"/>
     
          
    <td width="100" align="center"><div align="center"><input type="hidden" width="50" name="Jnum" value=""/><input type="text" width="200" name="JNam" value=""/></div></td>
    
    
    <td align="center" width="25"><div align="center"><input type="text" width="25" name="ODate" value="<%=strdate%>"/></div></td>
    <td align="center" width="25"><div align="center"><input type="text" width="25" name="DDate" value=""/></div></td>
    <td align="center" width="15"><div align="center"><input type="text" width="15" name="FSize" value=""/></div></td>
    <td align="center" width="25"><div align="center"><input type="text" width="25" name="MType" value=""/></div></td>
</tr>
    
    
    <tr>
     <td align="center" width="100">Stone Information</td>
      <td align="center" colspan="5"><div align="center"><textarea cols="80" name="Sinfo" value="" form="form1"></textarea>
    
  </div></td>
    </tr>
    <tr>
     <td align="center" width="100">Instructions <input type="hidden" name="Uname" value="<%=strUname%>"><input type="hidden" name="UPWD" value="<%=strPWD%>"><input type="submit" name="button" value="ADD" /></td>
    <td align="center" Colspan="5"><div align="center">
            <textarea cols="80" name="Comments" value="" form="form1"></textarea>
          </div></td>
    </tr>
    </form>
  </table>



<%

else%>
<form action=login.asp?choo=Login target="_self" method="post">
  <input type="hidden" name="upwd" value="<%=strPwd%>">
  <input type="hidden" name="uname" value="<%=strUname%>">

  <input type="submit" name="AJ" value="Add Job">
  </form>
</p>
</div>


<%


End if

Dim Conn8,strSQL8,objRec8
Set Conn8 = Server.Createobject("ADODB.Connection")   
Conn8.Open connectstr


strSQL8 = "SELECT * FROM MEMDB where Cnam='"&strConam&"' AND Status='OPEN' AND CR='NO'"

Set objRec8 = Server.CreateObject("ADODB.Recordset")   
objRec8.Open strSQL8, Conn8, 1,3

If objRec8.EOF then

else%>

<div align="center" class="style14 style20 style35 style44">
  <p class="style37"><div class="style40 style42"><span class="style75"><a name="pending">Pending Jobs</a>
  (not yet accepted)</span></div>
  <div class="style66 style44 style48">These are jobs that have not yet been accepted.This order will begin processing once this record has moved to &quot;Open Jobs&quot;</div>
  <form action=login.asp#pending method=post target="_self">
  <input type="hidden" name="upwd" value="<%=strPwd%>">
  <input type="hidden" name="uname" value="<%=strUname%>">

  <input type="submit" name="button" value="REFRESH">
  </form>
  
  
 </p>
</div>



<span class="style44">
<% 

End If

Dim strBut, strEno, strEnoChk



strEno=request.form("eno")

strBut=request.form("button")

Do While Not objRec8.EOF
strEnoChk=objRec8.Fields("JNam").Value
%>
</span>
<table width="980" border="1" align="center" bgcolor="#666666">
    <tr align="center">
      
               <td width="86" align="center" class="style44"><div align="center" class="style34 style42"><span class="style51 style62"><span class="style74">Job</span>  <a name="<%=objRec8.Fields("RecNo").Value%>" class="style49 style49"><strong><%=objRec8.Fields("RecNo").Value%></strong></a></span></div></td>
              
    
 <td width="146" align="center" class="style44">Order Date: <strong><%=objRec8.Fields("ODate").Value%></strong></td>
 <td width="345" align="center" class="style34 style44"><span class="style54"><%=objRec8.Fields("JNam").Value%></span></td>
   
     <td width="146" align="center" class="style44">Requested Due Date: <strong><%=objRec8.Fields("DDate").Value%></strong></td>
   
    <td width="144" align="center" class="style44">Size # <strong><%=objRec8.Fields("FSize").Value%></strong></td>
    <td width="144" align="center" class="style44"><div>Metal: <strong><%=objRec8.Fields("MType").Value%></strong></div></td>
    <td class="style44"></td>
  </tr>
<%'start normal



If strEno=strEnoChk Then
'start ENO IF

%>
<form id="userform" action=login.asp#<%=objRec8.Fields("RecNo").Value%> method="post" target="_self">
  <span class="style44">
<input type="hidden" name="Uname" value="<%=strUname%>">
<input type="hidden" name="Upwd" value="<%=strPwd%>">
<input type="hidden" name="Uprec" value="<%=objRec8.Fields("RecNo").Value%>">
  </span>
  <tr width="980">
<td width="40" align="center" class="style44"><div align="center" class="style31 style49 style49 style44"><span class="style42 style34 style38 style49 style49"><a name="<%=objRec8.Fields("RecNo").Value%>"><%=objRec8.Fields("RecNo").Value%></a></span></div></td>
        <td width="60" align="center" class="style44"><div align="center"><strong><%=objRec8.Fields("ODate").Value%></strong></div></td>
		<td align="center" width="150"><div align="center" class="style23 style42 style35 style51 style44"><%=objRec8.Fields("JNam").Value%></div></td>
    <%
	DD=objRec8.Fields("DDate").Value%>
	




	
    <td align="center" width="60"><div align="center"class="style19 style16"><strong><input type="text" name="DDate" value="<%=objRec8.Fields("DDate").Value%>"></strong></div></td>

    <td align="center" width="50"><div align="center"><strong><input type="text" name="FSize" value="<%=objRec8.Fields("FSize").Value%>"></strong></div></td>
    <td align="center" width="50"><div align="center"><strong><input type="text" name="Mtype" value="<%=objRec8.Fields("MType").Value%>"></strong></div></td>
 </tr>
<tr>

 <td align="center" width="40"><strong>Stone info</strong></td>
  <td align="center"colspan="7"><div align="center" class="style28 style32 style42 style34" ><textarea rows="4" cols="125" name="Sinfo" form="userform"value="">
<%=objRec8.Fields("SInfo").Value%></textarea></div></td> 
 </tr>

 <tr>
 <td align="center" width="40"><strong>Instructions:<br>
 



 
 </strong></td>   
 <td align="center" colspan="6" width="880"><div align="center" class="style28 style42 style34 style44"><textarea rows="4" cols="125" name="Comments" form="userform"value="">
<%=objRec8.Fields("Instructions").Value%></textarea></div></td> 

<td><input type="submit" name="button" value="update"></form> 


<form action=login.asp#<%=objRec8.Fields("JNam").Value%> method=post target="_self">
  <input type="hidden" name="upwd" value="<%=strPwd%>">
  <input type="hidden" name="uname" value="<%=strUname%>">

  <input type="submit" name="button" value="Cancel ">
  </form></td>
  
  </tr>




<%
'end 
Else


%>




<tr>

 <td align="center" width="40"><strong>Stone info</strong></td>
  <td align="center"colspan="7"><div align="center" class="style70 style48 style44 style28" ><strong><%=objRec8.Fields("Sinfo").Value%></strong></div></td> 
 </tr>

 <tr>
 <td align="center" width="40"><strong>Instructions:




 
 </strong> </td>   
 <td align="center" colspan="7" width="880"><div align="center" class="style70 style48 style44 style28"><strong><%=objRec8.Fields("Instructions").Value%></strong></div></td> 
  </tr>
<tr>



<%
'start picture show


Dim Conn6,strSQL6,objRec6, curjob, curfile
Set Conn6 = Server.Createobject("ADODB.Connection")   

curjob=objRec8.Fields("RecNo").Value
Conn6.Open connectstr


strSQL6 = "SELECT filenam from Uploads where job='"&curjob&"'"

Set objRec6 = Server.CreateObject("ADODB.Recordset")   
objRec6.Open strSQL6, Conn6, 1,3

If objRec6.EOF Then
%>
<TD colspan="8"><%Response.write ("No images uploaded....")%></td></tr>
<tr><td colspan="8">
<form action=login.asp#<%=objRec8.Fields("RecNo").Value%> method="post" target="_self"><A Href="uploadtester.asp?recno=<%=objRec8.Fields("RecNo").Value%>&jnam=<%=objRec8.Fields("JNam").Value%>&pwd=<%=strPwd%>&eml=<%=strUname%>" target="_self" class="style42 style35 style60">Upload Photos</a> 
<input type="hidden" name="Uname" value="<%=strUname%>">
<input type="hidden" name="Upwd" value="<%=strPwd%>">
<input type="hidden" name="Eno" value="<%=objRec8.Fields("Jnam").Value%>">
<input type="hidden" name="cant" value="<%=objRec8.Fields("Jnam").Value%>">
<input type="submit" name="button" value="edit">
<input type="submit" name="button" value="cancel">
</form></td></tr>

<%
Else

%>

<TD align="center" colspan="8">

<%

Do While Not objRec6.EOF

curfile=objRec6.Fields("filenam").Value
%>




<a href="./uploads/<%=curfile%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','./uploads/<%=curfile%>',1)"><img name="Image2" border="0" height="60" width="80" src="./uploads/<%=curfile%>"></a>

<%objRec6.MoveNext   
Loop 
objRec6.Close()   
Conn6.Close()   
Set objRec6 = Nothing  
Set Conn6 = Nothing 
%>
</td></Tr>
<tr><td colspan="8">
<form action=login.asp#<%=objRec8.Fields("RecNo").Value%> method="post" target="_self"><A Href="uploadtester.asp?recno=<%=objRec8.Fields("RecNo").Value%>&jnam=<%=objRec8.Fields("JNam").Value%>&pwd=<%=strPwd%>&eml=<%=strUname%>" target="_self" class="style42 style35 style60">Upload Photos</a> 
<input type="hidden" name="Uname" value="<%=strUname%>">
<input type="hidden" name="Upwd" value="<%=strPwd%>">
<input type="hidden" name="Eno" value="<%=objRec8.Fields("Jnam").Value%>">
<input type="hidden" name="cant" value="<%=objRec8.Fields("Jnam").Value%>">
<input type="submit" name="button" value="edit">
<input type="submit" name="button" value="cancel">
</form></td></tr>




<%
End if


'end picture show

%>

       




<% 
End If 
objRec8.MoveNext   
Loop 
objRec8.Close()   
Conn8.Close()   
Set objRec8 = Nothing  
Set Conn8 = Nothing %>

</table>





<%

'end normal
Dim Conn2,strSQL2,objRec2
Set Conn2 = Server.Createobject("ADODB.Connection")   
Conn2.Open connectstr


strSQL2 = "SELECT * FROM MEMDB where Cnam='"&strConam&"' AND Status='OPEN' AND CR='YES'"

Set objRec2 = Server.CreateObject("ADODB.Recordset")   
objRec2.Open strSQL2, Conn2, 1,3

If objrec2.eof Then

Else


%>
<div align="center" class="style14 style20 style35"><p class="style37"><a name="open" class="style61"><span class="style79">Open Jobs</span>
  <form action=login.asp#open method=post target="_self">
  <input type="hidden" name="upwd" value="<%=strPwd%>">
  <input type="hidden" name="uname" value="<%=strUname%>">

  <input type="submit" name="button" value="REFRESH">
  </form></a>

</p></div>


<% End If

Do While Not objRec2.EOF

%>
<table width="980" border="1" align="center" bgcolor="#666666">
    <tr align="center">
      <td>
      
      
      <form action="login.asp#<%=objRec2.Fields("JNam").Value%>" method="post" target="_self">
	  <input type="hidden" name="upwd" value="<%=strPwd%>">
  <input type="hidden" name="uname" value="<%=strUname%>">
   <input type="hidden" name="jview2" value="<%=objRec2.Fields("JNam").Value%>">
   <input type="hidden" name="view" value="view">
    
<%
dim viewcheck, viewcheck2, viewcheck3

viewcheck3=request.form("checker")

viewcheck2=request.form("jview2")
strview=request.form("button")
viewcheck=objRec2.Fields("JNam").Value

if strview="view" and viewcheck=viewcheck2  then


%>
 
  
    <input type="submit" value="minimize" name="button">
  
  
  <%else%>
  <input type="hidden" name="jview" value="<%=objRec2.Fields("JNam").Value%>">
  <input type="submit" value="view" name="button2">
  <%
end if%>
      </form><%
dim strmess, strmess2

strmess2="no"

strmess=objRec2.Fields("Mread").Value
if strmess=strmess2 then
%>
<span class="style72">New Message</span>
<%else%>


<%end if%>	  </td>
               <td align="center" width="135"><span class="style35 style42 style60">Job # <a name="<%=objRec2.Fields("RecNo").Value%>"><strong><%=objRec2.Fields("RecNo").Value%></a><a name="<%=objRec2.Fields("Jnam").Value%>"></a></span><span class="style60"></strong></span></td>
              
    
 <td align="center" width="89">Order Date: <strong><%=objRec2.Fields("ODate").Value%></strong></td>
 <td align="center" width="111"><span class="style52"><%=objRec2.Fields("JNam").Value%></span></td>
   
     <td align="center" width="89">Requested Due Date: <strong><%=objRec2.Fields("DDate").Value%></strong></td>
   
    <td align="center" width="75">Size # <strong><%=objRec2.Fields("FSize").Value%></strong></td>
    <td align="center" width="75"><div>Metal: <strong><%=objRec2.Fields("Mtype").Value%> </strong></div></td>
    
    
    
   <td align="center" width="139">CAD: <strong><%=objRec2.Fields("CC").Value%> </strong></td>
<td align="center" width="152">CAM: <strong><%=objRec2.Fields("CAM").Value%></strong></td>
  </tr>

<%' start view check
Dim StrJview, strJviewChk

strJviewChk=objRec2.Fields("JNam").Value

strJview=request.form("jview")


If  strJviewChk=strJview Then




%>
     

<tr bgcolor="#CCCCCC">
<td></td>
 <td align="center" width="40"><strong>Stone info</strong></td>
  <td align="center"colspan="7"><div align="center" class="style44 style28" ><strong><%=objRec2.Fields("Sinfo").Value%></strong></div></td> 
 </tr>

 <tr bgcolor="#CCCCCC">
 <td></td>
 <td align="center" width="40"><strong>Instructions:</strong><br><A Href="uploadtester.asp?recno=<%=objRec2.Fields("RecNo").Value%>&jnam=<%=objRec2.Fields("JNam").Value%>&pwd=<%=strPwd%>&eml=<%=strUname%>" target="_self" class="style42 style60">Upload Photos</a></td>   
 <td align="center" colspan="7"><div align="center" class="style28 style44 style42 style69"><%=objRec2.Fields("Instructions").Value%></div> <br>
 <%
 Dim strCK

 strCK=objRec2.Fields("Change").Value

 If strCK="yes" Then%>
 Latest Instructions:  <span class="style52">NOT READ</span><BR>
 <BR>

<%End If

If strCK="no" Then
dim strResp

strResp=objRec2.Fields("Resp").Value
%>
 <span class="style36"><span class="style44">Latest Instructions:</span> <span class="style42">RECEIVED</span></span>
<%
 end if
 
%>
 </span></td>
 </tR>
 <TR bgcolor="#CCCCCC">
 <td colspan="8"><textarea rows="4" cols="100" name="Inst" form="userform2"value="">
Add instructions here...</textarea>
 </td><TD><form action="login.asp#<%=objRec2.Fields("JNam").Value%>" target="_self" method="post" id="userform2">
 <input type="hidden" name="UName" value="<%=strUname%>">
 <input type="hidden" name="UPwd" value="<%=strPwd%>">
 <input type="hidden" name="instold" value="<%=objRec2.Fields("Instructions").Value%>">
 <input type="hidden" name="JNam" value="<%=objRec2.Fields("JNam").Value%>">
 <input type="hidden" name="Jview" value="<%=objRec2.Fields("JNam").Value%>">
 <input type="hidden" name="Curin" value="<%=objRec2.Fields("RecNo").Value%>">
 
 <input type="submit" name="button" value="Add Instructions"><br>
 You can not edit an accepted job but you may add instructions
 (request a change)
 </form> </TD></tr>
<%
Dim Conn16,strSQL16,objRec16, curjob1, curfile1
Set Conn16 = Server.Createobject("ADODB.Connection")   

curjob1=objRec2.Fields("RecNo").Value
Conn16.Open connectstr


strSQL16 = "SELECT filenam from Uploads where job='"&curjob1&"'"

Set objRec16 = Server.CreateObject("ADODB.Recordset")   
objRec16.Open strSQL16, Conn16, 1,3

If objRec16.EOF Then
%>
<tr bgcolor="#CCCCCC" align="left">
<TD colspan="9" align="left"><%Response.write ("No images uploaded....")%></td>
</tr>
<%
Else%>
<Tr>
<TD align="center" colspan="9">
<%
Do While Not objRec16.EOF

curfile1=objRec16.Fields("filenam").Value
%>

<a href="./uploads/<%=curfile1%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','./uploads/<%=curfile1%>',1)"><img name="Image2" border="0" height="60" width="80" src="./uploads/<%=curfile1%>"></a>

<%objRec16.MoveNext   
Loop 
objRec16.Close()   
Conn16.Close()   
Set objRec16 = Nothing  
Set Conn16 = Nothing 

End if


'end picture show


Else


End if
'end view check

objRec2.MoveNext   
Loop 
objRec2.Close()   
Conn2.Close()   
Set objRec2 = Nothing  
Set Conn2 = Nothing 

Dim Conn3,strSQL3,objRec3
Set Conn3 = Server.Createobject("ADODB.Connection")   
Conn3.Open connectstr


strSQL3 = "SELECT * FROM MEMDB where Cnam='"&strConam&"' AND Status<>'OPEN'"
Set objRec3 = Server.CreateObject("ADODB.Recordset")   
objRec3.Open strSQL3, Conn3%>
</td>
</tr> </table>

<%If objrec3.eof Then 

Else
%>
<div align="center" class="style14 style21">Completed Jobs</div>

<table align="center" width="600" border="1" bgcolor="#666666">
    <tr>
      <td width="40">	  </td>
               <td align="center" width="40"><strong>Job #</strong></td>
              
    
 <td align="center" width="150"><strong>Job Name</strong></td>
    <td align="center" width="60"><strong>Order Date</strong></td>
     
  
    <td align="center" width="40"><strong>Completed</strong></td>
  </tr>
      
<%End If 
Do While Not objRec3.EOF





%>   

   
     
      
      
          <tr >
    
      
	  <td width="40">

<form action="custdetails.asp" action="post" target="_self"> 
	  <input type="hidden" name="reco" value="<%=objRec3.Fields("RecNo").Value%>">
<input type="hidden" name="upwd" value="<%=strPwd%>">
<input type="hidden" name="uname" value="<%=strUname%>">

<input type="submit" value="view">	</form>  </td>
            <td width="40" align="center"><div align="center"><a name="<%=objRec3.Fields("RecNo").Value%>"><strong><%=objRec3.Fields("RecNo").Value%></a></div></td>
          
  
        <td align="center" width="150"><div align="center"><%=objRec3.Fields("JNam").Value%></div></td>
    <td align="center" width="60"><div align="center"><%=objRec3.Fields("ODate").Value%></div></td>
       <td align="center" width="40"><div align="center"><%=objRec3.Fields("Status").Value%></div></td>
      </tr>

    <%


objRec3.MoveNext   
Loop 
objRec3.Close()   
Conn3.Close()   
Set objRec3 = Nothing  
Set Conn3 = Nothing 


%></table>


<%


Else If strpasschecker=strPWD Then %>
  
  
  <script language="javascript">
<!--
alert("Your log in information was correct, but your account is not yet active.  Please try again later.");
location.href="customer.asp";
//-->
    </script> 
  
  <%Else%>
  <script language="javascript">
<!--
alert("Your log in information is incorrect.");
location.href="customer.asp";
//-->
    </script> 
  
  
  
  <%
End if
End if



Else



%>
</p>
<form action="dupchk.asp" method="post" accept-charset="utf-8" >
  <p>
    <input type="hidden" name="formID" value="31796863620159" />
  </p>
  <p class="style30">Please register to gain access to your job board</p>
  <div class="form-all">
    <ul class="form-section"><li class="form-line form-line-column" id="id_1"><div id="cid_1" class="form-input-wide"><span class="form-sub-label-container"><input class="form-textbox" type="text" size="10" name="fname" id="first_1" />
            <label class="form-sub-label" for="first_1" id="sublabel_first"> First Name </label></span><span class="form-sub-label-container"><input class="form-textbox" type="text" size="15" name="lname" id="last_1" />
            <label class="form-sub-label" for="last_1" id="sublabel_last"> Last Name </label></span>        </div>
      </li>
      <li class="form-line form-line-column" id="id_5">
        <div id="cid_5" class="form-input-wide">
          <input type="text" class=" form-textbox" data-type="input-textbox" id="input_5" name="pass" size="20" />
        </div>
      </li>
      <li class="form-line form-line-column" id="id_6">Password</li>
      <li class="form-line form-line-column">
        <div id="cid_6" class="form-input-wide">
          <input type="text" class=" form-textbox" data-type="input-textbox" id="input_6" name="pass2" size="20" />
        </div>
      </li>
      <li class="form-line form-line-column" id="id_9">
        <label class="form-label-top" id="label_9" for="input_9">        Password (confirm) 
       <br>
        </label>
        <div id="cid_9" class="form-input-wide"><span class="form-sub-label-container"><input class="form-textbox" type="tel" name="acod" id="input_9_area" size="3">
            -
            <label class="form-sub-label" for="input_9_area" id="sublabel_area"> Area Code </label></span><span class="form-sub-label-container"><input class="form-textbox" type="tel" name="number" id="input_9_phone" size="8">
            <label class="form-sub-label" for="input_9_phone" id="sublabel_phone"> Phone Number </label></span>        </div>
      </li>
      <li class="form-line form-line-column" id="id_3">
        <label class="form-label-top" id="label_3" for="input_3"></label>
        <div id="cid_3" class="form-input-wide">
          <input type="email" class=" form-textbox validate[Email]" id="input_3" name="email" size="30" />
        </div>
      </li>
      <li class="form-line form-line-column" id="id_7">
        <label class="form-label-top" id="label_3" for="label">E-mail </label>

        <label class="form-label-top" id="label_7" for="input_7"> <br>
        </label>
        <div id="cid_7" class="form-input-wide">
          <input type="email" class=" form-textbox validate[Email]" id="input_7" name="email2" size="30" />
        </div>
      </li>
      <li class="form-line" id="id_8"><span class="form-line form-line-column">
        <label class="form-label-top" id="label_7" for="label">Confirm E-mail </label>
      </span>
        <div id="cid_8" class="form-input">
          <table summary="" class="form-address-table" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td colspan="2"><span class="form-sub-label-container"><input class="form-textbox form-address-line" type="text" name="CoNam" id="input_8_addr_line1" />
                  <label class="form-sub-label" for="input_8_addr_line1" id="sublabel_8_addr_line1"> Company Name</label>
              </span>              </td>
            </tr>
            <tr>
              <td colspan="2"><span class="form-sub-label-container"><input class="form-textbox form-address-line" type="text" name="ADR" id="input_8_addr_line2" size="46" />
                  <label class="form-sub-label" for="input_8_addr_line2" id="sublabel_8_addr_line2">Street Address</label>
              </span>              </td>
            </tr>
            <tr>
              <td width="50%"><span class="form-sub-label-container"><input class="form-textbox form-address-city" type="text" name="Cty" id="input_8_city" size="21" />
                  <label class="form-sub-label" for="input_8_city" id="sublabel_8_city"> City </label></span>              </td>
              <td><span class="form-sub-label-container"><input class="form-textbox form-address-state" type="text" name="Sta" id="input_8_state" size="22" />
                  <label class="form-sub-label" for="input_8_state" id="sublabel_8_state"> State / Province </label></span>              </td>
            </tr>
            <tr>
              <td width="50%" function zip(){var iterator=Prototype.K,args=$A(arguments);if(Object.isFunction(args.last())) iterator=args.pop();var collections=[this].concat(args).map($A);return this.map(function(value,index){return iterator(collections.pluck(index));});}><span class="form-sub-label-container"><input class="form-textbox form-address-postal" type="text" name="Zip" id="input_8_postal" size="10" />
                  <label class="form-sub-label" for="input_8_postal" id="sublabel_8_postal"> Postal / Zip Code </label></span>
              </td>
              <td><span class="form-sub-label-container"><select class="form-dropdown form-address-country" name="Cun" id="input_8_country">
                    <option value="" selected> Please Select </option>
                    <option value="United States"> United States </option>
                    <option value="Afghanistan"> Afghanistan </option>
                    <option value="Albania"> Albania </option>
                    <option value="Algeria"> Algeria </option>
                    <option value="American Samoa"> American Samoa </option>
                    <option value="Andorra"> Andorra </option>
                    <option value="Angola"> Angola </option>
                    <option value="Anguilla"> Anguilla </option>
                    <option value="Antigua and Barbuda"> Antigua and Barbuda </option>
                    <option value="Argentina"> Argentina </option>
                    <option value="Armenia"> Armenia </option>
                    <option value="Aruba"> Aruba </option>
                    <option value="Australia"> Australia </option>
                    <option value="Austria"> Austria </option>
                    <option value="Azerbaijan"> Azerbaijan </option>
                    <option value="The Bahamas"> The Bahamas </option>
                    <option value="Bahrain"> Bahrain </option>
                    <option value="Bangladesh"> Bangladesh </option>
                    <option value="Barbados"> Barbados </option>
                    <option value="Belarus"> Belarus </option>
                    <option value="Belgium"> Belgium </option>
                    <option value="Belize"> Belize </option>
                    <option value="Benin"> Benin </option>
                    <option value="Bermuda"> Bermuda </option>
                    <option value="Bhutan"> Bhutan </option>
                    <option value="Bolivia"> Bolivia </option>
                    <option value="Bosnia and Herzegovina"> Bosnia and Herzegovina </option>
                    <option value="Botswana"> Botswana </option>
                    <option value="Brazil"> Brazil </option>
                    <option value="Brunei"> Brunei </option>
                    <option value="Bulgaria"> Bulgaria </option>
                    <option value="Burkina Faso"> Burkina Faso </option>
                    <option value="Burundi"> Burundi </option>
                    <option value="Cambodia"> Cambodia </option>
                    <option value="Cameroon"> Cameroon </option>
                    <option value="Canada"> Canada </option>
                    <option value="Cape Verde"> Cape Verde </option>
                    <option value="Cayman Islands"> Cayman Islands </option>
                    <option value="Central African Republic"> Central African Republic </option>
                    <option value="Chad"> Chad </option>
                    <option value="Chile"> Chile </option>
                    <option value="People's Republic of China"> People's Republic of China </option>
                    <option value="Republic of China"> Republic of China </option>
                    <option value="Christmas Island"> Christmas Island </option>
                    <option value="Cocos (Keeling) Islands"> Cocos (Keeling) Islands </option>
                    <option value="Colombia"> Colombia </option>
                    <option value="Comoros"> Comoros </option>
                    <option value="Congo"> Congo </option>
                    <option value="Cook Islands"> Cook Islands </option>
                    <option value="Costa Rica"> Costa Rica </option>
                    <option value="Cote d'Ivoire"> Cote d'Ivoire </option>
                    <option value="Croatia"> Croatia </option>
                    <option value="Cuba"> Cuba </option>
                    <option value="Cyprus"> Cyprus </option>
                    <option value="Czech Republic"> Czech Republic </option>
                    <option value="Denmark"> Denmark </option>
                    <option value="Djibouti"> Djibouti </option>
                    <option value="Dominica"> Dominica </option>
                    <option value="Dominican Republic"> Dominican Republic </option>
                    <option value="Ecuador"> Ecuador </option>
                    <option value="Egypt"> Egypt </option>
                    <option value="El Salvador"> El Salvador </option>
                    <option value="Equatorial Guinea"> Equatorial Guinea </option>
                    <option value="Eritrea"> Eritrea </option>
                    <option value="Estonia"> Estonia </option>
                    <option value="Ethiopia"> Ethiopia </option>
                    <option value="Falkland Islands"> Falkland Islands </option>
                    <option value="Faroe Islands"> Faroe Islands </option>
                    <option value="Fiji"> Fiji </option>
                    <option value="Finland"> Finland </option>
                    <option value="France"> France </option>
                    <option value="French Polynesia"> French Polynesia </option>
                    <option value="Gabon"> Gabon </option>
                    <option value="The Gambia"> The Gambia </option>
                    <option value="Georgia"> Georgia </option>
                    <option value="Germany"> Germany </option>
                    <option value="Ghana"> Ghana </option>
                    <option value="Gibraltar"> Gibraltar </option>
                    <option value="Greece"> Greece </option>
                    <option value="Greenland"> Greenland </option>
                    <option value="Grenada"> Grenada </option>
                    <option value="Guadeloupe"> Guadeloupe </option>
                    <option value="Guam"> Guam </option>
                    <option value="Guatemala"> Guatemala </option>
                    <option value="Guernsey"> Guernsey </option>
                    <option value="Guinea"> Guinea </option>
                    <option value="Guinea-Bissau"> Guinea-Bissau </option>
                    <option value="Guyana"> Guyana </option>
                    <option value="Haiti"> Haiti </option>
                    <option value="Honduras"> Honduras </option>
                    <option value="Hong Kong"> Hong Kong </option>
                    <option value="Hungary"> Hungary </option>
                    <option value="Iceland"> Iceland </option>
                    <option value="India"> India </option>
                    <option value="Indonesia"> Indonesia </option>
                    <option value="Iran"> Iran </option>
                    <option value="Iraq"> Iraq </option>
                    <option value="Ireland"> Ireland </option>
                    <option value="Israel"> Israel </option>
                    <option value="Italy"> Italy </option>
                    <option value="Jamaica"> Jamaica </option>
                    <option value="Japan"> Japan </option>
                    <option value="Jersey"> Jersey </option>
                    <option value="Jordan"> Jordan </option>
                    <option value="Kazakhstan"> Kazakhstan </option>
                    <option value="Kenya"> Kenya </option>
                    <option value="Kiribati"> Kiribati </option>
                    <option value="North Korea"> North Korea </option>
                    <option value="South Korea"> South Korea </option>
                    <option value="Kosovo"> Kosovo </option>
                    <option value="Kuwait"> Kuwait </option>
                    <option value="Kyrgyzstan"> Kyrgyzstan </option>
                    <option value="Laos"> Laos </option>
                    <option value="Latvia"> Latvia </option>
                    <option value="Lebanon"> Lebanon </option>
                    <option value="Lesotho"> Lesotho </option>
                    <option value="Liberia"> Liberia </option>
                    <option value="Libya"> Libya </option>
                    <option value="Liechtenstein"> Liechtenstein </option>
                    <option value="Lithuania"> Lithuania </option>
                    <option value="Luxembourg"> Luxembourg </option>
                    <option value="Macau"> Macau </option>
                    <option value="Macedonia"> Macedonia </option>
                    <option value="Madagascar"> Madagascar </option>
                    <option value="Malawi"> Malawi </option>
                    <option value="Malaysia"> Malaysia </option>
                    <option value="Maldives"> Maldives </option>
                    <option value="Mali"> Mali </option>
                    <option value="Malta"> Malta </option>
                    <option value="Marshall Islands"> Marshall Islands </option>
                    <option value="Martinique"> Martinique </option>
                    <option value="Mauritania"> Mauritania </option>
                    <option value="Mauritius"> Mauritius </option>
                    <option value="Mayotte"> Mayotte </option>
                    <option value="Mexico"> Mexico </option>
                    <option value="Micronesia"> Micronesia </option>
                    <option value="Moldova"> Moldova </option>
                    <option value="Monaco"> Monaco </option>
                    <option value="Mongolia"> Mongolia </option>
                    <option value="Montenegro"> Montenegro </option>
                    <option value="Montserrat"> Montserrat </option>
                    <option value="Morocco"> Morocco </option>
                    <option value="Mozambique"> Mozambique </option>
                    <option value="Myanmar"> Myanmar </option>
                    <option value="Nagorno-Karabakh"> Nagorno-Karabakh </option>
                    <option value="Namibia"> Namibia </option>
                    <option value="Nauru"> Nauru </option>
                    <option value="Nepal"> Nepal </option>
                    <option value="Netherlands"> Netherlands </option>
                    <option value="Netherlands Antilles"> Netherlands Antilles </option>
                    <option value="New Caledonia"> New Caledonia </option>
                    <option value="New Zealand"> New Zealand </option>
                    <option value="Nicaragua"> Nicaragua </option>
                    <option value="Niger"> Niger </option>
                    <option value="Nigeria"> Nigeria </option>
                    <option value="Niue"> Niue </option>
                    <option value="Norfolk Island"> Norfolk Island </option>
                    <option value="Turkish Republic of Northern Cyprus"> Turkish Republic of Northern Cyprus </option>
                    <option value="Northern Mariana"> Northern Mariana </option>
                    <option value="Norway"> Norway </option>
                    <option value="Oman"> Oman </option>
                    <option value="Pakistan"> Pakistan </option>
                    <option value="Palau"> Palau </option>
                    <option value="Palestine"> Palestine </option>
                    <option value="Panama"> Panama </option>
                    <option value="Papua New Guinea"> Papua New Guinea </option>
                    <option value="Paraguay"> Paraguay </option>
                    <option value="Peru"> Peru </option>
                    <option value="Philippines"> Philippines </option>
                    <option value="Pitcairn Islands"> Pitcairn Islands </option>
                    <option value="Poland"> Poland </option>
                    <option value="Portugal"> Portugal </option>
                    <option value="Puerto Rico"> Puerto Rico </option>
                    <option value="Qatar"> Qatar </option>
                    <option value="Romania"> Romania </option>
                    <option value="Russia"> Russia </option>
                    <option value="Rwanda"> Rwanda </option>
                    <option value="Saint Barthelemy"> Saint Barthelemy </option>
                    <option value="Saint Helena"> Saint Helena </option>
                    <option value="Saint Kitts and Nevis"> Saint Kitts and Nevis </option>
                    <option value="Saint Lucia"> Saint Lucia </option>
                    <option value="Saint Martin"> Saint Martin </option>
                    <option value="Saint Pierre and Miquelon"> Saint Pierre and Miquelon </option>
                    <option value="Saint Vincent and the Grenadines"> Saint Vincent and the Grenadines </option>
                    <option value="Samoa"> Samoa </option>
                    <option value="San Marino"> San Marino </option>
                    <option value="Sao Tome and Principe"> Sao Tome and Principe </option>
                    <option value="Saudi Arabia"> Saudi Arabia </option>
                    <option value="Senegal"> Senegal </option>
                    <option value="Serbia"> Serbia </option>
                    <option value="Seychelles"> Seychelles </option>
                    <option value="Sierra Leone"> Sierra Leone </option>
                    <option value="Singapore"> Singapore </option>
                    <option value="Slovakia"> Slovakia </option>
                    <option value="Slovenia"> Slovenia </option>
                    <option value="Solomon Islands"> Solomon Islands </option>
                    <option value="Somalia"> Somalia </option>
                    <option value="Somaliland"> Somaliland </option>
                    <option value="South Africa"> South Africa </option>
                    <option value="South Ossetia"> South Ossetia </option>
                    <option value="Spain"> Spain </option>
                    <option value="Sri Lanka"> Sri Lanka </option>
                    <option value="Sudan"> Sudan </option>
                    <option value="Suriname"> Suriname </option>
                    <option value="Svalbard"> Svalbard </option>
                    <option value="Swaziland"> Swaziland </option>
                    <option value="Sweden"> Sweden </option>
                    <option value="Switzerland"> Switzerland </option>
                    <option value="Syria"> Syria </option>
                    <option value="Taiwan"> Taiwan </option>
                    <option value="Tajikistan"> Tajikistan </option>
                    <option value="Tanzania"> Tanzania </option>
                    <option value="Thailand"> Thailand </option>
                    <option value="Timor-Leste"> Timor-Leste </option>
                    <option value="Togo"> Togo </option>
                    <option value="Tokelau"> Tokelau </option>
                    <option value="Tonga"> Tonga </option>
                    <option value="Transnistria Pridnestrovie"> Transnistria Pridnestrovie </option>
                    <option value="Trinidad and Tobago"> Trinidad and Tobago </option>
                    <option value="Tristan da Cunha"> Tristan da Cunha </option>
                    <option value="Tunisia"> Tunisia </option>
                    <option value="Turkey"> Turkey </option>
                    <option value="Turkmenistan"> Turkmenistan </option>
                    <option value="Turks and Caicos Islands"> Turks and Caicos Islands </option>
                    <option value="Tuvalu"> Tuvalu </option>
                    <option value="Uganda"> Uganda </option>
                    <option value="Ukraine"> Ukraine </option>
                    <option value="United Arab Emirates"> United Arab Emirates </option>
                    <option value="United Kingdom"> United Kingdom </option>
                    <option value="Uruguay"> Uruguay </option>
                    <option value="Uzbekistan"> Uzbekistan </option>
                    <option value="Vanuatu"> Vanuatu </option>
                    <option value="Vatican City"> Vatican City </option>
                    <option value="Venezuela"> Venezuela </option>
                    <option value="Vietnam"> Vietnam </option>
                    <option value="British Virgin Islands"> British Virgin Islands </option>
                    <option value="US Virgin Islands"> US Virgin Islands </option>
                    <option value="Wallis and Futuna"> Wallis and Futuna </option>
                    <option value="Western Sahara"> Western Sahara </option>
                    <option value="Yemen"> Yemen </option>
                    <option value="Zambia"> Zambia </option>
                    <option value="Zimbabwe"> Zimbabwe </option>
                    <option value="other"> Other </option>
                  </select>
                  <label class="form-sub-label" for="input_8_country" id="sublabel_8_country"> Country </label></span>              </td>
            </tr>
          </table>
        </div>
      </li>
      <li class="form-line" id="id_2">
        <div id="cid_2" class="form-input-wide">
          <div style="margin-left:156px" class="form-buttons-wrapper">
            <button id="input_2" type="submit" class="form-submit-button form-submit-button-3d_round_yellow">
              Submit            </button>
             
            <button id="input_reset_2" type="reset" class="form-submit-reset form-submit-button-3d_round_yellow">
              Clear Form            </button>
          </div>
        </div>
      </li>
      <li style="display:none">
        Should be Empty:
        <input type="text" name="website" value="" />
      </li>
    </ul>
  </div>
</form>

<%


End If




objRec7.Close()   
Conn7.Close()   
Set objRec7 = Nothing  
Set Conn7 = Nothing  




%>   
  </span></span>  </div>
</body>   
</html>  
