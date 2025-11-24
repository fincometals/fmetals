
 <% Option Explicit 
On Error Resume next
%>   
<html>   
<head>   
<title>FINCO ADMIN PANEL</title>   
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
body {
	background-image: url(http://www.fincometals.com/watches/gearb.jpg);
}
.style14 {
	font-family: Impact;
	font-size: 60px;
}
.style15 {color: #0000FF}
.style16 {color: #FF0000}
.style19 {font-size: 18px}
.style21 {
	font-family: "Calisto MT";
	color: #990033;
}
.style23 {color: #009900}
.style28 {color: #FF0000; font-weight: bold; }
.style29 {color: #0000FF; font-weight: bold; }
.style31 {
	color: #000000;
	font-weight: bold;
}
.style33 {color: #660000}
-->
</style></head>   
<body>   
<div align="center" class="style14 style21">FINCO JOB BOARD</div>

<div align="center">
  <p>
    <%   
Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic, strtyp
dim strPASS, strpass2, strpassM, rec, closer, strRead, strReadChk, strRes

strRes=request.form("resp")


strReadChk="read"


strRead=request.querystring("read")

closer=request.querystring("clos")

rec=request.querystring("reco")


strpass=Request.Querystring ("pass")

strpass2="99999"




if strpass <> strpass2 AND strpassM <> strpass2 then

%>
Your log in information is incorrect.  Are you sure you belong here?

<%
else
db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "InvDb"
connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword


If strRead=strReadChk then

Dim Conn8,strSQL8,objRec8, strNInst, strOInst, stroldin
Set Conn8 = Server.Createobject("ADODB.Connection")   

stroldin=request.form("oldin")

Conn8.Open connectstr

strNInst=strOldin&"<BR><span class=style66 style44 style48>Response:"&strRes&"</span>"

strNInst = Replace(strNInst, "'", "''")


strSQL8 = "UPDATE MemDB SET Instructions='"&strNInst&"',Change='no', Mread='no' WHERE RECNO='"&rec&"'" 






Conn8.Execute strSQL8
Conn8.Close
Set Conn8 = nothing
%>
<script language="javascript">
<!--
alert("Information saved...");

//-->
</script> 


<%Else%><%
End If


Dim Conn,strSQL,objRec, strmodel
Set Conn = Server.Createobject("ADODB.Connection")   
Conn.Open connectstr


strSQL = "SELECT * FROM MEMDB Where RECNO="& rec

Set objRec = Server.CreateObject("ADODB.Recordset")   
objRec.Open strSQL, Conn, 1,3   

%>   
    
    
    
 


  


 <table width="980" border="1" bgcolor="#666666">
    <tr>
      
               <td align="center" width="82"><strong>Job #</a></strong></td>
              
    <td align="center" width="100"><strong>Customer</strong></td>
 <td align="center" width="60"><strong>Order Date</strong></td>
 <td align="center" width="348"><strong>Job Name</strong></td>
   
     <td align="center" width="60"><strong>Due Date</strong></td>
   
    <td align="center" width="50"><strong>Size #</strong></td>
    <td align="center" width="50"><div align="center"><strong>Metal </strong></div></td>
    
    
    
   <td align="center" width="40"><strong>CAD</strong></td>
<td align="center" width="40"><strong>CAM</strong></td>
      <td></td>
  </tr>
      
    
     
       
      
          <tr width="980">
             <input type="hidden" name="Jnum" value="<%=objRec.Fields("RecNo").Value%>">
             
            <td width="82" align="center"><div align="center"><a name="<%=objRec.Fields("RecNo").Value%>"><%=objRec.Fields("RecNo").Value%></a></div></td>

	    <td width="100" align="center"><div align="center" class="style15"><strong><a href="details.asp?pass=<%=strpass2%>&reco=<%=objRec.Fields("RecNo").Value%>" target="_self" class=""> <%=objRec.Fields("CNam").Value%></a></strong></div></td>



        <td align="center" width="60"><div align="center"><strong><%=objRec.Fields("ODate").Value%></strong></div></td>
		<td align="center" width="348"><div align="center" class="style23"><strong><%=objRec.Fields("JNam").Value%></strong></div></td>
    
    <td align="center" width="60"><div align="center"><strong><%=objRec.Fields("DDate").Value%></strong></div></td>
    
    <td align="center" width="50"><div align="center"><strong><%=objRec.Fields("FSize").Value%></strong></div></td>
    <td align="center" width="50"><div align="center"><strong><%=objRec.Fields("MType").Value%></strong></div></td>




  <%dim strChecker1, strChecker2, strChecker3
	
	strChecker1=objRec.Fields("Status").Value
	strChecker2=objRec.Fields("CC").Value
	strChecker3=objRec.Fields("CAM").value
	
	
	
    
    If strChecker2="0PEN" then%>
	    <td align="center" width="40"><div align="center" class="style29"><%=objRec.Fields("CC").Value%></div></td>
        
        <%else %>
      <td align="center" width="40"><div align="center" class="style28"><%=objRec.Fields("CC").Value%></div></td>
      <%end if
      
      
       If strChecker3="0PEN" then%>
<td align="center" width="40"><div align="center" class="style29"><%=objRec.Fields("CAM").Value%></div></td>        
<% else %>
		<td align="center" width="40"><div align="center" class="style28"><%=objRec.Fields("CAM").Value%></div></td> 
        <% end if%>       
      </tr>

    <tr>
	  <%Dim strChange
	  
	  strChange=objRec.Fields("Change").Value


	  %>
	  <td align="center" colspan="2"><div align="right"><strong>Instructions:
	  
	  
	  <br>
      
      <form action="details.asp?reco=<%=rec%>&pass=<%=strpass2%>&read=read" target="_self" method="post" id="usrfrm">
      <input type="hidden" name="oldin" value="<%=objRec.Fields("Instructions").Value%>">
        <strong>
        <textarea cols="30" rows="2" name=resp form="usrfrm"></textarea>
        </strong>
        <input type="submit" value="Mark Read">
</form>

	  
	  <%
	  
	  dim strReed
	  
	  strReed=objRec.Fields("MRead").Value
	  
	  if strReed="no" then%>
      <span class="style33">Reply not viewed</span>
      <%else%>
	  
	  
	  <%
	  end if
	  
	  
	  %></strong></div></td>
	  <td align="center" colspan="8"><div align="center" class="style16"><span class="style31"><%=objRec.Fields("Instructions").Value%></span></div></td>
    </tr>
  <tr>
	  
	  <td align="center" colspan="2"><div align="right"><strong>Stone Info:</strong></div></td>
	  <td align="center" colspan="8"><div align="center" class="style15"><span class="style31"><%=objRec.Fields("SInfo").Value%></span></div></td>
    </tr>
        <tr>
	  
	  <td align="center" colspan="2"><div align="right"><strong>Notes:</strong></div></td>
	  <td align="center" colspan="8"><div align="center" class="style15"><span class="style31"><%=objRec.Fields("Comments").Value%></span></div></td>
    </tr>
      <tr>
	  
	  <td align="center" colspan="2"><div align="right"><strong>Elve's Notes:</strong></div></td>
	  <td align="center" colspan="8"><div align="center" class="style15"><span class="style31"><%=objRec.Fields("TimInst").Value%></span></div></td>
    </tr>
         <tr>
	  
	  <td align="center" colspan="2"><div align="right"><strong>Files:</strong></div></td>
	  <td align="center" colspan="8"> <%Dim Conn18,strSQL18,objRec18, curjob18, curfile18
Set Conn18 = Server.Createobject("ADODB.Connection")   
Conn18.Open connectstr

curjob18=objRec.Fields("RecNo").Value


strSQL18 = "SELECT filenam from files where RecNo='"&curjob18&"'"

Set objRec18 = Server.CreateObject("ADODB.Recordset")   
objRec18.Open strSQL18, Conn18, 1,3

Do While Not objRec18.EOF
dim strasdf
curfile18=objRec18.Fields("Filenam").Value

if curfile18="" then

else

strasdf="ftp://Ocean350:Mon1tor!@fincometals.com/files/"&curfile18

%>

<a href="<%=strasdf%>" target="_blank"><%=curfile18%></a>
<%end if
objRec18.MoveNext   
Loop  %>
      
      
      
      
      </td>
    </tr>

	  <tr>
	  <td colspan="2"><A Href="uploadtester2.asp?recno=<%=objRec.Fields("RecNo").Value%>&jnam=<%=objRec.Fields("JNam").Value%>" target="_self" class="style42 style35">Upload Photos</a>
	  
	  
	  
	  <a href="jobupdater.asp?pass=<%=strpass%>&cur=<%=rec%>#<%=rec%>" class="style19 16, style"><strong>GO BACK</strong></A></td>
<% Dim Conn6,strSQL6,objRec6, curjob, curfile
Set Conn6 = Server.Createobject("ADODB.Connection")   

curjob=objRec.Fields("RecNo").Value
Conn6.Open connectstr


strSQL6 = "SELECT filenam from Uploads where job='"&curjob&"'"

Set objRec6 = Server.CreateObject("ADODB.Recordset")   
objRec6.Open strSQL6, Conn6, 1,3

If objRec6.EOF Then
%>
<TD><%Response.write ("No images uploaded....")%></td>

<%
Else



Do While Not objRec6.EOF

curfile=objRec6.Fields("filenam").Value
%>


<TD align="center">

<a href="./uploads/<%=curfile%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','./uploads/<%=curfile%>',1)"><img name="Image2" border="0" height="60" width="80" src="./uploads/<%=curfile%>"></a></td>

<%objRec6.MoveNext   
Loop 
objRec6.Close()   
Conn6.Close()   
Set objRec6 = Nothing  
Set Conn6 = Nothing 

End if%>
</tr>
</table>
  



<%




objRec.Close()   
Conn.Close()   
Set objRec = Nothing  
Set Conn = Nothing  

End If


%>   
  </span></span>
</div>
</body>   
</html>  
