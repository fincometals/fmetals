
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
dim strPASS, strpass2, strpassM, rec, closer

closer=request.querystring("clos")

rec=request.querystring("reco")


strpass=Request.Querystring ("pass")

strpass2="88888"





db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "InvDb"
connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword




Dim Conn,strSQL,objRec, strmodel
Set Conn = Server.Createobject("ADODB.Connection")   
Conn.Open connectstr


strSQL = "SELECT * FROM MEMDB Where RECNO="& rec

Set objRec = Server.CreateObject("ADODB.Recordset")   
objRec.Open strSQL, Conn, 1,3   

%>   
    
    
    
 


  


 <table width="980" border="1">
    <tr>
      
               <td align="center" width="40"><strong>Job #</a></strong></td>
              
    <td align="center" width="100"><strong>Customer</strong></td>
 <td align="center" width="60"><strong>Order Date</strong></td>
 <td align="center" width="75"><strong>Job Name</strong></td>
   
     <td align="center" width="60"><strong>Due Date</strong></td>
   
    <td align="center" width="50"><strong>Size #</strong></td>
    <td align="center" width="50"><div align="center"><strong>Metal </strong></div></td>
    
    
    
   <td align="center" width="40"><strong>CAD</strong></td>
<td align="center" width="40"><strong>CAM</strong></td>
      <td></td>
  </tr>
      
    
     
       
      
          <tr width="980">
             <input type="hidden" name="Jnum" value="<%=objRec.Fields("RecNo").Value%>">
             
            <td width="40" align="center"><div align="center"><strong><a name="<%=objRec.Fields("RecNo").Value%>"><%=objRec.Fields("RecNo").Value%></strong></div></a></td>

	    <td width="100" align="center"><div align="center" class="style15"><strong><a href="details.asp?pass=<%=strpass2%>&reco=<%=objRec.Fields("RecNo").Value%>" target="_self" class=""> <%=objRec.Fields("CNam").Value%></a></strong></div></td>



        <td align="center" width="60"><div align="center"><strong><%=objRec.Fields("ODate").Value%></strong></div></td>
		<td align="center" width="75"><div align="center" class="style23"><strong><%=objRec.Fields("JNam").Value%></strong></div></td>
    
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
	  
	  <td align="center" width="50" colspan="2"><div align="right"><strong>Instructions:</strong></div></td>
	  <td align="center" colspan="8" width="900"><div align="center" class="style16"><strong><%=objRec.Fields("Comments").Value%></strong></div></td>
	  </tr>
  <tr>
	  
	  <td align="center" width="50" colspan="2"><div align="right"><strong>Stone Info:</strong></div></td>
	  <td align="center" colspan="8" width="900"><div align="center" class="style15"><strong><%=objRec.Fields("SInfo").Value%></strong></div></td>
	  </tr>

	  <tr>
	  <td><a href="login.asp?cust=<%=objRec.Fields("CNam").Value%>" class="style19 16, style"><strong>GO BACK</strong></A></td>
	
</tr>

	

</table>
  



<%




objRec.Close()   
Conn.Close()   
Set objRec = Nothing  
Set Conn = Nothing  




%>   
  </span></span>
  </div>
</body>   
</html>  
