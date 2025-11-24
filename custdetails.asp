
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
.style21 {
	font-family: "Calisto MT";
	color: #990033;
}
.style28 {color: #FF0000; font-weight: bold; }
.style31 {
	color: #000000;
	font-weight: bold;
}
.style33 {font-size: 24px; font-family: Impact;}
.style34 {color: #660000}
-->
</style></head>   
<body>   
<div align="center" class="style14 style21">JOB BOARD</div>
<div align="center">
  <p>
    <%   
Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic, strtyp
dim strPASS, strpass2, strpassM, rec, strpwd, struname



rec=request.querystring("reco")

strpwd=request.querystring("upwd")
struname=request.querystring("uname")










%>


<%

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
    
    
    
 


  


 <table width="980" border="1" bgcolor="#666666">
    <tr>
      
               <td align="center" width="82"><strong>Job #</a></strong></td>
              
    <td align="center" width="100"><strong>Customer</strong></td>
  <td align="center" width="348"><strong>Job Name</strong></td>

   
    <td align="center" width="50"><strong>Size #</strong></td>
    <td align="center" width="50"><div align="center"><strong>Metal </strong></div></td>
     <td align="center" width="60"><strong>Order Date</strong></td>

   
     <td align="center" width="60"><strong>Due Date</strong></td>
    
    
   <td align="center" width="40"><strong>Completed</strong></td>

     
  </tr>
      
    
     
       
      
          <tr width="980">
             <input type="hidden" name="Jnum" value="<%=objRec.Fields("RecNo").Value%>">
             
            <td width="82" align="center"><div align="center"><strong><a name="<%=objRec.Fields("RecNo").Value%>"><%=objRec.Fields("RecNo").Value%></strong></div></a></td>

	    <td width="100" align="center"><div align="center" class="style15"><strong><a href="details.asp?pass=<%=strpass2%>&reco=<%=objRec.Fields("RecNo").Value%>" target="_self" class=""> <%=objRec.Fields("CNam").Value%></a></strong></div></td>


<td align="center" width="391"><div align="center" class="style34"><span class="style33"><%=objRec.Fields("JNam").Value%></span></div></td>
       <td align="center" width="50"><div align="center"><strong><%=objRec.Fields("FSize").Value%></strong></div></td>
    <td align="center" width="50"><div align="center"><strong><%=objRec.Fields("MType").Value%></strong></div></td>
		<td align="center" width="60"><div align="center"><strong><%=objRec.Fields("ODate").Value%></strong></div></td>
		
    
    <td align="center" width="60"><div align="center"><strong><%=objRec.Fields("DDate").Value%></strong></div></td>
    
    

  
 
      
      

		<td align="center" width="40"><div align="center" class="style28"><%=objRec.Fields("status").Value%></div></td> 
    

 
      </tr>

    <tr>
	  <%Dim strChange
	  
	  strChange=objRec.Fields("Change").Value


	  %>
	  <td align="center" colspan="2"><div align="right"><strong>Instructions:</strong></div></td>
	  <td align="center" colspan="7"><div align="center" class="style16"><span class="style31"><%=objRec.Fields("Instructions").Value%></span></div></td>
    </tr>
  <tr>
	  
	  <td align="center" colspan="2"><div align="right"><strong>Stone Info:</strong></div></td>
	  <td align="center" colspan="7"><div align="center" class="style15"><span class="style31"><%=objRec.Fields("SInfo").Value%></span></div></td>
    </tr>
       

	  <tr>
	  <td colspan="2">
	  
	  <form action="login.asp#<%=objRec.Fields("RecNo").Value%>" target="_self" method="post">
	  <input type="hidden" name="upwd" value="<%=strPwd%>">
	  <input type="hidden" name="uname" value="<%=strUname%>">
	  <input type="hidden" name="button" value="Log in">
	  <input type="submit" value="GO BACK">




	  </form>
	  
	  
	  
	  
	  </td>
<% Dim Conn6,strSQL6,objRec6, curjob, curfile
Set Conn6 = Server.Createobject("ADODB.Connection")   

curjob=objRec.Fields("RecNo").Value
Conn6.Open connectstr


strSQL6 = "SELECT filenam from Uploads where job='"&curjob&"'"

Set objRec6 = Server.CreateObject("ADODB.Recordset")   
objRec6.Open strSQL6, Conn6, 1,3

If objRec6.EOF Then
%>
<TD colspan="6"><%Response.write ("No images uploaded....")%></td>

<%
Else

%><td></td><%

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




%>   
  </span></span>
  </div>
</body>   
</html>  
