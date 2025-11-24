
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
.style8 {
	color: #FFFFFF;
	font-family: Impact;
}
.style14 {
	font-family: Impact;
	font-size: 60px;
}
.style16 {color: #FF0000}
.style19 {font-size: 18px}
.style23 {color: #660000}
.style24 {color: #000033}
.style25 {color: #660033}
-->
</style></head>   
<body>   

<div align="center">
  <p>
    <%   
Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic, strtyp
dim strPASS, strpass2

strpass=Request.Querystring ("pass")
strpass2="99999"

if strpass <> strpass2 then

%>
Your log in information was incorrect.  Are you sure you belong here?

<%
else
db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "InvDb"
connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword




Dim Conn,strSQL,objRec, strmodel, strCAS
Set Conn = Server.Createobject("ADODB.Connection")   
Conn.Open connectstr

strCas=request.form("cas")

If strCAS="ALL" OR strCAS="" Then

strSQL = "SELECT * FROM MEMDB ORDER BY Odate, Cnam, Status DESC"


Else

strSQL="SELECT * FROM MEMDB WHERE CNAM='"&strCAS&"' ORDER BY Odate, Cnam, Status DESC"

End if
Set objRec = Server.CreateObject("ADODB.Recordset")   
objRec.Open strSQL, Conn, 1,3   
  
If objRec.EOF Then  
Response.write (" No records found.") %>
<p class="style14 style16"><span class="style23">ALL JOBS</span><a href="jobupdater.asp?pass=<%=strpass2%>" target="_self" class="style19 style24">(OPEN)  </a><a href="jobupdater2.asp?pass=<%=strpass2%>" target="_self" class="style19 style25">  (CLOSED)</a></p>
<%
Else  
  
Dim PageLen,PageNo,TotalRecord,TotalPage,No,intID   
PageLen = 50
PageNo = Request.QueryString("Page") 

if PageNo = "" Then PageNo = 1   
TotalRecord = objRec.RecordCount   
objRec.PageSize = PageLen   
TotalPage = objRec.PageCount   
objRec.AbsolutePage = PageNo   

No=1 
Dim I
i=0
%>
</p>


  <p class="style14 style16"><span class="style23">ALL JOBS</span><a href="jobupdater.asp?pass=<%=strpass2%>" target="_self" class="style19 style24">(OPEN)  </a><a href="jobupdater2.asp?pass=<%=strpass2%>" target="_self" class="style19 style25">  (CLOSED)</a></p>
   
	
	
  <span class="style8">Total : <%=TotalRecord%> Jobs.  Page <%=PageNo%> of <%=TotalPage%>   <br>
  <% IF Cint(PageNo) > 1 then %>   
  <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=1"><< First</a>   <br><br>
  <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=PageNo-1%>">< Back</a>   
  <% End IF%>   
  <% IF Cint(PageNo) < TotalPage Then %>   
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=PageNo+1%>">Next ></a>  </span>
  <form action="jobupdater3.asp?Pass=88888" method="POST" class="style8">

<input type="hidden" name="typ" value="<%=strTyp%>" />

<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=TotalPage%>&strtyp=<%=strtyp%>">Last >></a>   
</form>
  
  <span class="style8">
  <% End IF%>   
  <br>   
  Go to   
  <% For intID = 1 To TotalPage%>   
  <% if intID = Cint(PageNo) Then%>   
  <b><%=intID%></b>   
  <%Else%>   
  <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=intID%>"><%=intID%></a>   
  <%  End if%>   
  <%Next%>   
  <%end If  %>

<%
Dim Conn2,strSQL2,objRec2
Set Conn2 = Server.Createobject("ADODB.Connection")   
Conn2.Open connectstr


strSQL2 = "SELECT DISTINCT CNAM FROM MEMDB ORDER BY CNam"

Set objRec2 = Server.CreateObject("ADODB.Recordset")   
objRec2.Open strSQL2, Conn2, 1,3



  
 

%>


<table>
<form action="jobupdater3.asp?pass=<%=strpass%>" method="post" target="_self">

<tr>
<td>SEARCH BY CUSTOMER:</td>
<td><select name="CAS" selected value="<%Response.write(strCAS)%>">

                 <option value="ALL"> Show All </option>
				 <%Do While Not objRec2.EOF%>
				 
				 
				 <option value="<%=objRec2.Fields("CNam").Value%>"> <%=objRec2.Fields("CNam").Value%> </option>
			
          



<%
objRec2.MoveNext   
Loop 




%>

</select>     </td>
<td><input type="submit" value="GO"></td>
</tr>
 
</form>
</table>
<%



objRec2.Close()   
Conn2.Close()   
Set objRec2 = Nothing  
Set Conn2 = Nothing 
%>



  <table width="980" border="1">
    <tr>      
               <td align="center" width="40"><strong>Job #</strong></td>
              
    <td align="center" width="100"><strong>Customer</strong></td>
 <td align="center" width="75"><strong>Job Name</strong></td>
    <td align="center" width="60"><strong>Order Date</strong></td>
     <td align="center" width="60"><strong>Due Date</strong></td>
   
    <td align="center" width="50"><strong>Finger Size</strong></td>
    <td align="center" width="50"><strong>Metal Type</strong></td>
    <td align="center" width="100"><strong>Stone Information</strong></td>
  
    <td align="center" width="40"><strong>Completed</strong></td>
   
    
  </tr>
      
<%
Do While Not objRec.EOF and No <= PageLen   

i=i+1



%>   

     
        <form action="jobupdate2.asp?pass=<%=strpass2%>" method="post" target="_self">
      
          <tr width="980">
             <input type="hidden" name="Jnum" value="<%=objRec.Fields("RecNo").Value%>">
             
            <td width="40" align="center"><div align="center"><%=objRec.Fields("RecNo").Value%></div></td>
          
    <td width="100" align="center"><div align="center"><%=objRec.Fields("CNam").Value%></div></td>
        <td align="center" width="75"><div align="center"><%=objRec.Fields("JNam").Value%></div></td>
    <td align="center" width="60"><div align="center"><%=objRec.Fields("ODate").Value%></div></td>
    <td align="center" width="60"><div align="center"><%=objRec.Fields("DDate").Value%></div></td>
    
    <td align="center" width="50"><div align="center"><%=objRec.Fields("FSize").Value%></div></td>
    <td align="center" width="50"><div align="center"><%=objRec.Fields("MType").Value%></div></td>
    <td align="center" width="100"><div align="center"><%=objRec.Fields("Sinfo").Value%></div></td>
   
    <td align="center" width="40"><div align="center"><%=objRec.Fields("Status").Value%></div></td>
    
            
		

    
      </tr>
     
	
        </form>
      
  </div>
    <%No = No + 1   
objRec.MoveNext   
Loop  



%>  

</table>




</tr> 
  </table>
<%

objRec.Close()   
Conn.Close()   
Set objRec = Nothing  
Set Conn = Nothing  
end if

%>   


</body>   
</html>  
