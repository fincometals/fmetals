
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
dim strPASS, strpass2, strR1, strR2, strButton

strR1=request.form("Reno")
strR2=request.form("Reco")
strButton=request.form("button")

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


if strbutton <> "" then 



Dim Conn45,strSQL45,objRec45
Set Conn45 = Server.Createobject("ADODB.Connection")   
Conn45.Open connectstr

if strButton = "Read" then

strSQL45="UPDATE Guest SET Red='yes' WHERE Recno='"&strR1&"'"

else
strSQL45="UPDATE GuestUploads SET Red='yes' WHERE Rec='"&strR2&"'"
end if

conn45.execute strSQL45

conn45.close
set conn45=nothing


end if







Dim Conn,strSQL,objRec, strmodel, strCAS
Set Conn = Server.Createobject("ADODB.Connection")   
Conn.Open connectstr





strSQL = "SELECT * FROM Guest Where Red='no' ORDER BY Utime DESC"


Set objRec = Server.CreateObject("ADODB.Recordset")   
objRec.Open strSQL, Conn, 1,3   
  
If objRec.EOF Then  
Response.write (" No new requests.") %>
<p class="style14 style16"><span class="style24">Inquiries</span><a href="jobupdater.asp?pass=<%=strpass2%>" target="_self" class="style19 style24">(OPEN)  </a><a href="jobupdater2.asp?pass=<%=strpass2%>" target="_self" class="style19 style25">  (CLOSED)</a><a href="jobupdater3.asp?pass=<%=strpass2%>" target="_self" class="style19 style24">(ALL JOBS)  </a></p>
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


  <p class="style14 style16"><span class="style24">Inquiries</span><a href="jobupdater.asp?pass=<%=strpass2%>" target="_self" class="style19 style24">(OPEN)  </a><a href="jobupdater2.asp?pass=<%=strpass2%>" target="_self" class="style19 style25">  (CLOSED)</a></p>
   
	
	
  <span class="style8">Total : <%=TotalRecord%> Request(s).  Page <%=PageNo%> of <%=TotalPage%>   <br>
  <% IF Cint(PageNo) > 1 then %>   
  <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=1"><< First</a>   <br><br>
  <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=PageNo-1%>">< Back</a>   
  <% End IF%>   
  <% IF Cint(PageNo) < TotalPage Then %>   
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=PageNo+1%>">Next ></a>  </span>


<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?pass=<%=strpass2%>&Page=<%=TotalPage%>&strtyp=<%=strtyp%>">Last >></a>   

  
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





  <table width="980" border="1" align="center">
    <tr>      
               <td align="center" width="40"><strong>REC#</strong></td>
              
    <td align="center" width="100"><strong>Customer</strong></td>
 <td align="center" width="75"><strong>Company</strong></td>
    <td align="center" width="60"><strong>Phone</strong></td>
     <td align="center" width="60"><strong>Email</strong></td>
   
    <td align="center" width="50"><strong>Requested</strong></td>

    
  </tr>
 
      
<%
Do While Not objRec.EOF and No <= PageLen   

i=i+1


dim currec

currec=objRec.Fields("RecNo").Value
%>   

     
      
      
          <tr width="980">
        
             
            <td width="40" align="center"><div align="center"><form action="./inquiry.asp?pass=<%=strpass%>" target="_self" method="post"><input type="submit" name="button" value="Read"><input type="hidden" name="RENO" value="<%=currec%>"></form><%=objRec.Fields("RecNo").Value%></div></td>
          
    <td width="100" align="center"><div align="center"><%=objRec.Fields("FNam").Value%> <%=objRec.Fields("LNam").Value%></div></td>
        <td align="center" width="75"><div align="center"><%=objRec.Fields("CoNam").Value%></div></td>
    <td align="center" width="60"><div align="center"><%=objRec.Fields("Acd").Value%>-<%=objRec.Fields("Pnum").Value%></div></td>
    <td align="center" width="60"><div align="center"><a href="mailto:<%=objRec.Fields("Eml").Value%>"><%=objRec.Fields("Eml").Value%></a></div></td>
    
    <td align="center" width="50"><div align="center"><%=objRec.Fields("Utime").Value%></div></td>
  
		

    
      </tr>
      <tr><td colspan="6"><%=objRec.Fields("Comments").Value%></td></tr>
     
	
      
      
 
    <%No = No + 1   
objRec.MoveNext   
Loop  



%>  

</table>





<%

objRec.Close()   
Conn.Close()   
Set objRec = Nothing  
Set Conn = Nothing  
end if


Dim Conn2,strSQL2,objRec2
Set Conn2 = Server.Createobject("ADODB.Connection")   
Conn2.Open connectstr






strSQL2 = "SELECT * FROM GuestUploads Where Red='no' ORDER BY Utime DESC"


Set objRec2 = Server.CreateObject("ADODB.Recordset")   
objRec2.Open strSQL2, Conn2, 1,3   
%>
 <p class="style14 style16"><span class="style24">Uploads</span></p>
<table border="1" width="980" align="center">
<tr>
<td><strong>Filename</strong></td>
<td><strong>Requested on</strong></td>
<td><strong>File Preview</strong></td>
<%DO while not objRec2.EOF%>
<tr><td><a href="./guestuploads/<%=objRec2.Fields("Filenam").Value%>"><%=objRec2.Fields("Filenam").Value%></a></td>

<td><%=objRec2.Fields("Utime").Value%></td>
<td><a href="./guestuploads/<%=objRec2.Fields("Filenam").Value%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','./guestuploads/<%=objRec2.Fields("Filenam").Value%>',1)"><img name="Image2" border="0" height="60" width="80" src="./guestuploads/<%=objRec2.Fields("Filenam").Value%>"></a><form action="./inquiry.asp?pass=<%=strpass%>" target="_self" method="post"><input type="hidden" name="RECO" value="<%=objRec2.Fields("Rec").Value%>"><input type="submit" name="button" value="Mark Read"> </form></td>


</tr>

<%objrec2.movenext
loop%>

</table>
<p>
  <%objRec2.Close()   
Conn2.Close()   
Set objRec2 = Nothing  
Set Conn2 = Nothing  



Dim Conn9,strSQL9,objRec9
Set Conn9 = Server.Createobject("ADODB.Connection")   
Conn9.Open connectstr





strSQL9 = "SELECT * FROM Guest Where Red='yes' ORDER BY Utime DESC"


Set objRec9 = Server.CreateObject("ADODB.Recordset")   
objRec9.Open strSQL9, Conn9, 1,3   %>
  
</p>

 <p class="style14 style16"><span class="style23">Past Inquiries</span></p>
<table width="980" border="1" align="center">
   <tr>      
               <td align="center" width="40"><strong>REC#</strong></td>
              
    <td align="center" width="100"><strong>Customer</strong></td>
 <td align="center" width="75"><strong>Company</strong></td>
    <td align="center" width="60"><strong>Phone</strong></td>
     <td align="center" width="60"><strong>Email</strong></td>
   
    <td align="center" width="50"><strong>Requested</strong></td>

    
  </tr>
 
      
<%
Do While Not objRec9.EOF







%>   

     
      
      
          <tr width="980">
        
             
            <td width="40" align="center"><div align="center"><%=objRec9.Fields("RecNo").Value%></div></td>
          
    <td width="100" align="center"><div align="center"><%=objRec9.Fields("FNam").Value%> <%=objRec9.Fields("LNam").Value%></div></td>
        <td align="center" width="75"><div align="center"><%=objRec9.Fields("CoNam").Value%></div></td>
    <td align="center" width="60"><div align="center"><%=objRec9.Fields("Acd").Value%>-<%=objRec9.Fields("Pnum").Value%></div></td>
    <td align="center" width="60"><div align="center"><a href="mailto:<%=objRec9.Fields("Eml").Value%>"><%=objRec9.Fields("Eml").Value%></a></div></td>
    
    <td align="center" width="50"><div align="center"><%=objRec9.Fields("Utime").Value%></div></td>
  
		

    
      </tr>
      <tr><td colspan="6"><%=objRec9.Fields("Comments").Value%></td></tr>
     
	
      
      
 
  <% 
objRec9.MoveNext   
Loop  



%>  

</table>





<%

objRec9.Close()   
Conn9.Close()   
Set objRec9 = Nothing  
Set Conn9 = Nothing  



Dim Conn22,strSQL22,objRec22
Set Conn22 = Server.Createobject("ADODB.Connection")   
Conn22.Open connectstr






strSQL22 = "SELECT * FROM GuestUploads Where Red='yes' ORDER BY Utime DESC"


Set objRec22 = Server.CreateObject("ADODB.Recordset")   
objRec22.Open strSQL22, Conn22, 1,3   
%>
 <p class="style14 style16"><span class="style23">Past Uploads</span></p>
<table border="1" width="980" align="center">
<tr>
<td><strong>Filename</strong></td>
<td><strong>Requested on</strong></td>
<td><strong>File Preview</strong></td>
<%DO while not objRec22.EOF%>
<tr><td><a href="./guestuploads/<%=objRec22.Fields("Filenam").Value%>"><%=objRec22.Fields("Filenam").Value%></a></td>

<td><%=objRec22.Fields("Utime").Value%></td>
<td><a href="./guestuploads/<%=objRec22.Fields("Filenam").Value%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','./guestuploads/<%=objRec22.Fields("Filenam").Value%>',1)"><img name="Image2" border="0" height="60" width="80" src="./guestuploads/<%=objRec22.Fields("Filenam").Value%>"></a></td>


</tr>

<%objrec22.movenext
loop%>

</table>
<p>
  <%objRec22.Close()   
Conn22.Close()   
Set objRec22 = Nothing  
Set Conn22 = Nothing  
%>   


</body>   
</html>  
