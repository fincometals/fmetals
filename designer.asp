
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
.style16 {
	color: #000033
}
.style19 {color: #000066}
.style21 {
	color: #006633
}
.style23 {font-family: Impact}
.style24 {font-family: Impact; font-size: 60px; color: #660000; }
.style27 {color: #000033; font-family: Impact; }
.style28 {color: #003300}
.style29 {color: #660000}
.style30 {color: #000000}
.style32 {font-family: "Calisto MT"}
.style33 {
	color: #000033;
	font-family: "Calisto MT";
	font-weight: bold;
}
.style35 {
	color: #000000;
	font-family: "Calisto MT";
	font-weight: bold;
}
-->
</style>
<script type="text/javascript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>   
<body onLoad="MM_preloadImages('./uploads/<%=curfile%>','./uploads/<%=curfile2%>')">   
<div align="center" class="style14">JOB CENTRAL</div>
<div align="center">
  <p>
    <%   
Dim connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic, strtyp
dim strPASS, strpass2, strpass4

strpass4=request.QueryString("pass")

if strpass4<>"" then 

strpass=strpass4
else



strpass=Request.querystring ("pass")
end if


strpass2="TIMDS"
strpass=UCase(strpass)

if strpass <> strpass2 then%>
Your log in information was incorrect.  Are you sure you belong here?<%
else
db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "InvDb"
connectstr = "Driver={SQL Server};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword


Dim Conn2,strSQL2,objRec2, strbutton, strRE, strTod, strbut
Set Conn2 = Server.Createobject("ADODB.Connection")   
Conn2.Open connectstr

strRE=request.QueryString("rec")
strbutton=request.querystring("button")

strbut=request.form("but")


Dim strR, strTimInstN, strTimInstOld, strnews, strN

strN=now()

strTimInstOld=request.form("TI")

strR=request.form("R")

strnews=request.form("news")

strTimInstN=strTimInstOld&"<BR><span class=style16><strong>TIM:"&strnews&"*"&strN&"</strong></span>"

strTimInstN=Replace(strTimInstN, "'", "''")

strTod=now()

if strbutton="Complete" then

strSQL2 = "UPDATE MemDB SET TimDone='yes' WHERE RecNo='"&strRE&"'"%>    


<script language="javascript">
<!--
alert("Updated.");

//-->
    </script> <%
end if

if strbutton="ReOpen" then

strSQL2 = "UPDATE MemDB SET TimDone='no' WHERE RecNo='"&strRE&"'"%>    


<script language="javascript">
<!--
alert("Updated.");

//-->
    </script> <%
end If

If strbut="send" Then



strSQL2="UPDATE MemDB SET TimInst='"&strTimInstN&"', TimMes='yes' WHERE RecNo='"&strR&"'"%>
<script language="javascript">
<!--
alert("Updated.");

//-->
    </script> <%
end If


Conn2.Execute strSQL2

Conn2.Close()   
 
Set Conn2 = Nothing 


Dim Conn,strSQL,objRec, strmodel, strCAS
Set Conn = Server.Createobject("ADODB.Connection")   
Conn.Open connectstr



strSQL="SELECT * FROM MEMDB WHERE Tim='Tim' AND TimDone='no' ORDER BY ODATE, Cnam ASC"


Set objRec = Server.CreateObject("ADODB.Recordset")   
objRec.Open strSQL, Conn, 1,3   
  
If objRec.EOF Then  
Response.write (" No records found.") %>
<p class="style14 style16">Open Jobs</p>


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


  <p class="style14 style21">Open Jobs</p>
     <form action="designer.asp" method="post" target="_self">
  <input type="hidden" name="pass" value="<%=strpass%>">
  <input type="submit" value="Refresh">
  </form> 

	
	
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
  <%end If %>
  </span>
  <table width="980" border="1" bgcolor="#999999">
    <%
Do While Not objRec.EOF and No <= PageLen   

i=i+1



%>
    <div align="center"><span class="style8"><br>
      </span>
        <div align="center"></div>
      <tr width="980">
          <td width="175" align="center"><span class="style8 style16"><span class="style32">Job #</span><%=objRec.Fields("RecNo").Value%></span></td>
        <td align="center" width="387"><span class="style27"><span class="style32">Job Name</span>:<%=objRec.Fields("JNam").Value%></span></td>
        <td align="center" width="210"><div align="center"><span class="style27"><span class="style32">Order Date</span>:<%=objRec.Fields("ODate").Value%></span></div></td>
        <td align="center" width="180"><div align="center"><span class="style27"><span class="style32">Finger Size</span>:<%=objRec.Fields("FSize").Value%></span></div></td>
        <td rowspan="4"><span class="style8"><a href="uploadtester3.asp?pwd=<%=strpass%>&recno=<%=objRec.Fields("RecNo").Value%>&JNam=<%=objRec.Fields("JNam").Value%>">Upload Files</a><br> <br><br>
        </span><span class="style8"><a href="designer.asp?pass=<%=strpass%>&rec=<%=objRec.Fields("RecNo").Value%>&button=Complete">SUBMIT DONE</a> </span></td>
      </tr>
        <tr>
          <td align="center" width="175"><span class="style33">Stone Information</span></td>
          <td  colspan="3" align="left" ><span class="style35"><%=objRec.Fields("Sinfo").Value%></span></td>
      </tr>
        <tr>
          <td width="175" align="center"><span class="style33">Instructions:</span></td>
          <td colspan="3" align="left"><span class="style35"><%=objRec.Fields("TimInst").Value%></span></td>
      </tr>
        <tr>
          <td width="175" align="center"><span class="style21"><strong>Send Message:</strong></span></td>
          <td colspan="3" align="left"><form action="designer.asp" method="post" target="_self" id="formal">
              <span class="style8">
                <input type="text" name="news" >
              <input type="hidden" name="pass2" value="<%=strPass%>">
              <input type="hidden" name="R" value="<%=objRec.Fields("RecNo").Value%>">
              <input type="hidden" name="TI" value="<%=objRec.Fields("TimInst").Value%>">
              <input type="submit" name="but" value="send">
            </span>
          </form>
              <span class="style8">
                <%dim strTimMes
 
 strTimMes=objRec.Fields("TimMes").Value
 
 if strTimMes="yes" then
 %>
                <span class="style30">New Message:</span> <span class="style29">NOT READ</span>
              <%end if
 
 if strTimMes="no" then
%>
                <span class="style28">NEW REPLY</span>
              <%end if%>
              </span></td>
        </tr>
    </div>
    <% Dim curjob
	curjob=objRec.Fields("RecNo").Value 




%>
    <tr>
      <td colspan="4"><span class="style8">
        <%

Dim Conn3,strSQL3,objRec3, curfile
Set Conn3 = Server.Createobject("ADODB.Connection")   
Conn3.Open connectstr



strSQL3 = "SELECT filenam from Uploads where job='"&curjob&"'"

Set objRec3 = Server.CreateObject("ADODB.Recordset")   
objRec3.Open strSQL3, Conn3, 1,3

Do While Not objRec3.EOF

curfile=objRec3.Fields("Filenam").Value
%>
        <a href="./uploads/<%=curfile%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('999','','./uploads/<%=curfile%>',1)"><img name="Image2" border="0" height="60" width="80" src="./uploads/<%=curfile%>"></a>
        <%
objRec3.MoveNext   
Loop  

objRec3.Close()   
Conn3.Close()   
Set objRec3 = Nothing  
Set Conn3 = Nothing 

%>
      </span>
      <%Dim Conn8,strSQL8,objRec8, curjob8, curfile8
Set Conn8 = Server.Createobject("ADODB.Connection")   
Conn8.Open connectstr

curjob8=objRec.Fields("RecNo").Value


strSQL8 = "SELECT filenam from files where RecNo='"&curjob8&"'"

Set objRec8 = Server.CreateObject("ADODB.Recordset")   
objRec8.Open strSQL8, Conn8, 1,3



Do While Not objRec8.EOF

curfile8=objRec8.Fields("Filenam").Value

if curfile8="" then

else%>


<%=curfile8%>***
<%end if
objRec8.MoveNext   
Loop  
      
    objRec.MoveNext   
Loop  %>  
      </td>
    </tr>
  </table>
  <span class="style8">
  <p class="style14 style16">Submitted Jobs</p>
  <table bgcolor="#999999" border="1">
    <tr>      
      <td align="center" width="40"><span class="style16">Job #</span></td>
                
  
   <td width="387" align="center" class="style16">Job Name</td>
      <td align="center" width="40"><span class="style16">Order Date</span></td>
    
   
      <td width="40" align="center" class="style16">Finger Size</td>
    </tr>   <%
Dim Conn4,strSQL4,objRec4
Set Conn4 = Server.Createobject("ADODB.Connection")   
Conn4.Open connectstr
Set objRec4 = Server.CreateObject("ADODB.Recordset")   



strSQL4="SELECT * FROM MEMDB WHERE Tim='Tim' AND TimDone='yes'"
 

objRec4.Open strSQL4, Conn4, 1,3

Do While Not objRec4.EOF
%>
  <form action="jobupdate2.asp?pass=<%=strpass2%>" method="post" target="_self">
    <tr>
      
      <input type="hidden" name="Jnum" value="<%=objRec4.Fields("RecNo").Value%>">
      
      <td width="40" align="center"><span class="style19"><%=objRec4.Fields("RecNo").Value%></span></td>
            

          <td align="center" width="387"><span class="style19"><%=objRec4.Fields("JNam").Value%></span></td>
          <td align="center" width="40"><div align="center"><%=objRec4.Fields("ODate").Value%></div></td>
  
    
      <td align="center" width="40"><div align="center"><%=objRec4.Fields("FSize").Value%></div></td>
      <td rowspan="4">
        <%

Dim Conn6,strSQL6,objRec6, curfile2, curjob2
Set Conn6 = Server.Createobject("ADODB.Connection")   
Conn6.Open connectstr

curjob2=objRec4.Fields("RecNo").Value

strSQL6 = "SELECT filenam from Uploads where job='"&curjob2&"'"

Set objRec6 = Server.CreateObject("ADODB.Recordset")   
objRec6.Open strSQL6, Conn6, 1,3

Do While Not objRec6.EOF

curfile2=objRec6.Fields("Filenam").Value
%>
        
  <a href="./uploads/<%=curfile2%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('999','','./uploads/<%=curfile2%>',1)"><img name="Image2" border="0" height="60" width="80" src="./uploads/<%=curfile2%>"></a>
        
  <%
objRec6.MoveNext   
Loop  

objRec6.Close()   
Conn6.Close()   
Set objRec6 = Nothing  
Set Conn6 = Nothing 




%>      </td>
	  <td>
	    
	    <A href="designer.asp?pass=<%=strpass%>&rec=<%=objRec4.Fields("RecNo").Value%>&button=ReOpen">ReOpen Job</A>      </td>
  </tr>
    
  <%
objRec4.MoveNext   
Loop  

objRec4.Close() 

Conn4.Close()   
 Set objRec4=nothing
Set Conn4 = Nothing 
%>
  </table>
  <%' start COMPLETED JOBS%>
  </span><span class="style23">
  <p class="style24">Completed</p>
  </span>
  <table bgcolor="#999999" border="1">
 <tr>      
               <td align="center" width="40"><span class="style16">Job #</span></td>
              
  
 <td width="150" align="center" class="style16">Job Name</td>
    <td align="center" width="40"><span class="style16">Order Date</span></td>
  
   
    <td width="40" align="center" class="style16">Finger Size</td>
   

   
 
  
    </tr>  
	
	
	<%
Dim Conn14,strSQL14,objRec14
Set Conn14 = Server.Createobject("ADODB.Connection")   
Conn14.Open connectstr
Set objRec14 = Server.CreateObject("ADODB.Recordset")   



strSQL14="SELECT * FROM MEMDB WHERE Tim='done'"
 

objRec14.Open strSQL14, Conn14

Do While Not objRec14.EOF
%>
<form action="jobupdate2.asp?pass=<%=strpass2%>" method="post" target="_self">
      <tr>
         
             <input type="hidden" name="Jnum" value="<%=objRec14.Fields("RecNo").Value%>">
             
            <td width="40" align="center"><span class="style19"><%=objRec14.Fields("RecNo").Value%></span></td>
          

            <td align="center" width="150"><span class="style19"><%=objRec14.Fields("JNam").Value%></span></td>
            <td align="center" width="40"><div align="center"><%=objRec14.Fields("ODate").Value%></div></td>
			<td><div align="center"><%=objRec14.Fields("FSize").Value%></div></td>

    
    
    <%

Dim Conn16,strSQL16,objRec16, curfile12, curjob12
Set Conn16 = Server.Createobject("ADODB.Connection")   
Conn16.Open connectstr

curjob12=objRec14.Fields("RecNo").Value

strSQL16 = "SELECT filenam from Uploads where job='"&curjob12&"'"

Set objRec16 = Server.CreateObject("ADODB.Recordset")   
objRec16.Open strSQL16, Conn16

Do While Not objRec16.EOF

curfile12=objRec16.Fields("Filenam").Value
%>
<td>
<a href="./uploads/<%=curfile12%>" target="_blank" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','./uploads/<%=curfile12%>',1)"><img name="Image2" border="0" height="60" width="80" src="./uploads/<%=curfile12%>"></a>
</td>
<%
objRec16.MoveNext   
Loop  

objRec16.Close()   
Conn16.Close()   
Set objRec16 = Nothing  
Set Conn16 = Nothing 




%>


    
    
    
    </td>

  
    


</tr>

<%
objRec14.MoveNext   
Loop  

objRec14.Close() 

Conn14.Close()   
 Set objRec14=nothing
Set Conn14 = Nothing 
%>


   
    
   
</table>



<%
' END COMPLETED JOBS

objRec.Close()   
Conn.Close()   
Set objRec = Nothing  
Set Conn = Nothing  
end if

%>   


</body>   
</html>  
