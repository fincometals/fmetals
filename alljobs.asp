<%




Dim oConn, oRs, strJN
Dim qry, connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic


db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "JobDB"


connectstr = "Provider=SQLNCLI;SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open connectstr

qry = "SELECT * FROM ORDERS WHERE Status='OPEN'"

Set oRS = oConn.Execute(qry)


%>


<HTML>
<HEAD>
<TITLE>
</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
body {
	background-image: url(http://www.fincometals.com/watches/gearb.jpg);
	background-repeat: repeat-y;
}
.style20 {
	font-size: 60px;
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
</HEAD>

<BODY onLoad="MM_preloadImages('http://www.fincometals.com/watches/banner.png')">
<div align="center">
  <p><span class="style20">OPEN JOBS</span></p>
  <table width="980" border="0" align="center">
  <tr>
    <td align="center" width="40">Job #</td>
    <td align="center" width="100">Customer</td>
    <td align="center" width="100">Customer ID</td>
    <td align="center" width="60">Order Date</td>
    <td align="center" width="60">Due Date</td>
    <td align="center" width="75">Job Name</td>
    <td align="center" width="50">Finger Size</td>
    <td align="center" width="50">Metal Type</td>
    <td align="center" width="100">Stone Information</td>
    <td align="center" width="100">Remarks</td>
    <td align="center" width="40">Status</td>
    <td align="center" width="80">Tracking Info</td>
  </tr>
</table>
  
            <%Do until oRs.EOF%>
          </div>

<p>&nbsp;</p>
<%strJN="<%Response.Write oRs.Fields("JNum")%>"%>
<table width="980" border="1" align="center">
  <tr>
    <td width="40" align="center"><%Response.Write oRs.Fields("JNum")%></td>
    <td width="100" align="center"><%Response.Write oRs.Fields("CNam")%></td>
    <td width="100" align="center"><%Response.Write oRs.Fields("CID")%></td>
    <td align="center" width="60"><%Response.Write oRs.Fields("ODate")%></td>
    <td align="center" width="60"><%Response.Write oRs.Fields("DDate")%></td>
    <td align="center" width="75"><%Response.Write oRs.Fields("JNam")%></td>
    <td align="center" width="50"><%Response.Write oRs.Fields("FSize")%></td>
    <td align="center" width="50"><%Response.Write oRs.Fields("MType")%></td>
    <td align="center" width="100"><%Response.Write oRs.Fields("Stones")%></td>
    <td align="center" width="100"><%Response.Write oRs.Fields("Remarks")%></td>
    <td align="center" width="40"><%Response.Write oRs.Fields("Status")%></td>
    <td align="center" width="80"><%Response.Write oRs.Fields("Tracking")%></td>
    <td>
    <form action="jobupdate.asp" method="post" target="_blank">
        
           <input type="hidden" width="0" name="JobNum" value="<%=oRs.Fields("JNum").Value%>"/>
            
 
    <input type="submit" value="CLOSE" />
      
        </form></td>
  </tr>
</table>
<br>


<%oRS.MoveNext
Loop
oRs.Close


Set oRs = nothing
Set oConn = nothing

%>
	     
</BODY>
</HTML> 
