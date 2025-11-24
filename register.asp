<%




Dim oConn, oRs
Dim qry, connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
dim db_pic


db_server = "jwk124.db.7626218.hostedresource.com"
db_name = "jwk124"
db_username = "jwk124"
db_userpassword = "Mon1tor!"
fieldname = "RecNo"
tablename = "MemDB"


connectstr = "Provider=SQLNCLI;SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_username & ";PWD=" & db_userpassword
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open connectstr

qry = "SELECT * FROM MemDb"

Set oRS = oConn.Execute(qry)


%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" />
<meta name="HandheldFriendly" content="true" />
<title>Form</title>
<link href="http://cdn.jotfor.ms/static/formCss.css?3.1.773" rel="stylesheet" type="text/css" />
<link type="text/css" rel="stylesheet" href="http://cdn.jotfor.ms/css/styles/nova.css?3.1.773" />
<link type="text/css" media="print" rel="stylesheet" href="http://cdn.jotfor.ms/css/printForm.css?3.1.773" />
<style type="text/css"> 
    .form-label{
        width:75px !important;
    }
    .form-label-left{
        width:75px !important;
    }
    .form-line{
        padding-top:12px;
        padding-bottom:12px;
    }
    .form-label-right{
        width:75px !important;
    }
    body, html{
        margin:0;
        padding:0;
        background:false;
    }
 
    .form-all{
        margin:0px auto;
        padding-top:0px;
        width:590px;
        color:Black !important;
        font-family:'Lucida Grande',' Lucida Sans Unicode',' Lucida Sans',' Verdana',' Tahoma',' sans-serif';
        font-size:14px;
    }
</style>
 
<script src="http://cdn.jotfor.ms/static/jotform.js?3.1.773" type="text/javascript"></script>
<script type="text/javascript"> 
   JotForm.init(function(){
      $('input_30').hint('ex: myname@example.com');
   });
</script>
</head>
<body>
<form class="jotform-form" action="registerok.asp" method="post" name="regform" >
  <input type="hidden" name="formID" value="31651605269152" />
  <div class="form-all">
    <ul class="form-section">
      <li id="cid_29" class="form-input-wide">
        <div class="form-header-group">
          <h2 id="header_29" class="form-header">
            Please register for an account
          </h2>
          <div id="subHeader_29" class="form-subHeader">
            Please fill out every field (CASE SENSITIVE)
          </div>
        </div>
      </li>
      <li class="form-line" id="id_15">
        <label class="form-label-left" id="label_15" for="input_15"> Name </label>
        <div id="cid_15" class="form-input"><span class="form-sub-label-container"><input class="form-textbox" type="text" size="10" name="fname" id="first_15" />
            <label class="form-sub-label" for="first_15" id="sublabel_first"> First Name </label></span><span class="form-sub-label-container"><input class="form-textbox" type="text" size="15" name="lname" id="last_15" />
            <label class="form-sub-label" for="last_15" id="sublabel_last"> Last Name </label></span>
        </div>
      </li>
      <li class="form-line form-line-column" id="id_19">
        <label class="form-label-left" id="label_19" for="input_19"> Password</label>
        <div id="cid_19" class="form-input">
          <input type="text" class=" form-textbox" data-type="input-textbox" id="input_19" name="pass" size="32" />
        </div>
      </li>
      <li class="form-line" id="id_27">
        <label class="form-label-left" id="label_27" for="input_27">Password (verify)</label>
        <div id="cid_27" class="form-input">
          <input type="text" class=" form-textbox" data-type="input-textbox" id="input_27" name="pass2" size="32" />
        </div>
      </li>
      <li class="form-line" id="id_30">
        <label class="form-label-left" id="label_30" for="input_30"> E-mail </label>
        <div id="cid_30" class="form-input">
          <input type="email" class=" form-textbox validate[Email]" id="input_30" name="email" size="32" />
        </div>
      </li>
	   <li class="form-line" id="id_39">
        <label class="form-label-left" id="label_30" for="input_39"> E-mail (verify)</label>
        <div id="cid_30" class="form-input">
          <input type="email" class=" form-textbox validate[Email]" id="input_39" name="email2" size="32" />
        </div>
      </li>
      <li class="form-line" id="id_31">
        <label class="form-label-left" id="label_31" for="input_31"> Phone Number </label>
        <div id="cid_31" class="form-input"><span class="form-sub-label-container"><input class="form-textbox" type="tel" name="Acod" id="input_31_area" size="3">
            -
            <label class="form-sub-label" for="input_31_area" id="sublabel_area"> Area Code </label></span><span class="form-sub-label-container"><input class="form-textbox" type="tel" name="number" id="input_31_phone" size="9">
            <label class="form-sub-label" for="input_31_phone" id="sublabel_phone"> Phone Number </label></span>
        </div>
      </li>
      <li class="form-line" id="id_28">Do you have 2 years experience?
        <label class="form-label-top" id="label_28" for="input_28"></label>
        <div id="cid_28" class="form-input-wide">
          <div class="form-multiple-column"><span class="form-radio-item"><input type="radio" class="form-radio" id="input_28_0" name="yup" value="Yes" />
              <label for="input_28_0"> Yes </label></span><span class="clearfix"></span><span class="form-radio-item"><input type="radio" class="form-radio" id="input_28_1" name="yup" value="No" />
              <label for="input_28_1"> No </label></span><span class="clearfix"></span>
          </div>
        </div>
      </li>
      <li class="form-line" id="id_26">
        <div id="cid_26" class="form-input-wide">
          <div style="text-align:left" class="form-buttons-wrapper">
            <input type="submit" name="GOGO" value="Register">
          </div>
        </div>
      </li>
    
    
    </ul>
  </div>

</form>






     <%Do until oRs.EOF%>


<%oRS.MoveNext
Loop
oRs.Close


Set oRs = nothing
Set oConn = nothing

%>
</body>
</html>


	     

