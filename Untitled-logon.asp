<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/AA.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_AA_STRING
Recordset1_cmd.CommandText = "SELECT * FROM tb1" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>DREAMER子网页-3</title>
<link rel="shortcut icon" href="pic/LOGO111.ico)" />
 <script type="text/javascript">function del(s,d,a){var textareaVal = document.getElementById(s);var textareaVal_old = document.getElementById(a);if(textareaVal.value == 1){ textareaVal_old.value=d.value;  textareaVal.value = 2;  d.value = '';}else if(d.value.length < 1){ d.value = textareaVal_old.value; textareaVal.value = 1;}}</script>
<style>
img{vertical-align:middle;}
li{font-family:"方正兰亭超细黑简体"}
a{text-decoration:none;}
a{color:#9DCD69;}
img{border:0;}

#top{
	width:100%;
	height:50px;
	background:#FFF;}
#topLOGO{
	float:left;
	width:10px;}
#topLOGON{
	float:left;
	width:75px;
	margin:12px 0 0 1090px;}
#upbox{
	width:1200px;
	height:50px;
	margin:0 auto;}
#medium{
	width:100%;
	height:600px;
	background-color:#E8E8E8;
	padding:10px 0 0 0;}
#mediumbox{
	background:url(pic/ban.png);
	width:1200px;
	height:500px;
	margin:10px auto;
	padding:0;}
#juzhong{
	clear:both;
	float:left;
	width:380px;
	height:180px;
	margin:100px 430px;
	padding:0;}
#zhuanli{
	float:left;
	color:#CECECE;
	font-family:"Adobe 黑体 Std R";
	font-size:13px;
	width:300px;
	margin:90px 0 0 0;}
#BOTTOMJUZHONGlogo{
	float:right;
	width:170px;
	height:70px;
	margin:30px 0 0 500px;}
#bottombox{
	width:1200px;
	height:110px;
	margin:0 auto;}
#bottom{
	clear:both;
	width:auto;
	height:150px;
	background-color:#FFF;}
#down{
	width:auto;
	height:150px;
	clear:both;
	background:#E8E8E8;}
#LOGONLV{
	width:270px;
	height:30px;
	margin:0 50px;
	padding:0;}
#nav{
	font-size:12px;
	float:left;
	background-color:#FFF;
	width:380px;
	height:48px;
	margin:2px auto;}
#nav ul{
	float:left;
	list-style:none;
	margin:6px 0 0 50px;
	padding:0;}
#nav ul li{
	padding:10px 10px;
	float:left;
	text-decoration:none;}
.button{
	width:262px;
	height:32px;
	margin:0;
	border:none;}
.sdw{
	border-radius:3px;border:1px solid #CECECE;}
 .sdw:focus{
transition:border linear .2s,box-shadow linear .5s; -moz-transition:border linear .2s,-moz-box-shadow linear .5s; -webkit-transition:border linear .2s,-webkit-box-shadow linear .5s; outline:none;border-color:rgba(12,164,245,.75); box-shadow:0 0 8px rgba(12,164,235,.5); -moz-box-shadow:0 0 8px rgba(12,164,220,.5); -webkit-box-shadow:0 0 8px rgba(12,164,220,3);border-radius:3px;border:1px solid #0CA4F5;}
}
</style>

<script type="text/javascript" src="youxiang.js"></script>
</head>

<body>
<div id="top">
<div id="upbox">
 <div id="topLOGO"><a target="_blank" href="untitled.html"><img src="pic/LOGO.png" width="160" height="45" /></div>
 <div id="topLOGON"><img src="pic/user.png" width="18" height="21" /><a href=Untitled-2.html><font color="#9DCD69" size="-1" face="方正兰亭超细黑简体" >登录/注册</font></a></div>
</div>
</div>

<div id="medium">
<div id="mediumbox">
<form action="" method="get">
<div id="juzhong">
  <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="hidden" value="1" id="textarea">
  <input type="hidden" name="textarea_old_value" id="textareaVal_old">
  <input class="sdw" type="text" onkeyup="value=value.replace(/[^\w\.\/]/ig,'')"onfocus="del('textarea',this,'textareaVal_old');" onblur="del('textarea',this,'textareaVal_old');" value="邮箱" size="40px" style="text-align:center;height:23px;color:#CECECE"/></p> 
  <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="hidden" value="1" id="textarea" >
  <input type="hidden" name="textarea_old_value" id="textareaVal_old">
  <input class="sdw" type="text" onfocus="del('textarea',this,'textareaVal_old');" onblur="del('textarea',this,'textareaVal_old');" value="用户密码" size="40px" style=" text-align:center;height:23px;color:#CECECE"  /> </p> 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="logon" type="submit" class="button" style="background:url(pic/LOGONLV.png)" border="none" value=""/>
</form>
<div id="nav">
  <ul>
    <a href="#"><li>注册</li></a>
    <a href="#"><li>忘记密码？</li></a>
    <a href="#"><li>没有收到验证邮箱？</li></a>
  </ul>
</div>
  
</div>

</div>
</div>

<div id="bottom">
<div id="bottombox">
<div id="zhuanli">Copyright©20 16 DREAMER  川ICP备15058894号</div>
<div id="BOTTOMJUZHONGlogo"><img src="pic/LOGO下.png" width="160" height="65" /></div>
</div>
</div>
<div id="down"></div>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
