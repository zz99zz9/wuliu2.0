<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>物流结算系统</title>
<link rel="stylesheet" type="text/css" href="css/Public.css"/>
<%if request.cookies("wuliuid")="" then
response.Redirect("login.asp")
end if
%>
<!--#include virtual="inc/conn.asp"-->
<!--#include virtual="inc/inc.asp"-->
<!--#include virtual="inc/md5.asp"-->
</head>
<%if ECount="" then%>
<body>
<%else%>
<body onload="javascript:ChangeDiv('<%=Riframe%>','zhlbr','zhb',3);ChangeDiv('<%=zhlb%>','zhlb','zhbt',2);">
<%end if%>
<div id="topmenu">
<div id="topmenuleft">
<ul>
<%if request.cookies("wuliuv")=0 then%>
<li class="b b6" onclick="javascript:location.href='Index.asp'"><span>交易管理</span></li>
<li class="b b1" onclick="javascript:location.href='Create_Exhibition.asp'"><span>创建展会</span></li>
<li class="b b7" onclick="javascript:location.href='Chartlist.asp'"><span>统计报表</span></li>
<li class="b b2" onclick="javascript:location.href='Operator.asp'"><span>操作员设置</span></li>
<li class="b b3" onclick="javascript:location.href='Operator_v.asp'"><span>访问权限设置</span></li>
<li class="b b4" onclick="javascript:location.href='Company.asp'"><span>预设管理</span></li><%end if%>
<li class="b b5" onclick="javascript:location.href='Quit.asp'"><span>安全退出</span></li>
</ul>
</div>

<div id="topmenuright">当前操作人员：<%=request.cookies("wuliuuser")%></div>
</div>