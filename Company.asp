<!--#include virtual="head.asp"-->
<!---->
<div id="bottom">
<div id="leftmenu">
<ul>
<li class="l1" onclick="javascript:location.href='Company.asp'">结算单位</li>
<li class="l" onclick="javascript:location.href='Manager.asp'">项目经理</li>
<li class="l" onclick="javascript:location.href='Subject.asp'">费用项目</li>
<li class="l" onclick="javascript:location.href='Class.asp'">类　　别</li>
<li class="l" onclick="javascript:location.href='Income.asp'">收入方式</li>
</ul>
</div>
<div id="rightContent">
<div id="rightContentadd2">
<%

act=request("act")
id=request("id")
Cus_name=left(trim(request("Cus_name")),10)
Cus_name2=left(trim(request("Cus_name2")),30)
if act="add" then
'验证用户名是否存在

	sql="select Cus_name from Customer where Cus_name='"&Cus_name&"'"  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('“"&Cus_name&"”此名称已经存在，请修改重试');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if

'添加
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Customer "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Cus_name")=Cus_name
	 news("Cus_name2")=Cus_name2
	 news("Cus_time")=now()
	 news("Cus_OpeID")=request.cookies("wuliuid") '用户id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('提交成功');location.href('Company.asp');</script>"
end if
'修改
if act="mod" then
'验证用户名是否存在
	sql="select Cus_name from Customer where Cus_name='"&Cus_name&"'and Cus_id<>"&id  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)
	
	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('“"&Cus_name&"”此名称已经存在，请修改重试');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if
'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Customer where Cus_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Cus_name")=Cus_name
	 news("Cus_name2")=Cus_name2
	 news("Cus_time")=now()
	 news("Cus_OpeID")=request.cookies("wuliuid") '用户id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('修改成功');location.href('Company.asp');</script>"
end if

'删除数据
if act="d" then

dsql="delete from Customer where Cus_id="&id
response.write dsql
conn.execute dsql
response.redirect"Company.asp"
end if


'搜索企业
if act="search" then

  if Cus_name<>"" then 
  sqlSe="where Cus_name like '%"&Cus_name&"%' or Cus_name2 like '%"&Cus_name&"%'"
  end if
end if%>




<%if act="m" and id<>""  then%>
<%set modrs=server.createobject("adodb.recordset")
modsql="select * from Customer where Cus_Id="&id

modrs.open modsql,conn,3,3%>
<form name="wuliuform" method="post" action="?id=<%=id%>" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="178">企业简称 <input name="Cus_name" type="text" id="Cus_name" size="15" value="<%=trim(modrs("Cus_name"))%>"/></td>
      <td width="138"><p></p></td>
      <td width="184">全称
        <input name="Cus_name2" type="text" id="Cus_name2" size="20" value="<%=trim(modrs("Cus_name2"))%>"/></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="修改企业" /><input type="hidden" name="act" value="mod" /></td>
    </tr>
  </table>
</form>
<%else%>
<form name="wuliuform" method="post" action="" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="180">企业简称 <input name="Cus_name" type="text" id="Cus_name" size="15" /></td>
      <td width="136"><p></p></td>
      <td width="184">全称
        <input name="Cus_name2" type="text" id="Cus_name2" size="20" /></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="新增企业" /><input type="hidden" name="act" value="add" /></td>
    </tr>
  </table>
</form>
<%end if%>
<form name="wuliuform" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="180">企业名（全称或简称） </td>
      <td><p>
        <input name="Cus_name" type="text" id="Cus_name" size="25" value="<%=Cus_name%>"/>
      </p>        <!--全称
        <input name="Cus_name2" type="text" id="Cus_name2" size="20" />--></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="搜索企业" /><input type="hidden" name="act" value="search" /></td>
    </tr>
  </table>
</form>
<script>
//提交验证
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 
  var oName = document.wuliuform.Sup_name.value;;
      if (oName==""){
	
	wuliuform.Sup_name.focus();
	
      return false;

    }
	else if(reg.test(oName)){
			wuliuform.Sup_name.focus();
	
      return false;
		}

    else{
	   return true;

    }
}

//过程验证
window.onload=function(){
  var aInput = document.getElementsByTagName('input');
  var oName = aInput[0];
    var oName2 = aInput[1];
  var aP = document.getElementsByTagName('p');
  var name_msg = aP[0];
    var name_msg2 = aP[1];
   var name_length = 0;
//简称

  oName.onfocus = function(){
    name_msg.style.display = "inline";
    name_msg.innerHTML = "<i class='info'></i>推荐使用中文名";
  }

  oName.onblur = function(){

    //含有非法字符            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 


    //不能为空
    if (this.value==""){
	
      name_msg.innerHTML = "<i class='no'></i>不能为空！";
    }
    else if(reg.test(this.value)){
      name_msg.innerHTML = '<i class="no"></i>含有非法字符！';
    }

    //OK
    else {
      name_msg.innerHTML = "<i class='yes'></i>OK！";

    }
  }
  
  //全称
    oName2.onfocus = function(){
    name_msg2.style.display = "inline";
    name_msg2.innerHTML = "<i class='info'></i>推荐使用中文名";
  }

  oName2.onblur = function(){

    //含有非法字符            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 


    //不能为空
    if (this.value==""){
	
      name_msg2.innerHTML = "<i class='no'></i>不能为空！";
    }
    else if(reg.test(this.value)){
      name_msg2.innerHTML = '<i class="no"></i>含有非法字符！';
    }

    //OK
    else {
      name_msg2.innerHTML = "<i class='yes'></i>OK！";

    }
  }

  }


</script>
</div>
<div id="rightContentlist">
<table style="width:100%">
<tr><th>系统ID</th>
<th>企业简称</th>
<th>企业全称</th><th>操作</th></tr>
<%  
'开始分页

dim intPage,page,pre,last,filepath 
'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Customer "&sqlSe&" order by Cus_OrderId desc,Cus_id"

rs.PageSize = 50000 '这里设定每页显示的记录数
rs.CursorLocation = 3

rs.open sql,conn,3,3
if err.number<>0 then
				response.write "数据库中暂时无相关数据"
				end if
				if rs.eof And rs.bof then
       				Response.Write "<p align='center' > 对不起，没有查询到您需要的信息！</p>"
   				else
	  				pre = true
last = true
page = trim(Request.QueryString("page"))

if len(page) = 0 then
intpage = 1
pre = false
else
if cint(page) =< 1 then
intpage = 1
pre = false
else
if cint(page) >= rs.PageCount then
intpage = rs.PageCount
last = false
else
intpage = cint(page)
end if
end if
end if
if not rs.eof then
rs.AbsolutePage = intpage
end if 
do while not rs.eof and count<rs.PageSize
%> 
<tr onmousemove="changeTrColor(this)"><td><%=rs("Cus_id")%></td><td><%=rs("Cus_name")%></td><td><%=rs("Cus_name2")%></td><td><a href="?act=m&id=<%=rs("Cus_id")%>"><img src="images/m.gif" border=0/></a>　<a href="javascript:del()"><img src="images/d.gif" border=0/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del(){                   
 if(window.confirm("确实要删除“<%=trim(rs("Cus_name"))%>”吗？")){            
  window.location = "Company.asp?act=d&id=<%=rs("Cus_id")%>"; 
  //提交的url         
  }else{             
  return;         
  }   
    } //--> 
    </SCRIPT></td></tr>
<%
 rs.movenext
 count=count+1
 loop
end if
%>  
</table>
<script type="text/javascript">
function changeTrColor(obj){    
    var _table=obj.parentNode;
    for (var i=0;i<_table.rows.length;i++){
        _table.rows[i].style.backgroundColor="";
    }    
    obj.style.backgroundColor="#E2CFBC";
}
</script>
</div>
</div>
</div>

</body>
</html>
