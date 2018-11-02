<!--#include virtual="head.asp"-->
<!---->
<%

act=request("act")
id=request("id")
Exh_code=request("Exh_code")
Exh_name=request("Exh_name")
Exh_class=request("Exh_class")
Exh_volume=request("Exh_volume")
Exh_kg=request("Exh_kg")
Exh_Supid=request("Exh_Supid")
Exh_mark=request("Exh_mark")
Exh_year=request("Exh_year")
Exh_moon=request("Exh_moon")
Exh_f=request("Exh_f")
w_time=Exh_year&"."&Exh_moon
if Exh_f="" then Exh_f=0

if act="add" then

'验证用户名是否存在

	sql="select Exh_code from Exhibition where Exh_code='"&Exh_code&"'"  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('“"&Exh_code&"”此工作号已经存在，请修改重试');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if

'添加
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Exhibition "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Exh_code")=Exh_code
	 news("Exh_name")=Exh_name
	 news("Exh_class")=Exh_class
	 news("Exh_volume")=Exh_volume
	 news("Exh_kg")=Exh_kg
	 news("Exh_Supid")=Exh_Supid
	 news("Exh_mark")=Exh_mark
	 news("Exh_year")=Exh_year
	 news("Exh_moon")=Exh_moon
	 news("Exh_favorites")=Exh_f
	 news("Exh_addtime")=now()
   news("w_time")=w_time
	 news("Exh_OpeID")=request.cookies("wuliuid") '用户id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('展会添加成功');location.href('index.asp');</script>"
end if


%>
<script>
//提交验证
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 
  var Exh_Code = document.wuliuform.Exh_Code.value;;
      if (Exh_Code==""){
	
	wuliuform.Exh_Code.focus();
	
      return false;

    }
	else if(reg.test(Exh_Code)){
			wuliuform.Exh_Code.focus();
	
      return false;
		}

    else{
	   return true;

    }
}

//过程验证
window.onload=function(){
  var aInput = document.getElementsByTagName('input');
  var oExh_Code = aInput[0];
  var oExh_name = aInput[1];
  var oExh_volume = aInput[2];
  var oExh_kg = aInput[3];
  var oExh_mark = aInput[4];

  var aP = document.getElementsByTagName('p');
  var Code_msg = aP[0];
  var name_msg = aP[1];
  var mark_msg = aP[6];
  var Code_length = 0;
//会员名

  oExh_Code.onfocus = function(){
    Code_msg.style.display = "inline";
    Code_msg.innerHTML = "<i class='info'></i>请正确输入工作号";
  }

  oExh_Code.onblur = function(){

    //含有非法字符            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 


    //不能为空
    if (this.value==""){
	
      Code_msg.innerHTML = "<i class='no'></i>不能为空！";
    }
    else if(reg.test(this.value)){
      Code_msg.innerHTML = '<i class="no"></i>含有非法字符！';
    }

    //OK
    else {
      Code_msg.innerHTML = "<i class='yes'></i>OK！";

    }
  }
  
  //展会名称验证
    oExh_name.onfocus = function(){
    name_msg.style.display = "inline";
    name_msg.innerHTML = "<i class='info'></i>建议使用中文名";
  }

  oExh_name.onblur = function(){

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
  //唛头验证
/*    oExh_mark.onfocus = function(){
    mark_msg.style.display = "inline";
    mark_msg.innerHTML = "<i class='info'></i>请输入唛头名称";
  }

  oExh_mark.onblur = function(){

    //含有非法字符            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 


    //不能为空
    if (this.value==""){
	
      mark_msg.innerHTML = "<i class='no'></i>不能为空！";
    }
/*    else if(reg.test(this.value)){
      mark_msg.innerHTML = '<i class="no"></i>含有非法字符！';
    }*/

    //OK
/*    else {
      mark_msg.innerHTML = "<i class='yes'></i>OK！";

    }
  }*/

  }


</script>
<div id="bottom">
  <div id="CenterContent">
  	<div id="Toptitle">创建展会</div>
    <div id="Content">
    <form name="wuliuform" method="post" action="?id=<%=id%>" onSubmit="return checkform();">
      <table width="100%" border="0" cellspacing="8" cellpadding="0">
        <tr>
          <td width="20%" height="30" align="right">工 作 号</td>
          <td width="45%"><input type="text" name="Exh_Code" id="Exh_Code" /></td>
          <td width="35%"><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">展会名称</td>
          <td><input type="text" name="Exh_name" id="Exh_name" /></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">类　　别</td>
          <td>
          <select name="Exh_class" id="Exh_class">
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Class order by Cla_OrderId desc,Cla_id"
rs.open sql,conn,3,3
do while not rs.eof%>
            <option value="<%=rs("Cla_id")%>"><%=rs("Cla_name")%></option>
<%
 rs.movenext

 loop
rs.close
set rs=nothing
%> 
          </select></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">体　　积</td>
          <td><input type="text" name="Exh_volume" id="Exh_volume" /></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">公　　斤</td>
          <td><input type="text" name="Exh_kg" id="Exh_kg" /></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">项目主管</td>
          <td>
          <select name="Exh_Supid" id="Exh_Supid">
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Supervisor order by Sup_OrderId desc,Sup_id"
rs.open sql,conn,3,3
do while not rs.eof%>
            <option value="<%=rs("Sup_id")%>"><%=rs("Sup_name")%></option>
<%
 rs.movenext

 loop
rs.close
set rs=nothing
%> 
          </select></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">年　　月</td>
          <td><select name="Exh_year" id="Exh_year">
            <option value="2016">2016</option>
            <option value="2017">2017</option>
            <option value="2018">2018</option>
            <option value="2019">2019</option>
            <option value="2020">2020</option>
            <option value="2021">2021</option>
            <option value="2022">2022</option>
          </select>
            <select name="Exh_moon" id="Exh_moon">
              <option value="1">1</option>
              <option value="2">2</option>
              <option value="3">3</option>
              <option value="4">4</option>
              <option value="5">5</option>
              <option value="6">6</option>
              <option value="7">7</option>
              <option value="8">8</option>
              <option value="9">9</option>
              <option value="10">10</option>
              <option value="11">11</option>
              <option value="12">12</option>
            </select></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="30" align="right">唛　　头</td>
          <td><input type="text" name="Exh_mark" id="Exh_mark" /></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">&nbsp;</td>
          <td>&nbsp;</td>
          <td></td>
        </tr>
        <tr>
          <td height="30" align="right">&nbsp;</td>
          <td><input type="submit" name="button" id="button" value="发布展会" /><input type="hidden" value="add" id="act" name="act" /></td>
          <td>&nbsp;</td>
        </tr>
      </table></form>
    </div>
  </div>
</div>

</body>
</html>
