<%EExh_id=request.QueryString("Exh_id")%>

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
Exh_mark=trim(request("Exh_mark"))
Exh_year=request("Exh_year")
Exh_moon=request("Exh_moon")
Exh_f=request("Exh_f")
<!--  -->
if Exh_moon<10 then
    w_time=Exh_year&".0"&Exh_moon
    else
    w_time=Exh_year&"."&Exh_moon
    end if


if act="add" then

'��֤�û����Ƿ����



'���
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Exhibition where Exh_id="&EExh_id
	 news.open sql,conn,3,3	 
	' news.addnew	 
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
	 news("Exh_OpeID")=request.cookies("wuliuid") '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('չ����Ϣ�޸ĳɹ�');location.href('index.asp');</script>"
end if


%>
<script>
//�ύ��֤
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 
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

//������֤
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
//��Ա��

  oExh_Code.onfocus = function(){
    Code_msg.style.display = "inline";
    Code_msg.innerHTML = "<i class='info'></i>����ȷ���빤����";
  }

  oExh_Code.onblur = function(){

    //���зǷ��ַ�            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 


    //����Ϊ��
    if (this.value==""){
	
      Code_msg.innerHTML = "<i class='no'></i>����Ϊ�գ�";
    }
    else if(reg.test(this.value)){
      Code_msg.innerHTML = '<i class="no"></i>���зǷ��ַ���';
    }

    //OK
    else {
      Code_msg.innerHTML = "<i class='yes'></i>OK��";

    }
  }
  
  //չ��������֤
    oExh_name.onfocus = function(){
    name_msg.style.display = "inline";
    name_msg.innerHTML = "<i class='info'></i>����ʹ��������";
  }

  oExh_name.onblur = function(){

    //���зǷ��ַ�            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 


    //����Ϊ��
    if (this.value==""){
	
      name_msg.innerHTML = "<i class='no'></i>����Ϊ�գ�";
    }
    else if(reg.test(this.value)){
      name_msg.innerHTML = '<i class="no"></i>���зǷ��ַ���';
    }

    //OK
    else {
      name_msg.innerHTML = "<i class='yes'></i>OK��";

    }
  }
  //��ͷ��֤
/*    oExh_mark.onfocus = function(){
    mark_msg.style.display = "inline";
    mark_msg.innerHTML = "<i class='info'></i>��������ͷ����";
  }

  oExh_mark.onblur = function(){

    //���зǷ��ַ�            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 


    //����Ϊ��
    if (this.value==""){
	
      mark_msg.innerHTML = "<i class='no'></i>����Ϊ�գ�";
    }
/*    else if(reg.test(this.value)){
      mark_msg.innerHTML = '<i class="no"></i>���зǷ��ַ���';
    }*/

    //OK
/*    else {
      mark_msg.innerHTML = "<i class='yes'></i>OK��";

    }
  }*/

  }


</script>
<div id="bottom">
  <div id="CenterContent">
  	<div id="Toptitle">�޸�չ����Ϣ</div>
    <div id="Content">
    <%set rs1=server.createobject("adodb.recordset")
sql1="select * from Exhibition where Exh_id="&EExh_id
rs1.open sql1,conn,3,3%>
    <form name="wuliuform" method="post" action="?Exh_id=<%=EExh_id%>" onSubmit="return checkform();">
      <table width="100%" border="0" cellspacing="8" cellpadding="0">
        <tr>
          <td width="20%" height="30" align="right">�� �� ��</td>
          <td width="45%"><input type="text" name="Exh_Code" id="Exh_Code" value="<%=trim(rs1("Exh_Code"))%>"/></td>
          <td width="35%"><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">չ������</td>
          <td><input type="text" name="Exh_name" id="Exh_name" value="<%=trim(rs1("Exh_name"))%>"/></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">�ࡡ����</td>
          <td>
          <select name="Exh_class" id="Exh_class">
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Class order by Cla_OrderId desc,Cla_id"
rs.open sql,conn,3,3
do while not rs.eof%>
            <option value="<%=rs("Cla_id")%>" <%if int(rs1("Exh_class"))=int(rs("Cla_id")) then%> selected="selected"<%end if%>><%=rs("Cla_name")%></option>
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
          <td height="30" align="right">�塡����</td>
          <td><input type="text" name="Exh_volume" id="Exh_volume" value="<%=rs1("Exh_volume")%>"/></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">��������</td>
          <td><input type="text" name="Exh_kg" id="Exh_kg" value="<%=trim(rs1("Exh_kg"))%>"/></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">��Ŀ����</td>
          <td>
          <select name="Exh_Supid" id="Exh_Supid">
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Supervisor order by Sup_OrderId desc,Sup_id"
rs.open sql,conn,3,3
do while not rs.eof%>
            <option value="<%=rs("Sup_id")%>" <%if int(rs1("Exh_Supid"))=int(rs("Sup_id")) then%> selected="selected"<%end if%>><%=rs("Sup_name")%></option>
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
          <td height="30" align="right">�ꡡ����</td>
          <td><select name="Exh_year" id="Exh_year">
            <option value="2016" <%if int(rs1("Exh_year"))=2016 then%> selected="selected"<%end if%>>2016</option>
            <option value="2017" <%if int(rs1("Exh_year"))=2017 then%> selected="selected"<%end if%>>2017</option>
            <option value="2018" <%if int(rs1("Exh_year"))=2018 then%> selected="selected"<%end if%>>2018</option>
            <option value="2019" <%if int(rs1("Exh_year"))=2019 then%> selected="selected"<%end if%>>2019</option>
            <option value="2020" <%if int(rs1("Exh_year"))=2020 then%> selected="selected"<%end if%>>2020</option>
            <option value="2021" <%if int(rs1("Exh_year"))=2021 then%> selected="selected"<%end if%>>2021</option>
            <option value="2022" <%if int(rs1("Exh_year"))=2022 then%> selected="selected"<%end if%>>2022</option>
          </select>
            <select name="Exh_moon" id="Exh_moon">
              <option value="1" <%if int(rs1("Exh_moon"))=1 then%> selected="selected"<%end if%>>1</option>
              <option value="2" <%if int(rs1("Exh_moon"))=2 then%> selected="selected"<%end if%>>2</option>
              <option value="3" <%if int(rs1("Exh_moon"))=3 then%> selected="selected"<%end if%>>3</option>
              <option value="4" <%if int(rs1("Exh_moon"))=4 then%> selected="selected"<%end if%>>4</option>
              <option value="5" <%if int(rs1("Exh_moon"))=5 then%> selected="selected"<%end if%>>5</option>
              <option value="6" <%if int(rs1("Exh_moon"))=6 then%> selected="selected"<%end if%>>6</option>
              <option value="7" <%if int(rs1("Exh_moon"))=7 then%> selected="selected"<%end if%>>7</option>
              <option value="8" <%if int(rs1("Exh_moon"))=8 then%> selected="selected"<%end if%>>8</option>
              <option value="9" <%if int(rs1("Exh_moon"))=9 then%> selected="selected"<%end if%>>9</option>
              <option value="10" <%if int(rs1("Exh_moon"))=10 then%> selected="selected"<%end if%>>10</option>
              <option value="11" <%if int(rs1("Exh_moon"))=11 then%> selected="selected"<%end if%>>11</option>
              <option value="12" <%if int(rs1("Exh_moon"))=12 then%> selected="selected"<%end if%>>12</option>
            </select></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="30" align="right">�顡��ͷ</td>
          <td><input type="text" name="Exh_mark" id="Exh_mark" value="<%=trim(rs1("Exh_mark"))%>"/></td>
          <td><p></p></td>
        </tr>
        <tr>
          <td height="30" align="right">&nbsp;</td>
          <td>&nbsp;</td>
          <td></td>
        </tr>
        <tr>
          <td height="30" align="right">&nbsp;</td>
          <td colspan="2"><input type="submit" name="button" id="button" value="�����޸�" /><input type="hidden" value="add" id="act" name="act" />��<input type="button" name="Back" id="Back" value="������ҳ" onclick="javascript:history.back()"/>��<input type="button" name="Del" id="Del" value="ɾ��չ��" onclick="javascript:del<%=EExh_id%>()"/><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del<%=EExh_id%>(){                   
 if(window.confirm("ȷʵҪɾ����չ����Ϣ��")){            
  window.location = "Del_Exhibition.asp?Exh_id=<%=EExh_id%>"; 
  //�ύ��url         
  }else{             
  return;         
  }   
    } //--> 
    </SCRIPT></td>
          </tr>
      </table></form>
    </div>
  </div>
</div>

</body>
</html>
