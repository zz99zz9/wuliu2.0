      	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
    <!--<td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd2')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd2')"><img src="images/l4.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>-->
    <td width="60" align="center"><a href="Out_cargo_volume.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td height="60">&nbsp;</td>
    </tr>
</table>
</div>


<div>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>项目编号</th>
      <th>展会名称</th>
      <th>类别</th>
            <th>项目主管</th>
      <th>展商数</th>
      <th>体积</th>
      <th>公斤</th>

    </tr>
    <% Server.ScriptTimeOut=950 %>
    <%  
'开始分页


'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select top 30 * from Exhibition order by Exh_id desc"
rs.PageSize = 10000 '这里设定每页显示的记录数
rs.CursorLocation = 3

rs.open sql,conn,3,3
if err.number<>0 then
				response.write "数据库中暂时无数据"
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
do while not rs.eof

%> 
    <tr onmousemove="changeTrColor(this)" >

      <td><%=rs("Exh_code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%call Show_class_name(rs("Exh_class"))%></td>
      <td><%call Show_Supervisor_name(rs("Exh_supid"))%></td>
      <td><%call Show_exh_count(rs("Exh_id"))%></td>
      <td><%=rs("Exh_volume")%></td>
<td><%=rs("Exh_kg")%></td>

    </tr>
    <%	

 rs.movenext
 count=count+1
 loop
end if
rs.close
set rs=nothing
if ys="" or ys=0 then
ys="0.00"
else
ys=FormatNumber(ys)
end if

if yf="" or yf=0 then
yf="0.00"
else
yf=FormatNumber(yf)
end if

if ye="" or ye=0 then
ye="0.00"
else
ye=FormatNumber(ye)
end if
%> 
    <tr>
         <td></td>
              <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td><%=ys%></td>
<td><%=yf%></td>
<td><%=ye%></td>

      <td></td>


    </tr>
  </table>
  <script type="text/javascript">
function changeTrColor(obj){ 
    var _table=obj.parentNode;
    for (var i=0;i<_table.rows.length;i++){
        _table.rows[i].style.backgroundColor="";
    }    
    obj.style.backgroundColor="#FEE8EA";
}
</script>
</div>