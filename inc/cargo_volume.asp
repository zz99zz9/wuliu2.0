      	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
    <!--<td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd2')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd2')"><img src="images/l4.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>-->
    <td width="60" align="center"><a href="Out_Accounts2.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td height="60">&nbsp;</td>
    </tr>
</table>
</div>


<div>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>��Ŀ���</th>
      <th>չ������</th>
      <th>���</th>
            <th>��Ŀ����</th>
      <th>չ����</th>
      <th>���</th>
      <th>����</th>

    </tr>
    <%  
'��ʼ��ҳ


'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition order by Exh_id desc"
rs.PageSize = 100 '�����趨ÿҳ��ʾ�ļ�¼��
rs.CursorLocation = 3

rs.open sql,conn,3,3
if err.number<>0 then
				response.write "���ݿ�����ʱ������"
				end if
				if rs.eof And rs.bof then
       				Response.Write "<p align='center' > �Բ���û�в�ѯ������Ҫ����Ϣ��</p>"
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
if rs("Exp_amount2")<>rs("Exp_amount1") then
%> 
    <tr onmousemove="changeTrColor(this)" <%if rs("Exp_amount1")<=rs("Exp_amount2") and rs("Exp_amount1")<>0 then%> style="color:#ff0000;"<%end if%>>

      <td><%=rs("Exh_code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%=rs("Exh_class")%></td>
      <td><%=rs("Exh_supid")%></td>
      <td>1</td>
      <td><%=rs("Exh_volume")%></td>
<td><%=rs("Exh_kg")%></td>

    </tr>
    <%		ys=ys+rs("Exp_amount1")
	yf=yf+rs("Exp_amount2")
		end if

	ye=ye+rs("Exp_amount1")-rs("Exp_amount2")
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