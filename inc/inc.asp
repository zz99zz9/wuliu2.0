<%
'****************************************
'��֤��Ŀ�����Ƿ����
'call Check_Sup_name()
'****************************************
sub Check_Sup_name()
sql="select Sup_name from Supervisor where Sup_name='"&Sup_name&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼

set rs = conn.execute(sql)

If not(rs.Eof And rs.Bof) Then
 Response.Write ("<script language='javascript'>alert('��"&Sup_name&"�����û����Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
response.end()
end if
end sub

'****************************************
'�ͻ�idת���ͻ�������д
'call Show_customer_name()
'****************************************
Sub Show_customer_name(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
sql="select * from [Customer] where Cus_id="&countt
'response.write sql

oRs.Open sql,conn,1,1
If Not oRs.eof Then
Content=oRs("Cus_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
Response.write trim(Content)
End sub
'****************************************
'��Ŀidת����Ŀ����
'call Show_Subject_name()
'****************************************
Sub Show_Subject_name(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [Subject] where Sub_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("Sub_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
Response.write Content
End sub
'****************************************
'����idת����������
'call Show_Revenue_name()
'****************************************
Sub Show_Income_name(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [Income] where Inc_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("Inc_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
Response.write Content
End sub
'****************************************
'����Աidת������
'call Show_operator_name()
'****************************************
Sub Show_operator_name(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [operator] where Ope_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("Ope_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
Response.write Content
End sub
'****************************************
'��Ŀ����idת������
'call Show_Supervisor_name()
'****************************************
Sub Show_Supervisor_name(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [Supervisor] where Sup_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("Sup_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
Response.write Content
End sub
'****************************************
'��Ŀidת��Ӧ�տ����
'call Revenue_sum1()
'****************************************
function Revenue_sum1(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
sql="select sum(Rev_amount1) as sum_amount from [hz_wuliu_xgwl].[dbo].[Revenue] where Rev_Exhid="&countt
'response.write sql
oRs.Open sql,conn,1,1
If oRs("sum_amount")<>"" Then
Content=FormatNumber(oRs("sum_amount"))
Else
Content="0.00"
End if
oRs.close
set oRs=Nothing

Response.write Content
'Response.write FormatNumber(Content,2)
Revenue_sum11=Content
end function
'****************************************
'��Ŀidת�����տ����
'call Revenue_sum2()
'****************************************
function Revenue_sum2(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
sql="select sum(Rev_amount2) as sum_amount from [hz_wuliu_xgwl].[dbo].[Revenue] where Rev_Exhid="&countt
oRs.Open sql,conn,1,1
'response.write sql
If oRs("sum_amount")<>"" and oRs("sum_amount")<>0 Then
Content=FormatNumber(oRs("sum_amount"))
Else
Content="0.00"
End if

oRs.close
set oRs=Nothing

Response.write Content
'Response.write FormatNumber(Content,2)
Revenue_sum22=Content
End function
'****************************************
'��Ŀidת���Ѹ������
'call Expense_sum1()
'****************************************
function Expense_sum1(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
sql="select sum(Exp_amount1) as sum_amount from [Expense] where Exp_Exhid="&countt
'response.write sql
oRs.Open sql,conn,1,1
If oRs("sum_amount")<>"" Then
Content=FormatNumber(trim(oRs("sum_amount")))
Else
Content="0.00"
End if

oRs.close
set oRs=Nothing
Response.write Content
Expense_sum11=Content
End function
'****************************************
'��Ŀidת���Ѹ������
'call Expense_sum2()
'****************************************
function Expense_sum2(countt)

set oRs=Server.CreateObject("ADODB.Recordset")
sql="select sum(Exp_amount2) as sum_amount from [Expense] where Exp_Exhid="&countt
'response.write sql
oRs.Open sql,conn,1,1
If oRs("sum_amount")<>"" Then
Content=FormatNumber(trim(oRs("sum_amount")))
Else
Content="0.00"
End if

oRs.close
set oRs=Nothing
Response.write Content
Expense_sum22=Content
End function
'****************************************
'Ԥ������Ӧ��-Ӧ����
'call yjlr()
'****************************************
function yjlr(countt)
set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select sum(Rev_amount1) as sum_amount from [Revenue] where Rev_Exhid="&countt,conn,1,1
if oRs("sum_amount")<>"" then
ys=oRs("sum_amount")
else
ys=0
end if
oRs.close
set oRs=Nothing

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select sum(Exp_amount1) as sum_amount from [Expense] where Exp_Exhid="&countt,conn,1,1
if oRs("sum_amount")<>"" then
yf=oRs("sum_amount")
else
yf=0
end if
oRs.close
set oRs=Nothing
if ys-yf=0 then
response.write "0.00"
else
Response.write FormatNumber(ys-yf)
end if
yjlrhz=int(ys)-int(yf)
End function
'****************************************
'Ŀǰ��������-�Ѹ���
'call mqlr()
'****************************************
function mqlr(countt)
set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select sum(Rev_amount2) as sum_amount from [Revenue] where Rev_Exhid="&countt,conn,1,1
if oRs("sum_amount")<>"" then
ys=oRs("sum_amount")
else
ys=0
end if
oRs.close
set oRs=Nothing

set oRs=Server.CreateObject("ADODB.Recordset")
sql="select sum(Exp_amount2) as sum_amount from [Expense] where Exp_Exhid="&countt
'response.write sql
oRs.Open sql,conn,1,1
if oRs("sum_amount")<>"" then
yf=oRs("sum_amount")
else
yf=0
end if
oRs.close
set oRs=Nothing
if ys-yf=0 then
response.write "0.00"
else
Response.write formatnumber(ys-yf)
end if
mqlrhz=int(ys)-int(yf)
End function
'****************************************
'δ�տ����-���գ�
'call mqlr()
'****************************************
function wsk(countt)
set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select sum(Rev_amount1) as sum_amount from [Revenue] where Rev_Exhid="&countt,conn,1,1
if oRs("sum_amount")<>"" then
ys=oRs("sum_amount")
else
ys=0
end if
oRs.close
set oRs=Nothing

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select sum(Rev_amount2) as sum_amount from [Revenue] where Rev_Exhid="&countt,conn,1,1
if oRs("sum_amount")<>"" then
ys2=oRs("sum_amount")
else
ys2=0
end if
oRs.close
set oRs=Nothing
if ys-ys2=0 then
response.write "0.00"
else
Response.write FormatNumber(ys-ys2)
end if
wskhz=int(ys)-int(ys2)
End function
'****************************************
'��Ŀ���idת���������
'call Show_class_name()
'****************************************
Sub Show_class_name(countt)

set oRs=Server.CreateObject("ADODB.Recordset")

oRs.Open "select * from Class where cla_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("cla_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
Response.write Content
End sub
'****************************************
'չ��idת��չ����������
'call Show_exh2mas_name()
'****************************************
Sub Show_exh2mas_name(countt)

if trim(countt)<>"" then
set oRs=Server.CreateObject("ADODB.Recordset")

oRs.Open "select * from Exhibition where exh_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("Exh_supid")
Else
Content="��Ч���"
End if
oRs.close
oRs.Open "select * from Supervisor where Sup_id="&Content,conn,1,1
If Not oRs.eof Then
Content=oRs("sup_name")
Else
Content="��Ч���"
End if
oRs.close
set oRs=Nothing
end if
Response.write Content
End sub
'****************************************
'չ��idת��չ����루չ���ţ�
'call Show_exh_code()
'****************************************
Sub Show_exh_code(countt)
if trim(countt)<>"" then
set oRs=Server.CreateObject("ADODB.Recordset")

 oRs.Open "select * from Exhibition where exh_id="&countt,conn,1,1
 If Not oRs.eof Then
 Content=oRs("Exh_code")
 Else
 Content="��Ч���"
 End if

oRs.close
set oRs=Nothing
end if
Response.write Content
End sub
'****************************************
'չ��idת��չ������
'call Show_exh_name()
'****************************************
Sub Show_exh_name(countt)
if trim(countt)<>"" then
set oRs=Server.CreateObject("ADODB.Recordset")

oRs.Open "select * from Exhibition where exh_id="&countt,conn,1,1
If Not oRs.eof Then
Content=oRs("Exh_name")
Else
Content="��Ч���"
End if

oRs.close
set oRs=Nothing
end if
Response.write Content
End sub
%>