<%EExh_id=request.QueryString("Exh_id")%>
<body style="background:#D4D0C8;">
<!--#include virtual="inc/conn.asp"-->
<%
	sql="select Exp_Exhid from Expense where Exp_Exhid="&EExh_id&""  ' ��ѯ���ݿ����Ƿ����ظ���¼


	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��չ������֧�����ô��ڣ�����ɾ����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
	rs.close
	set rs=nothing
	%>
	
    <%
	sql="select Rev_Exhid from Revenue where Rev_Exhid='"&EExh_id&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��չ������������ô��ڣ�����ɾ����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
		rs.close
	set rs=nothing
	
	%>
	
    <%
	dsql="delete from Exhibition where Exh_id="&EExh_id
'response.write dsql
conn.execute dsql
	 Response.Write ("<script language='javascript'>alert('չ����Ϣɾ���ɹ���');location.href('index.asp');</script>") ' ���ؽ�������б���ת��

%>