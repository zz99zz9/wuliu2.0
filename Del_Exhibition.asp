<%EExh_id=request.QueryString("Exh_id")%>
<body style="background:#D4D0C8;">
<!--#include virtual="inc/conn.asp"-->
<%
	sql="select Exp_Exhid from Expense where Exp_Exhid="&EExh_id&""  ' 查询数据库中是否有重复记录


	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('此展会尚有支出费用存在，不能删除！');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if
	rs.close
	set rs=nothing
	%>
	
    <%
	sql="select Rev_Exhid from Revenue where Rev_Exhid='"&EExh_id&"'"  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('此展会尚有收入费用存在，不能删除！');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if
		rs.close
	set rs=nothing
	
	%>
	
    <%
	dsql="delete from Exhibition where Exh_id="&EExh_id
'response.write dsql
conn.execute dsql
	 Response.Write ("<script language='javascript'>alert('展会信息删除成功！');location.href('index.asp');</script>") ' 返回结果并进行编码转义

%>