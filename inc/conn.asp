<%Dim Conn,ConnStr 
ConnStr="Driver={SQL Server};Server=(local);Uid=sa;Pwd=123456;Database=hz_wuliu_xgwl;"
'On Error Resume Next 
Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Open ConnStr%>

<%
sub CloseConn()
	conn.close
	set conn=nothing
end sub

Sub CloseoRs()
oRs.close
set oRs=nothing
End sub

Sub CloseRs()
Rs.close
set Rs=nothing
End sub
%>
