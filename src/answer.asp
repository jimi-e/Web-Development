<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
Dim connstr
connstr="provider=microsoft.ACE.oledb.12.0;data source=" & server.MapPath("score.accdb")
Set conn = Server.Createobject(ADODB.Connection")
conn.Open connstr
%>

<%
Set rs =Server.CreateObject("ADODB.Recordset")
rs.Open "score", conn,adOpenDynamic, 3
%>

<%
name = Request.form("inputname")
%>

<%
    dim score
    score = 0
    if Request.Form("q1") = "C" Then score = score + 10
    if Request.Form("q1") = "F" Then score = score + 10
    if Request.Form("q2") = "A" Then score = score + 20
    if Request.Form("q3") = "C" Then score = score + 20
    if Request.Form("q4") = "D" Then score = score + 10
    if Request.Form("q5") = "C" Then score = score + 10
    if Request.Form("q5") = "D" Then score = score + 10
    if Request.Form("q5") = "F" Then score = score + 10
    
%>

<%
rs.AddNew
rs("s_name") = name
rs("s_score") = score
rs.Update
%>

