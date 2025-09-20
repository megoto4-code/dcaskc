<html>
<body>

<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open (Server.MapPath("COMMENTS.mdb"))

sql="INSERT INTO contact (nam,email,"
sql=sql & "cnam,query,txt)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("txtName") & "',"
sql=sql & "'" & Request.Form("txtEmail") & "',"
sql=sql & "'" & Request.Form("txtCname") & "',"
sql=sql & "'" & Request.Form("ddlQuery") & "',"
sql=sql & "'" & Request.Form("txtText") & "')"

on error resume next
conn.Execute sql,recaffected
if err<>0 then
  Response.Write("No update permissions!")
else
  Response.Redirect "submission_successful.htm"
end if
conn.close
Session.Abandon
%>

</body>
</html> 
