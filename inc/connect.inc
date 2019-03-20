<%
ConStr ="Provider=SQLOLEDB;Data Source=aster.arvixe.com;database=Stresscom;uid=DDHUser;pwd=h0vercraft"

Set DbCon = Server.CreateObject("ADODB.Connection")
dbCon.open (conStr)
%>