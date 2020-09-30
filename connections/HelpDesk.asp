<%
' FileName="Connection_ado_conn_string.htm"
' Type="ADO"
' HTTP="false"
' Catalog=""
' Schema=""

'MM_HelpDesk_STRING = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=Helpdesk;Data Source=DEV2" 

'MM_HelpDesk_STRING = "Server=tcp:itdapsql.database.windows.net,1433;Initial Catalog=itdaphelpdeskdb;Persist Security Info=False;User ID=itdapadmin;Password={your_password};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

'MM_HelpDesk_STRING =     "Data Source=DEV\STSQLSERVER;Database=HelpDesk;Network Library=DBMSSOCN;User ID=sa;Password=;"


MM_HelpDesk_STRING = Server=tcp:itdapsql.database.windows.net,1433;Initial Catalog=HelpDesk;Persist Security Info=False;User ID=itdapadmin;Password=12345$$12345$$A;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;

%>