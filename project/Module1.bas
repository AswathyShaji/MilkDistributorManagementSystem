Attribute VB_Name = "Module1"
Public adocn As New ADODB.Connection
Public username As String
Public pswd As String
Public names As String
Public status As String
Public check As Boolean


Public Sub main()
adocn.ConnectionString = "Provider=SQLNCLI10.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MDMS;Data Source=ASUS-PC"
adocn.CursorLocation = adUseClient
adocn.Open
'FRM_USERSALARY.Show
FRM_LOGIN.Show
'MDIForm1.Show
End Sub
