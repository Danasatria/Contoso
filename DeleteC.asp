<%
    Set Rs = Server.CreateObject("ADODB.recordset")
    Set Conn = Server.CreateObject("ADODB.Connection")
    Set ObjCmd = Server.CreateObject("ADODB.Command")
    
    Conn.ConnectionString = "Provider=SQLOLEDB; Data Source=MSI; initial catalog=ContosoUniv; User ID=sa;password=dana10"
    Conn.open

    id = Request.QueryString("ID")

    if trim(id) = "" or isnull(id) or trim(id) = "0" then 
    
    end if

    Dim Sql
    Sql = "DELETE FROM Course WHERE ID = " & id
    
    Set ObjCmd.ActiveConnection = Conn 

    ObjCmd.CommandText = Sql
    
    Set RS = ObjCmd.Execute()

    response.redirect("Course-table.asp")

    Conn.Close  
%>