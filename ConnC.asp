<% 
    Set Rs = Server.CreateObject("ADODB.recordset")
    Set Conn = Server.CreateObject("ADODB.Connection")
    Set ObjCmd = Server.CreateObject("ADODB.Command")
    
    Conn.ConnectionString = "Provider=SQLOLEDB; Data Source=MSI; initial catalog=ContosoUniv; User ID=sa;password=dana10"
    Conn.open
    
    Dim id
    Dim Form_Title
    Dim Form_Credits
    
    
    id= request.form("ID")
    Form_Title= request.form("Title")
    Form_Credits= request.form("Credits")
    

    if (trim(id) = "") or (isnull(id)) then id = 0 end if

    if cint(id) <> 0 then
        Dim Sql1
        Sql1 = "UPDATE Course Set Title='" & Form_Title & "',Credits='" & Form_Credits & "' Where ID = " & id

        Set ObjCmd.ActiveConnection = Conn 

        ObjCmd.CommandText = Sql1
    
        Set RS = ObjCmd.Execute()

        response.redirect("Course-table.asp")

        Conn.Close
    else
        Dim Sql 
        Sql = "INSERT INTO Course(Title,Credits) VALUES('" & Form_Title & "','" & Form_Credits & "')"

        Set ObjCmd.ActiveConnection = Conn 

        ObjCmd.CommandText = Sql
    
        Set RS = ObjCmd.Execute()

        response.redirect("Course-table.asp")

        Conn.Close
    end if
    
%>

