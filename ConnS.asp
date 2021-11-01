<% 
    Set Rs = Server.CreateObject("ADODB.recordset")
    Set Conn = Server.CreateObject("ADODB.Connection")
    Set ObjCmd = Server.CreateObject("ADODB.Command")
    
    Conn.ConnectionString = "Provider=SQLOLEDB; Data Source=MSI; initial catalog=ContosoUniv; User ID=sa;password=dana10"
    Conn.open
    
    Dim id
    Dim From_LastName
    Dim From_FirstMidName
    Dim From_EnrollmentDate
    
    From_LastName= request.form("LastName")
    From_FirstMidName= request.form("FirstMidName")
    From_EnrollmentDate= request.form("EnrollmentDate")
    id= request.form("ID")

    if (trim(id) = "") or (isnull(id)) then id = 0 end if

    if cint(id) <> 0 then
        Dim Sql1
        Sql1 = "UPDATE Student Set LastName='" & From_LastName & "',FirstMidName='" & From_FirstMidName & "',EnrollmentDate= '" & From_EnrollmentDate & "' Where ID = " & id

        Set ObjCmd.ActiveConnection = Conn 

        ObjCmd.CommandText = Sql1
    
        Set RS = ObjCmd.Execute()

        response.redirect("Student-table.asp")

        Conn.Close
    else
        Dim Sql 
        Sql = "INSERT INTO Student(LastName,FirstMidName,EnrollmentDate) VALUES('" & From_LastName & "','" & From_FirstMidName & "','" & From_EnrollmentDate & "')"

        Set ObjCmd.ActiveConnection = Conn 

        ObjCmd.CommandText = Sql
    
        Set RS = ObjCmd.Execute()

        response.redirect("Student-table.asp")

        Conn.Close
    end if
    
%>

