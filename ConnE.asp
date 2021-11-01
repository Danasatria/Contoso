<%
    Set Rs = Server.CreateObject("ADODB.recordset")
    Set Conn = Server.CreateObject("ADODB.Connection")
    Set ObjCmd = Server.CreateObject("ADODB.Command")
    
    Conn.ConnectionString = "Provider=SQLOLEDB; Data Source=MSI; initial catalog=ContosoUniv; User ID=sa;password=dana10"
    Conn.open
    
    Dim id
    Dim Form_CourseID
    Dim Form_StudentID
    Dim Form_Grade

    Form_CourseID = request.form("CourseID")
    Form_StudentID = request.form("StudentID")
    Form_Grade = request.form("Grade")
    id = request.form("ID")

    if (trim(id) = "") or (isnull(id)) then id = 0 end if

    if cint(id) <> 0 then
        Dim Sql1 
        Sql1 = "UPDATE Enrollment Set CourseID='" & Form_CourseID & "',StudentID='" & Form_StudentID & "',Grade='" & Form_Grade & "' Where ID = " & id

        Set ObjCmd.ActiveConnection = Conn 

        ObjCmd.CommandText = Sql1
    
        Set RS = ObjCmd.Execute()

        response.redirect("Enroll-table.asp")

        Conn.Close
    else
        Dim Sql
        Sql = "INSERT INTO Enrollment(CourseID,StudentID,Grade) Values('" & Form_CourseID & "','" & Form_StudentID & "','" & Form_Grade & "')"

        Set ObjCmd.ActiveConnection = Conn 

        ObjCmd.CommandText = Sql
    
        Set RS = ObjCmd.Execute()

        response.redirect("Enroll-table.asp")

        Conn.Close

    end if

%>