<%
    id = Request.QueryString("ID")
    Dim Data
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.open("Provider=SQLOLEDB; Data Source=MSI; initial catalog=ContosoUniv; User ID=sa;password=dana10")
    if (trim(id) = "") or (isnull(id)) then id = 0 end if

    if (cint(id) <> 0) then 
        Set Data = Conn.execute("Select * from Enrollment where ID = " & id)
        if not Data.EOF then
            ID = Data("ID")

            Set Data = nothing
        end if
    end if
%>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Enrollment Form</title>

    <link href="css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.2/css/all.css" integrity="sha384-fnmOCqbTlWIlj8LyTjo7mOUStjsKC4pOpQbqyi7RrhN7udi9RwhKkMHpvLbHG9Sr" crossorigin="anonymous">
    <link href="assets/css/style.css" rel="stylesheet">
    <link href="assets/css/components.css" rel="stylesheet">
    <link rel="shortcut icon" href="assets/img/logo_starbhak.ico">
</head>
<body>
    <div id="app">
        <div class="main-wrapper">
            <div class="navbar-bg"></div>
            <nav class="navbar navbar-expand-lg main-navbar">
                <form class="form-inline mr-auto">
                    <ul class="navbar-nav mr-3">
                        <li><a href="#" data-toggle="sidebar" class="nav-link nav-link-lg"><i class="fas fa-bars"></i></a></li>
                        <li><a href="#" data-toggle="search" class="nav-link nav-link-lg d-sm-none"><i class="fas fa-search"></i></a></li>
                    </ul>
                    <div class="search-element">
                        <input class="form-control" type="search" placeholder="Search" aria-label="Search" data-width="250">
                        <button class="btn" type="submit"><i class="fas fa-search"></i></button>
                        <div class="search-backdrop"></div>
                        <div class="search-result">
                            <div class="search-header">
                                Histories
                            </div>
                            <div class="search-item">
                                <a href="#">How to hack NASA using CSS</a>
                                <a href="#" class="search-close"><i class="fas fa-times"></i></a>
                            </div>
                            <div class="search-item">
                                <a href="#">Kodinger.com</a>
                                <a href="#" class="search-close"><i class="fas fa-times"></i></a>
                            </div>
                            <div class="search-item">
                                <a href="#">#Stisla</a>
                                <a href="#" class="search-close"><i class="fas fa-times"></i></a>
                            </div>
                            <div class="search-header">
                                Result
                            </div>
                            <div class="search-item">
                                <a href="#">
                                    <img class="mr-3 rounded" width="30" src="assets/img/products/product-3-50.png" alt="product">
                                    oPhone S9 Limited Edition
                                </a>
                            </div>
                            <div class="search-item">
                                <a href="#">
                                    <img class="mr-3 rounded" width="30" src="assets/img/products/product-2-50.png" alt="product">
                                    Drone X2 New Gen-7
                                </a>
                            </div>
                            <div class="search-item">
                                <a href="#">
                                    <img class="mr-3 rounded" width="30" src="assets/img/products/product-1-50.png" alt="product">
                                    Headphone Blitz
                                </a>
                            </div>
                            <div class="search-header">
                                Projects
                            </div>
                            <div class="search-item">
                                <a href="#">
                                    <div class="search-icon bg-danger text-white mr-3">
                                        <i class="fas fa-code"></i>
                                    </div>
                                    Stisla Admin Template
                                </a>
                            </div>
                            <div class="search-item">
                                <a href="#">
                                    <div class="search-icon bg-primary text-white mr-3">
                                        <i class="fas fa-laptop"></i>
                                    </div>
                                    Create a new Homepage Design
                                </a>
                            </div>
                        </div>
                    </div>
                </form>
                <ul class="navbar-nav navbar-right">
                    <li class="dropdown dropdown-list-toggle"><a href="#" data-toggle="dropdown" class="nav-link nav-link-lg message-toggle beep"><i class="far fa-envelope"></i></a>
                        <div class="dropdown-menu dropdown-list dropdown-menu-right">
                            <div class="dropdown-header">Messages
                                <div class="float-right">
                                    <a href="#">Mark All As Read</a>
                                </div>
                            </div>
                            <div class="dropdown-list-content dropdown-list-message">
                                
                            </div>
                            <div class="dropdown-footer text-center">
                                <a href="#">View All <i class="fas fa-chevron-right"></i></a>
                            </div>
                        </div>
                    </li>
                    <li class="dropdown dropdown-list-toggle"><a href="#" data-toggle="dropdown" class="nav-link notification-toggle nav-link-lg beep"><i class="far fa-bell"></i></a>
                        <div class="dropdown-menu dropdown-list dropdown-menu-right">
                            <div class="dropdown-header">Notifications
                                <div class="float-right">
                                    <a href="#">Mark All As Read</a>
                                </div>
                            </div>
                            <div class="dropdown-list-content dropdown-list-icons">
                                
                            </div>
                            <div class="dropdown-footer text-center">
                                <a href="#">View All <i class="fas fa-chevron-right"></i></a>
                            </div>
                        </div>
                    </li>
                    <li class="dropdown"><a href="#" data-toggle="dropdown" class="nav-link dropdown-toggle nav-link-lg nav-link-user">
                        <img alt="image" src="assets/img/avatar/avatar-1.png" class="rounded-circle mr-1">
                        <div class="d-sm-none d-lg-inline-block">Hi, Ujang Maman</div></a>
                        <div class="dropdown-menu dropdown-menu-right">
                            <div class="dropdown-title">Logged in 5 min ago</div>
                            <a href="features-profile.html" class="dropdown-item has-icon">
                                <i class="far fa-user"></i> Profile
                            </a>
                            <a href="features-activities.html" class="dropdown-item has-icon">
                                <i class="fas fa-bolt"></i> Activities
                            </a>
                            <a href="features-settings.html" class="dropdown-item has-icon">
                                <i class="fas fa-cog"></i> Settings
                            </a>
                            <div class="dropdown-divider"></div>
                            <a href="#" class="dropdown-item has-icon text-danger">
                                <i class="fas fa-sign-out-alt"></i> Logout
                            </a>
                        </div>
                    </li>
                </ul>
            </nav>
            <div class="main-sidebar">
                <aside id="sidebar-wrapper">
                    <div class="sidebar-brand">
                        <a href="index.asp">Contoso</a>
                    </div>
                    <div class="sidebar-brand sidebar-brand-sm">
                        <a href="index.asp">CU</a>
                    </div>
                    <ul class="sidebar-menu">
                        <li class="menu-header">Dashboard</li>
                        <li><a class="nav-link" href="Index.asp"><i class="fas fa-home"></i><span>Home</span></a></li>
                        <li class="menu-header">Data</li>
                        <li class="nav-item dropdown">
                            <a href="#" class="nav-link has-dropdown active" data-toggle="dropdown"><i class="far fa-user"></i><span>Student</span></a>
                            <ul class="dropdown-menu">
                                <li><a class="nav-link" href="Student-table.asp"><i class="fas fa-table"></i><span>Table</span></a></li>
                                <li><a class="nav-link" href="Student-form.asp"><i class="fas fa-user-plus"></i><span>Add Data</span></a></li>
                            </ul>
                        </li>
                        <li class="nav-item dropdown">
                            <a href="#" class="nav-link has-dropdown" data-toggle="dropdown"><i class="fas fa-columns"></i><span>Course</span></a>
                            <ul class="dropdown-menu">
                                <li><a class="nav-link" href="Course-table.asp"><i class="fas fa-table"></i><span>Table</span></a></li>
                                <li><a class="nav-link" href="Course-form.asp"><i class="fas fa-plus"></i><span>Add Data</span></a></li>
                            </ul>
                        </li>
                        <li class="nav-item dropdown">
                            <a href="#" class="nav-link has-dropdown" data-toggle="dropdown"><i class="far fa-file-alt"></i><span>Enrollment</span></a>
                            <ul class="dropdown-menu">
                                <li><a class="nav-link" href="Enroll-table.asp"><i class="fas fa-table"></i><span>Table</span></a></li>
                                <li class="active"><a class="nav-link" href="Enroll-form.asp"><i class="fas fa-plus"></i><span>Add Data</span></a></li>
                            </ul>
                        </li>
                    </ul>
                </aside>
            </div>

            <!-- Main Content -->
            <div class="main-content">
                <section class="section">
                <div class="section-header">
                    <h1>Enrollment Form</h1>
                    <div class="section-header-breadcrumb">
                        <div class="breadcrumb-item active"><a href="Index.asp">Dashboard</a></div>
                        <div class="breadcrumb-item active"><a href="Enroll-table.asp">Table</a></div>
                        <div class="breadcrumb-item">Form</div>
                    </div>
                </div>

                <div class="section-body">
                    <div class="card">
                        <div class="card-body">
                            <form method="post" action="ConnE.asp">
                                <% 
                                    Dim DataS 
                                    Dim DataC
                                    Set Conn = Server.CreateObject("ADODB.Connection")
                                    Set Rec = Server.CreateObject("ADODB.recordset")
                                    Conn.open("Provider=SQLOLEDB; Data Source=MSI; initial catalog=ContosoUniv; User ID=sa;password=dana10")
    
                                    Set DataS = Conn.execute("Select * from Student")
                                    Set DataC = Conn.execute("Select * from Course")
                                %>
                                <div class="form-group">
                                    <input class="form-control" type="Hidden" name="ID" value="<%=ID%>" disable>
                                </div>
                                <div class="form-group">
                                    <label>Course</label>
                                    <select class="form-control" name="CourseID">
                                    <%
                                        do while not DataC.EOF
                                    %>
                                        <option value="<%=DataC("ID")%>"><%=DataC("Title")%></option>
                                    <%
                                        DataC.MoveNext()
                                        Loop
                                    %>
                                    </select>
                                </div>

                                <div class="form-group">
                                    <label>Student</label>
                                    <select class="form-control" name="StudentID">
                                    <%
                                        do while not DataS.EOF
                                    %>
                                        <option value="<%=DataS("ID")%>"><%=DataS("LastName")%> <span><%=DataS("FirstMidName")%></span></option>
                                    <%
                                        DataS.MoveNext()
                                        Loop
                                    %>
                                    </select>
                                </div>

                                <div class="form-group">
                                    <label>Grade</label>
                                    <select class="form-control" name="Grade">
                                        <option value="1">A</option>
                                        <option value="2">B</option>
                                        <option value="3">C</option>
                                        <option value="4">D</option>
                                        <option value="5">F</option>
                                    </select>
                                </div>
                                <div class="d-flex justify-content-end">
                                    <a href="Student-table.asp" class="btn btn-icon icon-left btn-danger mr-2"><i class="fas fa-times"></i> Back</a>
                                    <button type="submit" class="btn btn-icon icon-left btn-success "><i class="fas fa-check"></i> Submit</button>
                                </div>
                            </form>
                        </div>
                    </div>
                </section>
            </div>
        </div>
    </div>



















































    <!-- General JS Scripts -->
    <script src="https://code.jquery.com/jquery-3.3.1.min.js" integrity="sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8=" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js" integrity="sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js" integrity="sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.nicescroll/3.7.6/jquery.nicescroll.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.24.0/moment.min.js"></script>

    <!-- Template JS File -->
    <script src="js/bootstrap.min.js"></script>
    <script src="assets/js/stisla.js"></script>
    <script src="assets/js/scripts.js"></script>
    <script src="assets/js/custom.js"></script>
</body>

</html>