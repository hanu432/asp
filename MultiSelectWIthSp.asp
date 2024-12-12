<%
' Database connection variables
Dim conn, cmd, rs
Dim connString, spName

' Set up the connection string (update with your server, database, user, and password details)
connString = "Provider=SQLOLEDB;Data Source=your_server;Initial Catalog=your_database;User ID=your_user;Password=your_password;"
spName = "GetDropdownOptions"

' Create connection object
Set conn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set rs = Server.CreateObject("ADODB.Recordset")

' Open the connection
conn.Open connString

' Configure the command object to call the stored procedure
cmd.ActiveConnection = conn
cmd.CommandText = spName
cmd.CommandType = 4 ' 4 indicates a stored procedure

' Execute the command and get the results
Set rs = cmd.Execute

%>

<!-- Include Bootstrap CSS -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css">
<!-- Include Bootstrap Multiselect CSS -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/css/bootstrap-multiselect.css">

<main>
    <div class="container">
        <h1>Multi-Select Dropdown with Checkboxes</h1>
        <form id="dropdownForm" method="post" action="process.asp">
            <label for="multiSelect">Select Options (Max: 10):</label>
            <select id="multiSelect" name="options[]" multiple="multiple" class="form-control">
                <%
                ' Populate dropdown options from SQL data
                Do While Not rs.EOF
                    Response.Write("<option value='" & rs("Id") & "'>" & rs("Name") & "</option>")
                    rs.MoveNext
                Loop
                %>
            </select>
        </form>
    </div>

    <!-- Include jQuery -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <!-- Include Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/js/bootstrap.bundle.min.js"></script>
    <!-- Include Bootstrap Multiselect JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/js/bootstrap-multiselect.min.js"></script>

    <script>
        $(document).ready(function () {
            // Initialize Bootstrap Multiselect
            $('#multiSelect').multiselect({
                includeSelectAllOption: true,
                enableFiltering: true,
                buttonWidth: '300px',
                nonSelectedText: 'Select Options'
            });
        });
    </script>
</main>

<%
' Clean up objects
rs.Close
Set rs = Nothing
Set cmd = Nothing
conn.Close
Set conn = Nothing
%>
