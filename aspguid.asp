<%
    ' Dynamically generate a unique token on the server side
    Function CreateGUID()
        Dim objGUID
        Set objGUID = CreateObject("Scriptlet.TypeLib")
        CreateGUID = objGUID.GUID
        Set objGUID = Nothing
    End Function

    ' Generate a unique token for the form submission
    Dim uploadToken
    uploadToken = CreateGUID() ' Replace with your method to generate GUID if needed
%>

<!DOCTYPE html>
<html>
<head>
    <title>File Upload Form</title>
</head>
<body>

    <h2>File Upload Example</h2>

    <!-- Form to submit the file and hidden token -->
    <form method="post" action="<%= Request.ServerVariables("URL") %>">
        <!-- Hidden field to pass the token -->
        <input type="hidden" name="upload_token" value="<%= uploadToken %>">
        <input type="submit" value="Submit">
    </form>

    <% 
        ' Check if form has been submitted
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
            Dim formToken
            ' Retrieve the hidden field value
            formToken = Trim(Request.Form("upload_token"))
        
            ' Output the value of upload_token for debugging
            Response.Write "Form Token: " & formToken & "<br>"
        
            ' Validate and display message
            If Len(formToken) = 0 Then
                Response.Write "<p>Error: upload_token is missing or empty!</p>"
            Else
                Response.Write "<p>upload_token received successfully!</p>"
            End If
        End If
    %>

</body>
</html>
