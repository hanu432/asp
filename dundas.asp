<% 
' Server-side code for processing the form after it is submitted
Dim uploader, uploadFolder, fileUploaded, fileRenamed, impersonated
Dim fileId, objectId, objectCode, newFileName, fileType, fileComments, returnCode

' Initialize variables
fileId = ""
impersonated = False
fileUploaded = False
fileRenamed = False
uploadFolder = Server.MapPath("\FileStorageTmp") & "\"

' Create Dundas uploader object
Set uploader = Server.CreateObject("Dundas.Upload")

' Handle impersonation if needed (optional)
ipu = getSystemConfigKey(server.MapPath("/config/system.config"),"","Application","IPU")
If ipu <> "" Then
    ipw = getSystemConfigKey(server.MapPath("/config/system.config"),"","Application","IPW")
    uploader.ImpersonateUser(ipu, ipw, "US", 8) ' Example impersonation
    impersonated = True
End If

' Handle form submission
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Set max file size (optional)
    uploader.MaxFileSize = 10485760 ' 10MB

    ' Save the uploaded file to the server
    uploader.Save uploadFolder ' Save the uploaded file to the folder

    ' Collect form values
    fileId = uploader.Form.Item("file_id")
    newFileName = uploader.Form.Item("new_file_name")
    fileComments = uploader.Form.Item("file_comments")
    objectId = uploader.Form.Item("object_id")
    objectCode = uploader.Form.Item("object_code")
    returnCode = uploader.Form.Item("return_code")

    ' Check if the file was uploaded successfully
    If uploader.FileExists(uploader.Files(0).Path) Then
        fileUploaded = True
        Response.Write "<script>alert('File uploaded successfully!');</script>"
    Else
        fileUploaded = False
        Response.Write "<script>alert('File upload failed.');</script>"
    End If
End If
%>

<!-- HTML Form for File Upload -->
<form method="POST" enctype="multipart/form-data">
    <h3>File Upload</h3>

    <!-- Hidden Fields -->
    <input type="hidden" name="file_id" value="12345">
    <input type="hidden" name="new_file_name" value="uploadedFileName.txt">
    <input type="hidden" name="file_comments" value="Some file comments">
    <input type="hidden" name="object_id" value="001">
    <input type="hidden" name="object_code" value="XYZ123">
    <input type="hidden" name="return_code" value="200">

    <!-- File Upload Input -->
    <input type="file" name="uploadedFile">

    <!-- Submit Button -->
    <input type="submit" value="Upload">
</form>

<!-- Display Message After Submission -->
<%
If fileUploaded Then
    Response.Write "<p style='color: green;'>File uploaded successfully. File ID: " & fileId & "</p>"
Else
    Response.Write "<p style='color: red;'>Failed to upload the file.</p>"
End If
%>
