<html>
<head>
    <script type="text/vbscript">
        ' This function retrieves the selected options' text and value from the dropdown
        Sub GetSelectedOptions()
            Dim dropdown, selectedTexts, selectedValues, i, selectedText, selectedValue
            Set dropdown = Document.getElementById("myDropdown")

            selectedTexts = ""
            selectedValues = ""

            ' Loop through all options and check if they are selected
            For i = 0 To dropdown.options.length - 1
                If dropdown.options(i).selected Then
                    selectedText = dropdown.options(i).text
                    selectedValue = dropdown.options(i).value

                    ' Append the selected texts and values with a comma separator
                    If selectedTexts <> "" Then
                        selectedTexts = selectedTexts & ","
                        selectedValues = selectedValues & ","
                    End If

                    selectedTexts = selectedTexts & selectedText
                    selectedValues = selectedValues & selectedValue
                End If
            Next

            ' Check if any options are selected and display them as a comma-separated string
            If selectedTexts <> "" Then
                MsgBox "Selected Texts: " & selectedTexts
                MsgBox "Selected Values: " & selectedValues
            Else
                MsgBox "No options selected."
            End If
        End Sub
    </script>
</head>
<body>

    <form>
        <select id="myDropdown" multiple>
            <option value="1">Option 1</option>
            <option value="2">Option 2</option>
            <option value="3">Option 3</option>
            <option value="4">Option 4</option>
        </select>

        <br/>
        <input type="button" value="Get Selected Options" onclick="GetSelectedOptions()" />
    </form>

</body>
</html>
