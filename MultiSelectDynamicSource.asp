<%
' Mocking a stored procedure call with hardcoded data in an array
Dim options
options = Array("Option 1", "Option 2", "Option 3", "Option 4", "Option 5", "Option 6", "Option 7", "Option 8", "Option 9", "Option 10", "Option 11", "Option 12")

%>

<!-- Include Bootstrap CSS -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css">
<!-- Include Bootstrap Multiselect CSS -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.15/css/bootstrap-multiselect.css">
<!-- Include SweetAlert2 CSS -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">

<!-- Custom Styles -->
<style>
    body {
        font-family: Arial, sans-serif;
        margin: 20px;
    }
</style>

<main>
    <div class="container">
        <h1>Multi-Select Dropdown with Checkboxes</h1>
        <form id="dropdownForm" method="post" action="process.asp">
            <label for="multiSelect">Select Options (Max: 10):</label>
            <select id="multiSelect" name="options[]" multiple="multiple" class="form-control">
                <%
                ' Loop through the options array and generate the dropdown items dynamically
                Dim i
                For i = 0 To UBound(options)
                    Response.Write("<option value='" & (i + 1) & "'>" & options(i) & "</option>")
                Next
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
    <!-- Include SweetAlert2 JS -->
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js"></script>

    <script>
        $(document).ready(function () {
            // Initialize Bootstrap Multiselect
            $('#multiSelect').multiselect({
                includeSelectAllOption: true,  // Add a "Select All" checkbox
                enableFiltering: true,        // Enable search/filter functionality
                buttonWidth: '300px',         // Set button width
                nonSelectedText: 'Select Options' // Placeholder text
            });

            // Add real-time validation for maximum selection
            $('#multiSelect').on('change', function () {
                var selectedOptions = $(this).val(); // Get selected values

                if (selectedOptions && selectedOptions.length > 10) {
                    // Show a warning alert
                    Swal.fire({
                        icon: 'warning',
                        title: 'Selection Limit Exceeded',
                        text: 'You can select a maximum of 10 items.',
                        confirmButtonText: 'OK'
                    });

                    // Automatically deselect the last selected option
                    var lastSelected = selectedOptions[selectedOptions.length - 1]; // Get the last selected option
                    $(this).find('option[value="' + lastSelected + '"]').prop('selected', false); // Deselect it
                    $(this).multiselect('refresh'); // Refresh the dropdown to reflect changes
                }
            });
        });
    </script>
</main>
