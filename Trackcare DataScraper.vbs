Sub GetDataFromWeb()
    ' Phase 1: Logging in

    Dim IE, URL, doc, usernameField, passwordField, table, rows, Row, nextPageButton
    URL = "https://trakcarelabwebview.nhls.ac.za/trakcarelab/csp/system.Home.cls#/Component/SSUser.Logon"

    ' Create Internet Explorer instance
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True ' Set to False if you do not want to see the browser

    ' Navigate to the login page
    IE.Navigate URL
    WaitForIE IE ' Wait for the page to fully load

    ' Access the document
    Set doc = IE.Document

    ' Retrieve username and password fields
    On Error Resume Next
    Set usernameField = doc.getElementById("SSUser_Logon_0-item-USERNAME")
    Set passwordField = doc.getElementById("SSUser_Logon_0-item-PASSWORD")
    On Error GoTo 0

    ' Check if fields are found
    If usernameField Is Nothing Or passwordField Is Nothing Then
        MsgBox "Unable to find the login elements on the page."
        IE.Quit
        Set IE = Nothing
        Exit Sub
    End If

   
    ' Submit the form
    On Error Resume Next
    doc.forms(0).submit
    On Error GoTo 0
    WaitForIE IE ' Wait for the page to fully load after login

    ' Prompt the user to manually apply filters and proceed
    MsgBox "Please apply the filters and click OK to continue."

    ' Phase 2: Retrieving Data and Writing to Excel

    ' Create a new Excel application
    Dim xlApp, xlBook, xlSheet
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    ' Initialize the row index for writing data to Excel
    Dim rowIndex
    rowIndex = 1

    ' Loop to handle pagination
    Do
        ' Refresh document reference
        Set doc = IE.Document

        ' Get the table element
        Set table = doc.getElementById("tweb_DEBDebtor_FindList_0")

        ' Check if the table is found
        If table Is Nothing Then
            MsgBox "Unable to find the table element on the page."
            IE.Quit
            Set IE = Nothing
            Exit Sub
        End If

        ' Get the table rows
        Set rows = table.getElementsByTagName("tr")

        ' Write the table data into Excel
        Dim colIndex
        For Each Row In rows
            colIndex = 1

            ' Ensure the row has the expected number of cells
            If Row.cells.Length >= 8 Then
                ' Write data to Excel from the specified columns
                xlSheet.Cells(rowIndex, 1).Value = Row.cells(0).innerText ' Surname
                xlSheet.Cells(rowIndex, 2).Value = Row.cells(1).innerText ' GivenName
                xlSheet.Cells(rowIndex, 3).Value = Row.cells(2).innerText ' MRN
                xlSheet.Cells(rowIndex, 4).Value = Row.cells(3).innerText ' MRNLink
                xlSheet.Cells(rowIndex, 5).Value = Row.cells(4).innerText ' DOB
                xlSheet.Cells(rowIndex, 6).Value = Row.cells(5).innerText ' Species
                xlSheet.Cells(rowIndex, 7).Value = Row.cells(6).innerText ' HospitalURNo
                xlSheet.Cells(rowIndex, 8).Value = Row.cells(7).innerText ' UserLocation
            End If

            ' Increment the row index
            rowIndex = rowIndex + 1
        Next

        ' Check for next page button
        Set nextPageButton = Nothing
        On Error Resume Next
        Set nextPageButton = doc.querySelector("a.ng-binding.ng-scope[id*='nextPage']")
        On Error GoTo 0

        If nextPageButton Is Nothing Then
            Exit Do ' Exit the loop if there's no next page button
        Else
            nextPageButton.Click ' Click the next page button
            WaitForIE IE ' Wait for the page to fully load
        End If
    Loop

    ' Close Internet Explorer
    IE.Quit
    Set IE = Nothing

    ' Inform the user that the process is complete
    MsgBox "Data retrieval complete. Internet Explorer closed."
End Sub

Sub WaitForIE(IE)
    Do While IE.Busy Or IE.ReadyState <> 4
        WScript.Sleep 100
    Loop
    ' Additional wait to ensure elements are loaded
    WScript.Sleep 500
End Sub

' Call the main subroutine
GetDataFromWeb

