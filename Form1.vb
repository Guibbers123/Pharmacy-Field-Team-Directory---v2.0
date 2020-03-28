Imports Microsoft.Office.Interop
Public Class Form1
    Public MyArray As Object(,), lastRow As ULong, storeNum As UInteger
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim file = "", strFile, strPath, fileName As String, fileNum As UInteger
        Dim fileFound As Boolean
        Dim xlsWorkbooks As Excel.Workbooks = Nothing
        Dim appPath As String = Application.StartupPath
        Dim sheetName As String = ""
        Dim xls As Excel.Application
        Dim workbook As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim rng As Excel.Range

        Me.Width = 425

        xls = New Excel.Application
        xls.Visible = False
        xls.EnableEvents = False
        xls.DisplayAlerts = False

        fileName = "premises"
        sheetName = "Boots UK Ltd"
        fileFound = False
        strPath = appPath & "\"
        file = Dir(strPath)
        Do While file <> ""
            If InStr(1, UCase(file), UCase(fileName)) Then
                strFile = file
                fileFound = True
                fileNum = fileNum + 1
                If fileFound = True Then
                    workbook = xls.Workbooks.Open(strPath & file, [ReadOnly]:=True)
                    ws = workbook.Worksheets(sheetName)
                    rng = ws.UsedRange
                    MyArray = CType(rng.Value, Object(,))
                    workbook.Close()
                Else
                    MsgBox("No files found")
                End If
            End If
            file = Dir()
        Loop
        xls.DisplayAlerts = True
        xls.EnableEvents = True
        If fileFound = False Then
            MsgBox("Master Premises file is missing")
            Me.Close()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Erase MyArray
        Me.Close()
    End Sub
    Dim objapp As Excel.Application
    Dim objbook As Excel.Workbook
    Sub CheckForExcelFiles()
        Dim storeName As String
        Dim regionName, areaName, pharmacy, registration As String
        Dim address0, address1, address2, address3, address4, address5, postcode As String
        Dim country, phone, LAT, NHS, NHShours, ODS, CS, houseNumber As String
        Dim regionNumber, areaNumber As UInteger
        Dim x As ULong
        Dim xlsWorkbooks As Excel.Workbooks = Nothing
        Dim appPath As String = Application.StartupPath
        Dim sheetName As String = ""
        Dim xls As Excel.Application

        lastRow = UBound(MyArray, 1)
        If IsNumeric(TextBox1.Text) Then storeNum = CInt(TextBox1.Text)
        For x = 2 To lastRow
            If CInt(MyArray(x, 1)) = storeNum Then
                storeName = StrConv(MyArray(x, 2), vbProperCase)
                regionNumber = MyArray(x, 3)
                regionName = StrConv(MyArray(x, 4), vbProperCase)
                areaNumber = MyArray(x, 5)
                areaName = StrConv(MyArray(x, 6), vbProperCase)
                pharmacy = MyArray(x, 7)
                registration = MyArray(x, 9)
                country = MyArray(x, 18)
                phone = MyArray(x, 19)
                LAT = MyArray(x, 20)
                NHS = MyArray(x, 21)
                NHShours = MyArray(x, 22)
                ODS = MyArray(x, 23)
                CS = MyArray(x, 25)
                houseNumber = MyArray(x, 11)
                address0 = StrConv(houseNumber, vbProperCase)
                address1 = StrConv(MyArray(x, 12), vbProperCase)
                address2 = StrConv(MyArray(x, 13), vbProperCase)
                address3 = StrConv(MyArray(x, 14), vbProperCase)
                address4 = StrConv(MyArray(x, 15), vbProperCase)
                address5 = StrConv(MyArray(x, 16), vbProperCase)
                postcode = MyArray(x, 17)
                If address0 <> "" Then
                    TextBox13.Text = address0 & " "
                End If
                If address1 <> "" Then
                    TextBox13.Text = TextBox13.Text & address1 & vbCrLf
                End If
                If address2 <> "" Then
                    TextBox13.Text = TextBox13.Text & address2 & vbCrLf
                End If
                If address3 <> "" Then
                    TextBox13.Text = TextBox13.Text & address3 & vbCrLf
                End If
                If address4 <> "" Then
                    TextBox13.Text = TextBox13.Text & address4 & vbCrLf
                End If
                If address5 <> "" Then
                    TextBox13.Text = TextBox13.Text & address5 & vbCrLf
                End If

                TextBox13.Text = TextBox13.Text & postcode
                RemoveHandler TextBox2.TextChanged, AddressOf TextBox2_TextChanged
                TextBox2.Text = storeName
                TextBox3.Text = regionNumber
                TextBox4.Text = regionName
                TextBox5.Text = areaNumber
                TextBox6.Text = areaName
                TextBox7.Text = pharmacy
                TextBox8.Text = registration
                TextBox9.Text = NHS
                TextBox10.Text = NHShours
                TextBox11.Text = ODS
                TextBox12.Text = CS
                TextBox14.Text = phone

                AddHandler TextBox2.TextChanged, AddressOf TextBox2_TextChanged
                Exit For
            End If
        Next
    End Sub
    'Keep the application object and the workbook object global, so you can  
    'retrieve the data in Button2_Click that was set in Button1_Click.
    'Dim objApp As Excel.Application
    'Dim objBook As Excel._Workbook
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        CheckForExcelFiles()
    End Sub
    Private Sub TextBox1_keydown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            CheckForExcelFiles()
        End If
    End Sub
    Private Sub TextBox2_keydown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Down Then
            'set focus to listBox
            ListBox1.Select()
            ListBox1.Select()
            ListBox1.SelectedItem = 1
            ListBox1.SelectedItem = 1
        End If
    End Sub

    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As EventArgs) Handles ListBox1.MouseDoubleClick
        'Exit Sub
        Dim thisSelection, sbStr As String
        thisSelection = ListBox1.SelectedItem.ToString
        sbStr = thisSelection.Substring(0, 4)
        storeNum = CInt(sbStr)
        TextBox1.Text = storeNum
        ListBox1.Visible = False
        CheckForExcelFiles()
    End Sub
    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        Dim thisSelection, sbStr As String
        If e.KeyCode = 13 Then
            thisSelection = ListBox1.SelectedItem.ToString
            sbStr = thisSelection.Substring(0, 4)
            storeNum = CInt(sbStr)
            TextBox1.Text = storeNum
            ListBox1.Visible = False
            CheckForExcelFiles()
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim stName As String, foundList() As Object
        stName = TextBox2.Text
        If Len(stName) > 2 Then
            ListBox1.Visible = True
            foundList = obtainShortList(stName)
            ListBox1.Items.Clear()
            ListBox1.Items.AddRange(foundList)
        End If
    End Sub
    Private Sub TextBox1_GotFocus() Handles TextBox1.GotFocus
        TextBox1.Text = ""
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.Width <> 850 Then
            Me.Width = 850
            Button2.Text = "< Hide people details  "
        Else
            Me.Width = 425
            Button2.Text = "  Show people details >"
        End If
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub

    Private Sub TextBox2_GotFocus() Handles TextBox2.GotFocus
        TextBox2.Text = ""
    End Sub
    Function obtainShortList(stName) As Object
        Dim shortListArray(1000) As Object, x As ULong, y As ULong, stNum As String
        Dim mPadding As String
        y = 0
        lastRow = UBound(MyArray, 1)
        For x = 2 To lastRow
            If InStr(UCase(MyArray(x, 2)), UCase(stName)) Then
                stNum = MyArray(x, 1)
                mPadding = Convert.ToChar("0")
                stNum = stNum.PadLeft(4, mPadding)
                shortListArray(y) = stNum & " " & MyArray(x, 2)
                y = y + 1
                End If
        Next x
        ReDim Preserve shortListArray(y - 1)

        Return shortListArray
        'Erase shortListArray()
    End Function
End Class