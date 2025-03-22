Attribute VB_Name = "Export_Data_to_Ppt_Code"
Sub GenerateCompanyPresentations()
    Dim pptApp As Object, pptPres As Object, pptSlide As Object
    Dim pptTemplate As String, pptNewFile As String
    Dim xlSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim companyName As String, prevCompany As String
    Dim xlData As Object
    Dim key As Variant
    Dim shp As Object
    Dim newSlide As Object
    
    ' Set Excel Data Sheet
    Set xlSheet = ThisWorkbook.Sheets(1)
    
    ' Determine last row with data
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(xlUp).Row
    
    ' PowerPoint Template Path
    pptTemplate = ThisWorkbook.Path & "\Sample_Presentation.pptx"
    
    ' Start PowerPoint Application
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    On Error GoTo 0
    
    ' Initialize previous company tracker
    prevCompany = ""
    
    ' Loop through data
    For i = 2 To lastRow
        companyName = xlSheet.Cells(i, 1).Value  ' Read company name
        
        ' Check if a new company is starting, create a new PowerPoint file
        If companyName <> prevCompany Then
            If Not pptPres Is Nothing Then
                ' Delete Slide 2 before saving the file
                If pptPres.Slides.Count >= 2 Then
                    pptPres.Slides(2).Delete
                End If
                
                pptPres.Save
                pptPres.Close
            End If
            
            ' Save new file with company name
            pptNewFile = ThisWorkbook.Path & "\" & companyName & "_Report.pptx"
            FileCopy pptTemplate, pptNewFile
            
            ' Open the new PowerPoint file
            Set pptPres = pptApp.Presentations.Open(pptNewFile)
            
            ' Update placeholders in the first slide (Company Summary)
            Set pptSlide = pptPres.Slides(1)
            For Each shp In pptSlide.Shapes
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, "[Company Name]", companyName)
                    End If
                End If
            Next shp
        End If
        
        ' Store row data for replacements
        Set xlData = CreateObject("Scripting.Dictionary")
        xlData("[Company Name]") = companyName
        xlData("[Report Period]") = xlSheet.Cells(i, 2).Value
        xlData("[Revenue]") = Format(xlSheet.Cells(i, 3).Value, "#,##0")
        xlData("[Expenses]") = Format(xlSheet.Cells(i, 4).Value, "#,##0")
        xlData("[Net Profit]") = Format(xlSheet.Cells(i, 5).Value, "#,##0")
        
        ' Duplicate the financial summary slide (Slide 2)
        pptPres.Slides(2).Copy
        pptPres.Slides.Paste
        Set newSlide = pptPres.Slides(pptPres.Slides.Count) ' Get newly pasted slide
        
        ' Replace placeholders in the duplicated slide
        For Each shp In newSlide.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    For Each key In xlData.keys
                        If InStr(1, shp.TextFrame.TextRange.Text, key, vbTextCompare) > 0 Then
                            shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, key, xlData(key))
                        End If
                    Next key
                End If
            End If
        Next shp
        
        ' Update previous company tracker
        prevCompany = companyName
    Next i

    ' Final cleanup for last company
    If Not pptPres Is Nothing Then
        ' Delete Slide 2 for the last company's presentation
        If pptPres.Slides.Count >= 2 Then
            pptPres.Slides(2).Delete
        End If
        
        pptPres.Save
        pptPres.Close
    End If
    
    ' Cleanup
    pptApp.Quit
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub


