Attribute VB_Name = "ExternalLinkFinder"
Sub FindLinks()
' =================================================================================================
' Checks for any external links that give the following errors when a workbook is opened:
'   "This workbook contains links to other data sources", or
'   "This workbook contains one or more links that cannot be updated"
'
' Any links that are found are documented in a new sheet
' =================================================================================================
Application.ScreenUpdating = False

    Dim ListSheet As Worksheet
    Dim r As Integer
    Dim Sht As Worksheet
    Dim Box As Range
    Dim clForm As FormatCondition
    Dim ValCells As Range
    Dim Shp As Shape
    Dim PivotTbl As PivotTable
    Dim Tbl As ListObject
    
    ' ---------------------------------------------------------------------------------------------
    ' Setup the sheet for listing all external links
    ' ---------------------------------------------------------------------------------------------
    Sheets.Add After:=Sheets(Sheets.Count)
    Set ListSheet = Sheets(Sheets.Count)
    With ListSheet
        .Move Before:=Sheets(1)
        .Cells(1, 1).Value = "Type"
        .Cells(1, 2).Value = "Name"
        .Cells(1, 3).Value = "Location"
        .Cells(1, 4).Value = "Offender"
        .Cells(1, 5).Value = "Value"
    End With
    
    ' ---------------------------------------------------------------------------------------------
    ' Loop through each sheet, finding all external links
    ' ---------------------------------------------------------------------------------------------
    r = 2
    For Each Sht In ActiveWorkbook.Worksheets
        If Sht.Name <> ListSheet.Name Then
            
            ' Cell formulas and conditional formatting
            For Each Box In Sht.UsedRange
                
                ' Cell formulas
                If InStr(Box.Formula, "[") > 0 And InStr(Box.Formula, "]") > 0 Then
                    ListSheet.Cells(r, 1) = "Cell"
                    ListSheet.Cells(r, 3) = Sht.Name & ", " & Box.Address
                    ListSheet.Cells(r, 4) = "Formula"
                    ListSheet.Cells(r, 5) = "'" & Box.Formula
                    r = r + 1
                End If
                
                ' Conditional formatting formulas
                For Each clForm In Box.FormatConditions
                    If InStr(clForm.Formula1, "[") > 0 And InStr(clForm.Formula1, "]") > 0 Then
                        ListSheet.Cells(r, 1) = "Cell"
                        ListSheet.Cells(r, 3) = Sht.Name & ", " & Box.Address
                        ListSheet.Cells(r, 4) = "Conditional Formatting"
                        ListSheet.Cells(r, 5) = "'" & clForm.Formula1
                        r = r + 1
                    End If
                Next
            Next

            ' Cell validations
            On Error Resume Next
            Err = 0
            Set ValCells = Intersect(Sht.UsedRange.SpecialCells(xlCellTypeAllValidation), Sht.UsedRange)
            If Err = 0 Then
                For Each Box In ValCells
                    If InStr(Box.Validation.Formula1, "[") > 0 And InStr(Box.Validation.Formula1, "]") > 0 Then
                        ListSheet.Cells(r, 1) = "Cell"
                        ListSheet.Cells(r, 3) = Sht.Name & ", " & Box.Address
                        ListSheet.Cells(r, 4) = "Validation"
                        ListSheet.Cells(r, 5) = "'" & Box.Validation.Formula1
                        r = r + 1
                    End If
                Next
            End If

            ' Shapes
            For Each Shp In Sht.Shapes
        
                ' Chart formulas
                Err = 0
                ErrCheck = Shp.Chart.SeriesCollection(1).Formula
                If Err = 0 Then
                    For Each DataSeries In Shp.Chart.SeriesCollection
                        If InStr(DataSeries.Formula, "[") > 0 And InStr(DataSeries.Formula, "]") > 0 Then
                            ListSheet.Cells(r, 1) = "Shape"
                            ListSheet.Cells(r, 2) = Shp.Name
                            ListSheet.Cells(r, 3) = Sht.Name & ", " & Shp.TopLeftCell.Address
                            ListSheet.Cells(r, 4) = DataSeries.Name
                            ListSheet.Cells(r, 5) = "'" & DataSeries.Formula
                            r = r + 1
                        End If
                    Next
                End If
        
                ' Shape formulas
                Err = 0
                ErrCheck = Shp.DrawingObject.Formula
                If Err = 0 Then
                    If InStr(Shp.DrawingObject.Formula, "[") > 0 And InStr(Shp.DrawingObject.Formula, "]") > 0 Then
                        ListSheet.Cells(r, 1) = "Shape"
                        ListSheet.Cells(r, 2) = Shp.Name
                        ListSheet.Cells(r, 3) = Sht.Name & ", " & Shp.TopLeftCell.Address
                        ListSheet.Cells(r, 4) = "Formula"
                        ListSheet.Cells(r, 5) = "'" & Shp.DrawingObject.Formula
                        r = r + 1
                    End If
                End If
        
                ' Shape Macros
                Err = 0
                ErrCheck = Shp.OnAction
                If Err = 0 Then
                    If Len(Shp.OnAction) > 0 And InStr(Shp.OnAction, ActiveWorkbook.Name) = 0 Then
                        ListSheet.Cells(r, 1) = "Shape"
                        ListSheet.Cells(r, 2) = Shp.Name
                        ListSheet.Cells(r, 3) = Sht.Name & ", " & Shp.TopLeftCell.Address
                        ListSheet.Cells(r, 4) = "Assigned Macro"
                        ListSheet.Cells(r, 5) = "'" & Shp.OnAction
                        r = r + 1
                    End If
                End If
        
                ' Form Control Input Ranges
                Err = 0
                ErrCheck = Shp.ControlFormat.ListFillRange
                If Err = 0 Then
                    If InStr(Shp.ControlFormat.ListFillRange, "[") > 0 And InStr(Shp.ControlFormat.ListFillRange, "]") > 0 Then
                        ListSheet.Cells(r, 1) = "Form Control"
                        ListSheet.Cells(r, 2) = Shp.Name
                        ListSheet.Cells(r, 3) = Sht.Name & ", " & Shp.TopLeftCell.Address
                        ListSheet.Cells(r, 4) = "Input Range"
                        ListSheet.Cells(r, 5) = "'" & Shp.ControlFormat.ListFillRange
                        r = r + 1
                    End If
                End If
        
                ' Form Control Linked Cells
                Err = 0
                ErrCheck = Shp.ControlFormat.LinkedCell
                If Err = 0 Then
                    If InStr(Shp.ControlFormat.LinkedCell, "[") > 0 And InStr(Shp.ControlFormat.LinkedCell, "]") > 0 Then
                        ListSheet.Cells(r, 1) = "Form Control"
                        ListSheet.Cells(r, 2) = Shp.Name
                        ListSheet.Cells(r, 3) = Sht.Name & ", " & Shp.TopLeftCell.Address
                        ListSheet.Cells(r, 4) = "Linked Cell"
                        ListSheet.Cells(r, 5) = "'" & Shp.ControlFormat.LinkedCell
                        r = r + 1
                    End If
                End If
            Next
        
            ' Pivot tables
            For Each PivotTbl In Sht.PivotTables
                Err = 0
                ErrCheck = Len(PivotTbl.PivotCache.SourceDataFile)
                If Err = 0 Then
                    ListSheet.Cells(r, 1) = "Pivot Table"
                    ListSheet.Cells(r, 2) = PivotTbl.Name
                    ListSheet.Cells(r, 3) = Sht.Name & ", " & PivotTbl.TableRange2.Cells(1, 1).Address
                    ListSheet.Cells(r, 4) = "SourceData"
                    ListSheet.Cells(r, 5) = "'" & PivotTbl.PivotCache.SourceDataFile
                    r = r + 1
                End If
            Next
            
            ' Regular Tables
            For Each Tbl In Sht.ListObjects
                Err = 0
                ErrCheck = Tbl.QueryTable.SourceDataFile
                If Err = 0 Then
                    ListSheet.Cells(r, 1) = "Table"
                    ListSheet.Cells(r, 2) = Tbl.Name
                    ListSheet.Cells(r, 3) = Sht.Name & ", " & Tbl.Address
                    ListSheet.Cells(r, 4) = "SourceData"
                    ListSheet.Cells(r, 5) = "'" & Tbl.QueryTable.SourceDataFile
                    r = r + 1
                End If
            Next
            On Error GoTo 0
            
        End If
    Next
    
    ' Named Ranges
    For Each NamedRange In ActiveWorkbook.Names
        If InStr(NamedRange.RefersTo, "[") > 0 And InStr(NamedRange.RefersTo, "]") > 0 Then
            ListSheet.Cells(r, 1) = "Named Range"
            ListSheet.Cells(r, 2) = NamedRange.Name
            ListSheet.Cells(r, 4) = "Reference"
            ListSheet.Cells(r, 5) = "'" & NamedRange.RefersTo
            r = r + 1
        End If
    Next

    ListSheet.Select
    ListSheet.Cells.EntireColumn.AutoFit
    ListSheet.Range("A1").Select

Application.ScreenUpdating = True
End Sub
