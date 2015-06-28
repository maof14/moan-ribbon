Imports Microsoft.Office.Interop.Excel
Imports System.Diagnostics

' Class to generate the templates to be filled in by the user. In VS for easier handling an encapsulation from the user. 

Public Class CTemplateGenerator

    Private xlApp As Excel.Application

    Public Sub New()
        xlApp = Globals.ThisAddIn.Application
    End Sub

    ' Function to generate the template in Excel. 
    ' Todo: Have a save path somewhere in the project. 
    ' Return void. 
    Public Sub InitiateTemplate(ByVal scriptDict As Dictionary(Of String, String))

        Dim name As String = scriptDict("name")
        Dim scriptid As String = scriptDict("scriptid")
        Dim description As String = scriptDict("description")

        Dim i As Integer
        Dim username As String
        username = Environ$("username")

        '    If Dir(savePath, vbDirectory) = "" Then
        '        MkDir (savePath)
        '    End If

        Dim templateWorkbook As Workbook
        templateWorkbook = xlApp.Workbooks.Add

        ' Hopefully creates in the added workbook... 
        With xlApp
            .Cells.Select()
            With .Selection.Interior
                .Pattern = XlPattern.xlSolid
                .PatternColorIndex = UnclassifiedConstants.xlAutomatic
                .ThemeColor = 1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Selection.Font
                .Name = "Trebuchet MS"
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            With .Selection.Font
                .Name = "Trebuchet MS"
                .Size = 10
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            .Rows("1:1").Select()
            With .Selection.Font
                .Name = "Trebuchet MS"
                .Size = 28
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            .Rows("5:5").Select()
            With .Selection.Font
                .Name = "Trebuchet MS"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            .Range("B5").Select()
            .Range("B5").Select()
            With .Selection.Font
                .Name = "Trebuchet MS"
                .FontStyle = "Normal"
                .Bold = True
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = XlThemeColor.xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            With .Selection.Interior
                .Pattern = XlPattern.xlSolid
                .PatternColorIndex = UnclassifiedConstants.xlAutomatic
                .Color = 13311
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .Selection.Borders(XlBordersIndex.xlDiagonalDown).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlDiagonalUp).LineStyle = UnclassifiedConstants.xlNone
            With .Selection.Borders(XlBordersIndex.xlEdgeLeft)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 15
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Selection.Borders(XlBordersIndex.xlEdgeTop)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 15
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            With .Selection.Borders(XlBordersIndex.xlEdgeRight)
                .LineStyle = XlLineStyle.xlContinuous
                .ColorIndex = 15
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = UnclassifiedConstants.xlNone
            .Range("B1").Select()
            .ActiveCell.FormulaR1C1 = "MOAN Script template"
            .Range("B1").Select()
            With .Selection.Font
                .Color = -16763905
                .TintAndShade = 0
            End With
            .Rows("1:1").RowHeight = 51
        End With

        ' Second part - Create the custom headers from the database. 

        With xlApp
            For i = 6 To 7
                .Cells(i, 2).Select()
                With .Selection
                    With .Font
                        .ThemeColor = XlThemeColor.xlThemeColorLight1
                        .TintAndShade = 0.349986267
                        .ThemeFont = XlThemeFont.xlThemeFontNone
                    End With
                    With .Borders(XlBordersIndex.xlEdgeRight)
                        .LineStyle = XlLineStyle.xlContinuous
                        .ColorIndex = 2
                        .TintAndShade = 0
                        .Weight = XlBorderWeight.xlThin
                    End With
                    With .Borders(XlBordersIndex.xlEdgeLeft)
                        .LineStyle = XlLineStyle.xlContinuous
                        .ColorIndex = 2
                        .TintAndShade = 0
                        .Weight = XlBorderWeight.xlThin
                    End With
                    If i = 6 Then
                        With .Interior
                            .Pattern = XlPattern.xlSolid
                            .PatternColorIndex = UnclassifiedConstants.xlAutomatic
                            .ThemeColor = XlThemeColor.xlThemeColorDark2
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                End With
            Next i

            .Range("B6:B7").Select()
            .Selection.Copy()

            .Range("B6:B3000").PasteSpecial(XlPasteType.xlPasteFormats)

            .Range("G2:K2").Select()
            .Selection.Merge()

            With .Selection
                .HorizontalAlignment = UnclassifiedConstants.xlLeft
                .VerticalAlignment = UnclassifiedConstants.xlBottom
                .WrapText = False
                With .Font
                    .Size = 10
                    .Bold = True
                    .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                    .Color = -16763905
                    .TintAndShade = 0
                    .ThemeFont = XlThemeFont.xlThemeFontNone
                End With
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = UnclassifiedConstants.xlContext
                .MergeCells = True
            End With

            .Selection.Borders(XlBordersIndex.xlDiagonalDown).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlDiagonalUp).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlEdgeTop).LineStyle = UnclassifiedConstants.xlNone
            With .Selection.Borders(XlBordersIndex.xlEdgeBottom)
                .LineStyle = XlLineStyle.xlDouble
                .ThemeColor = 1
                .TintAndShade = -0.14996795556505
                .Weight = XlBorderWeight.xlThick
            End With
            .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = UnclassifiedConstants.xlNone

            .Selection.Value = scriptid

            .Cells(1, 1).Select()

            With .Selection.Font
                .ThemeColor = XlThemeColor.xlThemeColorDark1
                .TintAndShade = 0
            End With
            With .Selection.Interior
                .Pattern = XlPattern.xlSolid
                .PatternColorIndex = UnclassifiedConstants.xlAutomatic
                .ThemeColor = XlThemeColor.xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

            '    If (Not IsMissing(transactionString)) Then
            '        transaction = transactionString
            '    End If

            .Rows("1:1").Select()
            With .Selection.Font
                .Name = "Segoe UI Light"
                .FontStyle = "Regular"
                .Size = 28
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .Color = 13311
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            With .Selection.Interior
                .Pattern = XlPattern.xlSolid
                .PatternColorIndex = UnclassifiedConstants.xlAutomatic
                .Color = 13311
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .Range("B1").Select()
            With .Selection.Font
                .Name = "Segoe UI Light"
                .FontStyle = "Regular"
                .Size = 28
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = XlThemeColor.xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With
            .Rows("1:1").Select()
            .Selection.Borders(XlBordersIndex.xlDiagonalDown).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlDiagonalUp).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlEdgeLeft).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlEdgeTop).LineStyle = UnclassifiedConstants.xlNone
            With .Selection.Borders(XlBordersIndex.xlEdgeBottom)
                .LineStyle = XlLineStyle.xlContinuous
                .Color = -13395457
                .TintAndShade = 0
                .Weight = XlBorderWeight.xlThin
            End With
            .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = UnclassifiedConstants.xlNone

            ' Sväv med VS här.

            'For i = 5 To 5 + UBound(headers)
            '    .Cells(5, i - 3).Value = headers(i - 5)
            'Next i

            .Cells(1, 1).Select()
        End With

        ' More 

        Dim headers() As String
        headers = Split(scriptDict("headers"), ";")
        i = 0
        With xlApp
            For Each header In headers
                .Cells(5, i + 2).Value2 = header
                i = i + 1
            Next

            ' Description cell.
            .Range("G3:K4").Select()

            .Selection.Merge()

            With .Selection
                .WrapText = True
                .HorizontalAlignment = UnclassifiedConstants.xlLeft
                .VerticalAlignment = UnclassifiedConstants.xlTop
                .HorizontalAlignment = UnclassifiedConstants.xlLeft
                .VerticalAlignment = UnclassifiedConstants.xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = UnclassifiedConstants.xlContext
                .MergeCells = True
            End With

            With .Selection.Font
                .Name = "Trebuchet MS"
                .Size = 7
                .Bold = False
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = XlUnderlineStyle.xlUnderlineStyleNone
                .Color = -16763905
                .TintAndShade = 0
                .ThemeFont = XlThemeFont.xlThemeFontNone
            End With

            .Selection.value = description

            .ScreenUpdating = False
            .Cells(1, 2).EntireColumn.Copy()
            .Range(.Cells(1, 2), .Cells(1, 2 + i - 1)).EntireColumn.PasteSpecial(XlPasteType.xlPasteFormats)
            .Range(.Cells(5, 2), .Cells(5, 2 + i - 1)).Columns.AutoFit()
            .Columns(1).Select()
            .Selection.ColumnWidth = 2.14
            .Cells(6, 2).Select()
            .CutCopyMode = False
            .ScreenUpdating = True

            .DisplayAlerts = False
            ' Save the WB for later review by the runner. Need to have save path somewhere. Settings?
            ' .SaveAs(savePath & script & "_" & Format(Now(), "yyyymmddHhNnSs"))
            .DisplayAlerts = True

        End With

    End Sub



End Class
