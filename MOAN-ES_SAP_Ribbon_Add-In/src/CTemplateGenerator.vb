Imports Microsoft.Office.Interop.Excel
Imports System.Diagnostics

''' <summary>
''' A wrapper class to create the templates for the scripts. 
''' </summary>
''' <remarks></remarks>
Public Class CTemplateGenerator : Implements IDisposable

    Private xlApp As Excel.Application
    Private disposedValue As Boolean ' To detect redundant calls

    Public Sub New()
        Me.xlApp = Globals.ThisAddIn.Application
    End Sub

    ''' <summary>
    ''' Creates the template in Excel. 
    ''' </summary>
    ''' <param name="scriptDict">Information about the script from the script table in the database.</param>
    ''' <remarks></remarks>
    Public Sub initiateTemplate(ByVal scriptDict As Dictionary(Of String, String))

        Dim i As Integer

        ' For when created save dir for reports. 
        '    If Dir(savePath, vbDirectory) = "" Then
        '        MkDir (savePath)
        '    End If

        Dim templateWorkbook As Workbook
        templateWorkbook = xlApp.Workbooks.Add

        With xlApp
            .ScreenUpdating = False
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

            .Rows("2:3").Select()
            With .Selection.Interior
                .Pattern = XlPattern.xlSolid
                .PatternColorIndex = UnclassifiedConstants.xlAutomatic
                .Color = 13311
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
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
            .Range("B2").Select()
            .ActiveCell.FormulaR1C1 = scriptDict("description")
            .Range("B3").Select()
            .ActiveCell.FormulaR1C1 = "Transaction " & scriptDict("transaction")
            .Range("B2:B3").Select()
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
            With .Selection.Font
                .Name = "Segoe UI"
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
            With .Selection.Font
                .ThemeColor = XlThemeColor.xlThemeColorDark1
                .TintAndShade = 0
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
            .ActiveCell.FormulaR1C1 = scriptDict("name")
            With .Selection.Font
                .Color = -16763905
                .TintAndShade = 0
            End With
            .Rows("1:1").RowHeight = 51
        End With

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
            .Selection.Borders(XlBordersIndex.xlEdgeRight).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideVertical).LineStyle = UnclassifiedConstants.xlNone
            .Selection.Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = UnclassifiedConstants.xlNone
            .Cells(1, 1).Select()
        End With

        ' More 

        Dim headers() As String
        headers = Split(scriptDict("headers"), ";")
        i = 0

        Dim validation As String = scriptDict("validation")

        Dim validationSplit() As String = {}
        Dim validationColumn As String = ""
        Dim values() As String = {}
        Dim column As New Integer
        Dim doCreateValidation As Boolean = False

        If (validation.Length > 0) Then
            validationSplit = Split(validation, "=")
            validationColumn = validationSplit(0)
            values = Split(validationSplit(1), ";")
            doCreateValidation = True
        Else

        End If

        With xlApp
            For Each header In headers
                .Cells(5, i + 2).Value2 = header
                If header = validationColumn Then
                    column = i + 2
                End If
                i = i + 1
            Next

            If doCreateValidation Then
                .Cells(6, column).select()
                With .Selection.Validation
                    .Delete()
                    .Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, Operator:= _
                    XlFormatConditionOperator.xlBetween, Formula1:=Join(values, ";"))
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With
            End If

            ' Paste formats for input data.

            .Cells(1, 2).EntireColumn.Copy()
            .Range(.Cells(1, 2), .Cells(1, 2 + i - 1)).EntireColumn.PasteSpecial(XlPasteType.xlPasteFormats)
            If doCreateValidation Then
                .Cells(6, column).copy()
                .Range(.Cells(6, column), .Cells(3000, column)).PasteSpecial(XlPasteType.xlPasteValidation)
            End If
            .Range(.Cells(5, 2), .Cells(5, 2 + i - 1)).Columns.AutoFit()
            .Columns(1).Select()
            .Selection.ColumnWidth = 2.14
            .Cells(6, 2).Select()
            .CutCopyMode = False
            .ScreenUpdating = True

            '  For when created save dir for reports. 
            ' .DisplayAlerts = False
            ' Save the WB for later review by the runner. Need to have save path somewhere. Settings?
            ' .SaveAs(savePath & script & "_" & Format(Now(), "yyyymmddHhNnSs"))
            ' .DisplayAlerts = True

        End With

    End Sub

#Region "IDisposable Support"

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' Dispose managed state (managed objects).
            End If

        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
