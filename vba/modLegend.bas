Attribute VB_Name = "modLegend"
Option Explicit

' =============================================================
' Pass 6 - Legend
'
' Writes a two-column self-documenting color legend in a corner
' of the Consensus sheet, anchored at LEGEND_ANCHOR_CELL.
' Column 1: bordered swatch with the relevant fill.
' Column 2: plain-English description.
'
' All colors come from modConfig, so the legend stays in sync
' with the source of truth.
' =============================================================

Public Sub Run_Module_6_Legend()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim anchor As Range
    Set anchor = ws.Range(LEGEND_ANCHOR_CELL)

    ' Clear a generous block so stale entries don't linger.
    Dim clearRng As Range
    Set clearRng = ws.Range(anchor, anchor.Offset(40, 1))
    clearRng.Clear

    Dim r As Long: r = 0

    WriteHeader ws, anchor.Offset(r, 0), "Forecasting Suite - Color Legend": r = r + 2

    WriteSection ws, anchor.Offset(r, 0), "Field Forecast accuracy (FF Qty Final label cell)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_FIELD_GREEN, "Accuracy >= 80% (green)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_FIELD_YELLOW, "Accuracy 60-80% (yellow)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_FIELD_RED, "Accuracy < 60% (red)": r = r + 2

    WriteSection ws, anchor.Offset(r, 0), "FF vs BP seed (Opportunity row)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_FFvsBP_FF_USED, "FF used (FF > BP and FF is reliable)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_FFvsBP_BP_USED, "BP used (BP >= FF)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_FFvsBP_BP_FALLBACK, "BP fallback (FF > BP but FF unreliable)": r = r + 2

    WriteSection ws, anchor.Offset(r, 0), "SOF override deviation (next 3 months)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_SOF_HIGH, "Pass 2 was 20%+ above SOF before override": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_SOF_LOW, "Pass 2 was 20%+ below SOF before override": r = r + 2

    WriteSection ws, anchor.Offset(r, 0), "SO FAR run rate (current month)": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_SOFAR_UPGRADE, "Tracking 5%+ ahead of prior forecast": r = r + 1
    WriteSwatch ws, anchor.Offset(r, 0), CLR_SOFAR_DOWNGRADE, "Tracking 5%+ behind prior forecast": r = r + 2

    WriteSection ws, anchor.Offset(r, 0), "Inventory constraint (font color)": r = r + 1
    WriteFontSwatch ws, anchor.Offset(r, 0), CLR_INV_CONSTRAINED_TEXT, "Cumulative Opportunity exceeds Available Inventory"

    ' Column widths so the legend is readable
    anchor.EntireColumn.ColumnWidth = 6
    anchor.Offset(0, 1).EntireColumn.ColumnWidth = 60

    Exit Sub

Fail:
    Err.Source = "Module 6 - Legend"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub WriteHeader(ws As Worksheet, anchor As Range, text As String)
    anchor.Value = text
    anchor.Font.Bold = True
    anchor.Font.Size = 12
End Sub

Private Sub WriteSection(ws As Worksheet, anchor As Range, text As String)
    anchor.Value = text
    anchor.Font.Bold = True
End Sub

Private Sub WriteSwatch(ws As Worksheet, anchor As Range, colorVal As Long, description As String)
    anchor.Value = ""
    anchor.Interior.Color = colorVal
    anchor.Borders.LineStyle = xlContinuous
    anchor.Borders.Weight = xlThin
    anchor.Offset(0, 1).Value = description
    anchor.Offset(0, 1).HorizontalAlignment = xlLeft
End Sub

Private Sub WriteFontSwatch(ws As Worksheet, anchor As Range, colorVal As Long, description As String)
    anchor.Value = "Abc"
    anchor.Font.Color = colorVal
    anchor.Font.Bold = True
    anchor.HorizontalAlignment = xlCenter
    anchor.Borders.LineStyle = xlContinuous
    anchor.Borders.Weight = xlThin
    anchor.Offset(0, 1).Value = description
    anchor.Offset(0, 1).HorizontalAlignment = xlLeft
End Sub
