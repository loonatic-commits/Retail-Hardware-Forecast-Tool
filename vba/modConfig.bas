Attribute VB_Name = "modConfig"
Option Explicit

' =============================================================
' modConfig - shared constants and color initialization
' =============================================================

' --- Accuracy grading thresholds ---
Public Const ACCURACY_GREEN As Double = 0.8
Public Const ACCURACY_YELLOW As Double = 0.6
Public Const ACCURACY_LOOKBACK_MONTHS As Long = 3

' --- Deviation thresholds ---
Public Const SOF_VARIANCE As Double = 0.2          ' 20% for SOF flagging
Public Const SO_FAR_VARIANCE As Double = 0.05      ' 5% for run-rate flagging

' --- Sheet names ---
Public Const SHEET_CONSENSUS As String = "Consensus"
Public Const SHEET_SOF As String = "SOF"
Public Const SHEET_GRADING_CACHE As String = "_GradingCache"   ' hidden

' --- Legend placement in Consensus sheet (corner) ---
Public Const LEGEND_ANCHOR_CELL As String = "Q1"   ' adjust if Q-column overlaps data

' --- Key figure labels (exact strings as they appear in Column B) ---
' N-1 suffixed labels are prior-period data and are intentionally
' NOT referenced anywhere in the suite.
Public Const KF_BUSINESS_PLAN As String = "Business Plan Quantity"
Public Const KF_SELL_IN As String = "Sell-In Quantity (Gross)"
Public Const KF_BACK_ORDER As String = "Back Order Qty"
Public Const KF_SO_FAR As String = "SO FAR"
Public Const KF_FIELD_FCST_RAW As String = "Field Forecast"   ' not used by any pass
Public Const KF_FIELD_FCST_FINAL As String = "Field Forecast Qty Final"
Public Const KF_MARKETING_PRIMARY As String = "Marketing"      ' not used by any pass
Public Const KF_CONSENSUS As String = "Consensus"              ' not used by any pass
Public Const KF_MARKETING_SECONDARY As String = "Marketing Forecast Qty"
Public Const KF_OPPORTUNITY As String = "Opportunity Qty Final"
Public Const KF_MARKETING_OPP_FCST As String = "Marketing Opportunity Fcst Qty"
Public Const KF_ACTUALS_PLUS As String = "Actuals + Fcst Qty: Marketing"
Public Const KF_AVAIL_INV As String = "Available Inventory (Manual)"

' --- Sentinel for ungradable accuracy ---
Public Const ACCURACY_UNGRADABLE As Double = -1#

' --- Colors (assigned in InitColors since VBA constants can't call RGB) ---
Public CLR_FIELD_GREEN As Long
Public CLR_FIELD_YELLOW As Long
Public CLR_FIELD_RED As Long

Public CLR_FFvsBP_FF_USED As Long
Public CLR_FFvsBP_BP_USED As Long
Public CLR_FFvsBP_BP_FALLBACK As Long

Public CLR_SOF_HIGH As Long
Public CLR_SOF_LOW As Long

Public CLR_SOFAR_UPGRADE As Long
Public CLR_SOFAR_DOWNGRADE As Long

Public CLR_INV_CONSTRAINED_TEXT As Long

Public Sub InitColors()
    CLR_FIELD_GREEN = RGB(198, 239, 206)
    CLR_FIELD_YELLOW = RGB(255, 235, 156)
    CLR_FIELD_RED = RGB(255, 199, 206)

    CLR_FFvsBP_FF_USED = RGB(189, 215, 238)        ' light blue - FF accepted
    CLR_FFvsBP_BP_USED = RGB(226, 207, 245)        ' light purple - BP >= FF
    CLR_FFvsBP_BP_FALLBACK = RGB(255, 217, 217)    ' soft red - FF unreliable, fell back to BP

    CLR_SOF_HIGH = RGB(146, 208, 80)               ' bright green - Pass 2 was > SOF by 20%+ before override
    CLR_SOF_LOW = RGB(244, 176, 132)               ' orange - Pass 2 was < SOF by 20%+ before override

    CLR_SOFAR_UPGRADE = RGB(0, 176, 80)            ' strong green - current month run rate beating forecast
    CLR_SOFAR_DOWNGRADE = RGB(192, 0, 0)           ' strong red fill - current month run rate trailing forecast

    CLR_INV_CONSTRAINED_TEXT = RGB(192, 0, 0)      ' red font when inventory exhausted
End Sub
