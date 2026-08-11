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
Public Const SO_FAR_VARIANCE As Double = 0.5       ' 50% for current-month run-rate flagging
Public Const VARIANCE_DISSONANCE As Double = 0.15  ' 15% dissonance-flag threshold

' --- Forecast selection ---
' The constrained test compares availability against a demand signal.
' Demand signal = max(Consensus N-1, Field N-1) when this is True.
' Set to False to use max(Consensus N-1, Field) - the current-month
' field submission - instead.
Public Const CONSTRAINT_USE_FIELD_N1 As Boolean = True

' Month offsets, relative to the current month (N = 0)
Public Const OFFSET_AVAILABILITY_LAST As Long = 3   ' N..N+3 = availability when constrained
Public Const OFFSET_BACKORDER As Long = 4           ' N+4 = availability + backorder
Public Const OFFSET_CONSENSUS_HOLD As Long = 3      ' N+3 onward = Consensus N-1 (normal path)

' --- Info block placement ---
' Info block starts this many columns right of the last month column.
' 2 leaves one blank spacer column between the data and the block.
Public Const INFO_BLOCK_COL_OFFSET As Long = 2

' --- Sheet names ---
Public Const SHEET_CONSENSUS As String = "Consensus"
Public Const SHEET_SOF As String = "SOF"
Public Const SHEET_GRADING_CACHE As String = "_GradingCache"

' --- Key figure labels (exact strings as they appear in Column B) ---
Public Const KF_BUSINESS_PLAN As String = "Business Plan Quantity"
Public Const KF_SELL_IN As String = "Sell-In Quantity (Gross)"
Public Const KF_BACK_ORDER As String = "Back Order Qty"
Public Const KF_SO_FAR As String = "SO FAR"
Public Const KF_FIELD_FCST_FINAL As String = "Field Forecast Qty Final"
Public Const KF_FIELD_FCST_N1 As String = "Field Forecast Qty N-1"
Public Const KF_CONSENSUS_N1 As String = "Consensus Fcst Qty Final N-1"
Public Const KF_MARKETING_SECONDARY As String = "Marketing Forecast Qty"
Public Const KF_OPPORTUNITY As String = "Opportunity Qty Final"
Public Const KF_MARKETING_OPP_FCST As String = "Marketing Opportunity Fcst Qty"
Public Const KF_ACTUALS_PLUS As String = "Actuals + Fcst Qty: Marketing"
Public Const KF_AVAIL_INV As String = "Available Inventory (Manual)"

' --- Sentinel for ungradable accuracy ---
Public Const ACCURACY_UNGRADABLE As Double = -1#

' --- Colors ---
' Surviving color coding after the color-module removal:
'   Pass 1  field accuracy grade on the FF Qty Final label cell
'   Pass 5  red font on inventory-constrained Opportunity cells
'   Info    dissonance fill on flagged Opportunity month cells
Public CLR_FIELD_GREEN As Long
Public CLR_FIELD_YELLOW As Long
Public CLR_FIELD_RED As Long
Public CLR_INV_CONSTRAINED_TEXT As Long
Public CLR_DISSONANCE As Long

Public Sub InitColors()
    CLR_FIELD_GREEN = RGB(198, 239, 206)
    CLR_FIELD_YELLOW = RGB(255, 235, 156)
    CLR_FIELD_RED = RGB(255, 199, 206)

    CLR_INV_CONSTRAINED_TEXT = RGB(192, 0, 0)      ' red font when projected inventory < demand
    CLR_DISSONANCE = RGB(255, 192, 0)              ' gold fill where a 15% dissonance flag fires
End Sub
