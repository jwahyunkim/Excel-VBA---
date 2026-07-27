Attribute VB_Name = "modGanttConfig"
Option Explicit

Public Const GANTT_HEADER_ROW_MONTH As Long = 1
Public Const GANTT_HEADER_ROW_WEEK As Long = 2
Public Const GANTT_HEADER_ROW_DAY As Long = 3
Public Const GANTT_HEADER_ROW_DATE As Long = 4

Public Const HEADER_ROW As Long = 4
Public Const DATA_START_ROW As Long = 5

Public Const COL_NO As String = "B"
Public Const COL_LEVEL As String = "C"
Public Const COL_TASK As String = "D"
Public Const COL_NOTE As String = "E"
Public Const COL_PLAN_START As String = "F"
Public Const COL_PLAN_END As String = "G"
Public Const COL_ACTUAL_START As String = "H"
Public Const COL_ACTUAL_END As String = "I"
Public Const COL_PROGRESS As String = "J"
Public Const COL_NORMAL_PROGRESS As String = "K"
Public Const COL_MANUAL_PROGRESS As String = "L"
Public Const COL_MANUAL_STATUS As String = "M"
Public Const COL_WEEKLY_REPORT As String = "N"
Public Const COL_DEV_PROGRESS As String = "O"
Public Const COL_PLAN_DAYS As String = "P"
Public Const COL_ACTUAL_DAYS As String = "Q"
Public Const COL_STATUS As String = "R"
Public Const COL_GANTT_START As String = "S"

Public Const HOLIDAY_SHEET_NAME As String = "휴일설정"
Public Const HOLIDAY_HEADER_ROW As Long = 1
Public Const HOLIDAY_DATA_START_ROW As Long = 2

Public Const HOLIDAY_COL_DATE As String = "A"
Public Const HOLIDAY_COL_TYPE As String = "B"
Public Const HOLIDAY_COL_DESC As String = "C"

Public Const HOLIDAY_TYPE_HOLIDAY As String = "휴일"
Public Const HOLIDAY_TYPE_WORKDAY As String = "근무일"

Public Const HIDE_SETTING_TITLE_CELL As String = "F1"
Public Const HIDE_SETTING_LEVEL_LABEL_CELL As String = "F2"
Public Const HIDE_SETTING_LEVEL_VALUE_CELL As String = "G2"
Public Const HIDE_SETTING_PERIOD_LABEL_CELL As String = "F3"
Public Const HIDE_SETTING_PERIOD_VALUE_CELL As String = "G3"
Public Const DISPLAY_SETTING_TITLE_CELL As String = "I1"
Public Const DISPLAY_SETTING_START_LABEL_CELL As String = "I2"
Public Const DISPLAY_SETTING_START_VALUE_CELL As String = "J2"
Public Const DISPLAY_SETTING_END_LABEL_CELL As String = "I3"
Public Const DISPLAY_SETTING_END_VALUE_CELL As String = "J3"
Public Const DISPLAY_SETTING_GANTT_ONLY_LABEL_CELL As String = "I4"
Public Const DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL As String = "J4"
Public Const DISPLAY_SETTING_REPORT_ONLY_LABEL_CELL As String = "I5"
Public Const DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL As String = "J5"
Public Const HIDE_EXCLUDE_NO_HEADER_CELL As String = "F5"
Public Const HIDE_EXCLUDE_DATE_HEADER_CELL As String = "G5"
Public Const HIDE_EXCLUDE_NO_START_CELL As String = "F6"
Public Const HIDE_EXCLUDE_DATE_START_CELL As String = "G6"

Public Const STATUS_NORMAL As String = "정상"
Public Const STATUS_CAUTION As String = "주의"
Public Const STATUS_DELAY As String = "지연"
Public Const STATUS_DONE As String = "완료"
Public Const STATUS_HOLD As String = "보류"
Public Const STATUS_WEEKLY_REPORT As String = "주간보고"
Public Const STATUS_DEV_PROGRESS As String = "개발진행"
Public Const STATUS_REPORT_DONE As String = "보고완료"
Public Const STATUS_ERROR As String = "오류"
Public Const REPORT_FILTER_ALL As String = "전체"
Public Const REPORT_FILTER_EMPTY As String = "빈값"
Public Const STATUS_DONE_WITH_HOLD_SUFFIX As String = "(보류 포함)"
