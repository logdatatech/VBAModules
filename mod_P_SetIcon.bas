Attribute VB_Name = "mod_P_SetIcon"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module Purpose:	This module contains code to change the icon of the Excel
'
' Dependencies:	NONE
'
' Author(s):	nathan@vba.guru
'
' LastChanged:	2016.Aug.04
'
' ReUsable:	Yes
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' 
' window. The code is compatible with 64-bit Office.
#If VBA7 And Win64 Then
'''''''''''''''''''''''''''''
' 64 bit Excel
'''''''''''''''''''''''''''''
Private Declare PtrSafe Function SendMessageA Lib "USER32" _
      (ByVal HWnd As LongPtr, ByVal wMsg As Longlong, ByVal wParam As Longlong, _
      ByVal lParam As Longlong) As LongPtr

Private Declare PtrSafe Function ExtractIconA Lib "shell32.dll" _
      (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, _
      ByVal nIconIndex As LongPtr) As Long

Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
Private Const WM_SETICON = &H80

#Else
'''''''''''''''''''''''''''''
' 32 bit Excel
'''''''''''''''''''''''''''''
Private Declare Function SendMessageA Lib "USER32" _
      (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
      ByVal lParam As Long) As Long

Private Declare Function ExtractIconA Lib "shell32.dll" _
      (ByVal hInst As Long, ByVal lpszExeFileName As String, _
      ByVal nIconIndex As Long) As Long

Private Const ICON_SMALL As Long = 0&
Private Const ICON_BIG As Long = 1&
Private Const WM_SETICON As LongPtr = &H80
#End If


Sub SetIcon(FileName As String, Optional Index As Long = 0)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetIcon
' This procedure sets the icon in the upper left corner of
' the main Excel window. FileName is the name of the file
' containing the icon. It may be an .ico file, an .exe file,
' or a .dll file. If it is an .ico file, Index must be 0
' or omitted. If it is an .exe or .dll file, Index is the
' 0-based index to the icon resource.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If VBA7 And Win64 Then
    ' 64 bit Excel
    Dim HWnd As LongPtr
    Dim HIcon As LongPtr
#Else
    ' 32 bit Excel
    Dim HWnd As Long
    Dim HIcon As Long
#End If
    Dim N As Long
    Dim s As String
    If Dir(FileName, vbNormal) = vbNullString Then
        ' file not found, get out
        Exit Sub
    End If
    ' get the extension of the file.
    N = InStrRev(FileName, ".")
    s = LCase(Mid(FileName, N + 1))
    ' ensure we have a valid file type
    Select Case s
        Case "exe", "ico", "dll"
            ' OK
        Case Else
            ' invalid file type
            Err.Raise 5
    End Select
    HWnd = Application.HWnd
    If HWnd = 0 Then
        Exit Sub
    End If
    HIcon = ExtractIconA(0, FileName, Index)
    If HIcon <> 0 Then
        SendMessageA HWnd, WM_SETICON, ICON_SMALL, HIcon
    End If
End Sub

Sub ResetIconToExcel()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ResetIconToExcel
' This resets the Excel window's icon. It is assumed to
' be the first icon resource in the Excel.exe file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim FName As String
    FName = Application.Path & "\excel.exe"
    SetIcon FName
End Sub
