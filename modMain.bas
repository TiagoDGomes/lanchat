Attribute VB_Name = "modMain"
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngICC = ICC_USEREX_CLASSES
       .lngSize = LenB(iccex)
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Sub Main()
    InitCommonControlsVB
    Load frmChat
End Sub


