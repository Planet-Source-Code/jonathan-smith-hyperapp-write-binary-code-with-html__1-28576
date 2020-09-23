VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HACDRom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'HyperApp Demo Object
'CDRom - open and closes the CD-ROM =)

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCmd As String, ByVal lpszRet As String, ByVal uRet As Long, ByVal hCallback As Long) As Long
Private bIsOpen As Boolean

Public Property Get TrayOpen() As Boolean
    TrayOpen = bIsOpen
End Property

Public Sub ToggleTray()
    If TrayOpen Then
        Call mciSendString("set CDAudio door closed", 0&, 0, 0)
    Else
        Call mciSendString("set CDAudio door open", 0&, 0, 0)
    End If
    
    bIsOpen = Not bIsOpen
    
End Sub
