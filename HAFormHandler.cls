VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HAFormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'HyperApp Demo Object
'FormHandler - sample object for retrieving form information

Private m_hApp As HyperApp

Public Sub Init(happ As HyperApp)
    Set m_hApp = happ
    
End Sub

Public Sub CheckForm()

    MsgBox "POST Data: " & StrConv(m_hApp.GetLastPostData, vbUnicode)
    MsgBox "Headers: " & m_hApp.GetLastHeaders
    
    MsgBox "Full Name text box contains " & m_hApp.GetTextValue(0, "FullName")
    MsgBox "Email text box contains " & m_hApp.GetTextValue(0, "EMail")
    
End Sub
