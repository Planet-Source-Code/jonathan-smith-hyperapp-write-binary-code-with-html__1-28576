VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmInterface 
   Caption         =   "HyperApp Demonstration"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser brw 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      ExtentX         =   15901
      ExtentY         =   12303
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private HTMLApp As New HyperApp
Private CCDRom As New HACDRom
Private CFormHandler As New HAFormHandler

Private Sub Form_Load()
    
    brw.Navigate "file:///" & App.Path & "\hyperapp.htm"
    
    If Not HTMLApp.Init(brw) Then
        MsgBox "Error: Unable to initialize HyperApp"
        Exit Sub
    End If
    
    CFormHandler.Init HTMLApp
    
    HTMLApp.AddObject "CDRom", CCDRom
    HTMLApp.AddObject "FormHandler", CFormHandler
    
    
End Sub
