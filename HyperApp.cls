VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HyperApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Objects and object event controllers
Private Script As New ScriptControl
Private WithEvents Script_Sink As ScriptControl
Attribute Script_Sink.VB_VarHelpID = -1
Private MainIE As WebBrowser
Private WithEvents MainIE_Sink As WebBrowser
Attribute MainIE_Sink.VB_VarHelpID = -1

'Member variables
Private m_szPostData As String
Private m_szHeaders As String
Private m_HTMLDoc As HTMLDocument

Public Function Init(ByRef brw As WebBrowser) As Boolean
    On Error Resume Next
    Set MainIE = brw
    Set MainIE_Sink = MainIE
    If Err = 0 Then
        Init = True
    Else
        Init = False
    End If
    
End Function

Public Sub AddObject(ByVal szObjName As String, ByRef obj As Object)
'AddObject - adds publically accessable objects to the script interface

    Script.AddObject szObjName, obj, True
    
End Sub

Private Sub Class_Initialize()
    Script.AllowUI = True   'or false, your choice
    Script.language = "VBScript"    'or JavaScript/JScript, or even PerlScript or Python (if those are installed)
    Script.UseSafeSubset = False    'or true, not much difference
    
    Set Script_Sink = Script
    
End Sub

Private Sub Class_Terminate()
    Set Script_Sink = Nothing
    
End Sub

Private Sub MainIE_Sink_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim szStatement As String
    
    'Retrieve document information
    m_szPostData = PostData
    m_szHeaders = Headers
        
    
    If URL Like "happ://*" Then
        Cancel = True
        
        'Retrieve the script statement (which follows 'happ://')
        szStatement = Mid$(URL, 8)
        If Right$(szStatement, 1) = "/" Then szStatement = DecodeHex$(Left$(szStatement, Len(szStatement) - 1))
        
        'Execute the script statement
        Script.ExecuteStatement szStatement
        
    End If
    
End Sub

Public Function GetLastPostData() As String
    GetLastPostData = m_szPostData
    
End Function

Public Function GetLastHeaders() As String
    GetLastHeaders = m_szHeaders
    
End Function

Public Function GetTextValue(ByVal nIndex As Integer, ByVal szName As String) As String
    'Written by Stephan
    'Posted by Chris Kesler
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer
    
    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1


        If m_HTMLDoc.Forms(nIndex)(q).Name = szName Then
            GetTextValue = m_HTMLDoc.Forms(nIndex)(q).Value
            Exit For
        End If
    Next q
    
End Function

Public Sub SetTextValue(ByVal nIndex As Integer, ByVal szName As String, ByVal szText As String)
    'Written by Stephan
    'Posted by Chris Kesler
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer
    
    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1

        If m_HTMLDoc.Forms(nIndex)(q).Name = szName Then

            m_HTMLDoc.Forms(nIndex)(q).Value = szText
            Exit For
        End If
    Next q
    
End Sub

Public Function GetChecked(ByVal nIndex As Integer, ByVal szName As String) As Boolean
    'Written by Ultimatum (with inspiration from Stephan)
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer
    
    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1


        If m_HTMLDoc.Forms(nIndex)(q).Name = szName Then
            GetChecked = m_HTMLDoc.Forms(nIndex)(q).Checked
            Exit For
        End If
    Next q
    
End Function

Public Sub SetChecked(ByVal nIndex As Integer, ByVal szName As String, ByVal bChecked As Boolean)
    'Written by Stephan
    'Posted by Chris Kesler
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1
        If m_HTMLDoc.Forms(nIndex)(q).Name = szName Then
            m_HTMLDoc.Forms(nIndex)(q).Checked = bChecked
            Exit For
        End If
    Next q
    
End Sub

Public Function GetComboValue(ByVal nIndex As Integer, ByVal szName As String) As String
    'Written by Ultimatum (with inspiration from Stephan)
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer, i As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1
        If (m_HTMLDoc.Forms(nIndex)(q).Name = szName) Then
            i = m_HTMLDoc.Forms(nIndex)(q).Options.selectedIndex
            GetComboValue = m_HTMLDoc.Forms(nIndex)(q).Options(i).Value
        End If
    Next q
    
End Function


Public Sub SetComboByValue(ByVal nIndex As Integer, ByVal szName As String, ByVal szValue As String)
    'Written by Stephan
    'Posted by Chris Kesler
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer, i As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1
        If (m_HTMLDoc.Forms(nIndex)(q).Name = szName) Then
            For i = 0 To m_HTMLDoc.Forms(nIndex)(q).length - 1
                If m_HTMLDoc.Forms(nIndex)(q).Options(i).Value = szValue Then
                    m_HTMLDoc.Forms(nIndex)(q).Options(i).Selected = True
                    Exit For
                End If
            Next i
        End If
    Next q
End Sub

Public Function GetComboText(ByVal nIndex As Integer, ByVal szName As String) As String
    'Written by Ultimatum (with inspiration from Stephan)
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer, i As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1
        If (m_HTMLDoc.Forms(nIndex)(q).Name = szName) Then
            i = m_HTMLDoc.Forms(nIndex)(q).Options.selectedIndex
            GetComboText = m_HTMLDoc.Forms(nIndex)(q).Options(i).Text
        End If
    Next q
    
End Function


Public Sub SetComboByText(ByVal nIndex As Integer, ByVal szName As String, ByVal szText As String)
    'Written by Stephan
    'Posted by Chris Kesler
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer, i As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1
        If (m_HTMLDoc.Forms(nIndex)(q).Name = szName) Then
            For i = 0 To m_HTMLDoc.Forms(nIndex)(q).length - 1
                If m_HTMLDoc.Forms(nIndex)(q).Options(i).Value = szText Then
                    m_HTMLDoc.Forms(nIndex)(q).Options(i).Selected = True
                    Exit For
                End If
            Next i
        End If
    Next q
End Sub

Public Function GetRadioState(ByVal nIndex As Integer, ByVal szGroupID As String, ByVal szName As String) As Boolean
    'Written by Ultimatum (with inspiration from Stephan)
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1

        If (m_HTMLDoc.Forms(nIndex)(q).Name = szGroupID) And (m_HTMLDoc.Forms(nIndex)(q).Value = szName) Then
            GetRadioState = m_HTMLDoc.Forms(nIndex)(q).Checked
            Exit For
        End If
        
    Next q

End Function

Public Sub SetRadioState(ByVal nIndex As Integer, ByVal szGroupID As String, ByVal szName As String, ByVal bOn As Boolean)
    'Written by Stephan
    'Posted by Chris Kesler
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1

        If (m_HTMLDoc.Forms(nIndex)(q).Name = szGroupID) And (m_HTMLDoc.Forms(nIndex)(q).Value = szName) Then
            m_HTMLDoc.Forms(nIndex)(q).Checked = True
            Exit For
        End If
        
    Next q
    
End Sub

Public Function GetCheckedRadioFromGroup(ByVal nIndex As Integer, ByVal szGroupID As String) As String
    'Written by Ultimatum
    
    Set m_HTMLDoc = MainIE.document
    
    Dim q As Integer

    For q = 0 To m_HTMLDoc.Forms(nIndex).length - 1

        If (m_HTMLDoc.Forms(nIndex)(q).Name = szGroupID) And (m_HTMLDoc.Forms(nIndex)(q).Checked = True) Then
            GetCheckedRadioFromGroup = m_HTMLDoc.Forms(nIndex)(q).Value
            Exit For
        End If
        
    Next q
    
End Function

Public Function GetLink(ByVal nIndex As Integer) As String
    Set m_HTMLDoc = MainIE.document
    GetLink = m_HTMLDoc.links(nIndex).href
    
End Function

Public Function GetImage(ByVal nIndex As Integer) As String
    Set m_HTMLDoc = MainIE.document
    GetImage = m_HTMLDoc.images(nIndex).src
        
End Function

Public Function GetSource() As String
    Set m_HTMLDoc = MainIE.document
    GetSource = m_HTMLDoc.All(0).outerHTML
    
End Function

Public Sub LoadHTML(ByVal szSource As String)
    MainIE.navigate "about:" & szSource
    
End Sub

Public Function GetTitle() As String
    Set m_HTMLDoc = MainIE.document
    GetTitle = m_HTMLDoc.Title
    
End Function

Public Property Get DocumentObject() As Object
    Set DocumentObject = m_HTMLDoc
    
End Property

Public Function DecodeHex(ByVal szString As String) As String
    DecodeHex = szString
    DecodeHex = Replace$(DecodeHex, "%00", Chr$(0))
    DecodeHex = Replace$(DecodeHex, "%01", Chr$(1))
    DecodeHex = Replace$(DecodeHex, "%02", Chr$(2))
    DecodeHex = Replace$(DecodeHex, "%03", Chr$(3))
    DecodeHex = Replace$(DecodeHex, "%04", Chr$(4))
    DecodeHex = Replace$(DecodeHex, "%05", Chr$(5))
    DecodeHex = Replace$(DecodeHex, "%06", Chr$(6))
    DecodeHex = Replace$(DecodeHex, "%07", Chr$(7))
    DecodeHex = Replace$(DecodeHex, "%08", Chr$(8))
    DecodeHex = Replace$(DecodeHex, "%09", Chr$(9))
    DecodeHex = Replace$(DecodeHex, "%0A", Chr$(10))
    DecodeHex = Replace$(DecodeHex, "%0B", Chr$(11))
    DecodeHex = Replace$(DecodeHex, "%0C", Chr$(12))
    DecodeHex = Replace$(DecodeHex, "%0D", Chr$(13))
    DecodeHex = Replace$(DecodeHex, "%0E", Chr$(14))
    DecodeHex = Replace$(DecodeHex, "%0F", Chr$(15))
    DecodeHex = Replace$(DecodeHex, "%10", Chr$(16))
    DecodeHex = Replace$(DecodeHex, "%11", Chr$(17))
    DecodeHex = Replace$(DecodeHex, "%12", Chr$(18))
    DecodeHex = Replace$(DecodeHex, "%13", Chr$(19))
    DecodeHex = Replace$(DecodeHex, "%14", Chr$(20))
    DecodeHex = Replace$(DecodeHex, "%15", Chr$(21))
    DecodeHex = Replace$(DecodeHex, "%16", Chr$(22))
    DecodeHex = Replace$(DecodeHex, "%17", Chr$(23))
    DecodeHex = Replace$(DecodeHex, "%18", Chr$(24))
    DecodeHex = Replace$(DecodeHex, "%19", Chr$(25))
    DecodeHex = Replace$(DecodeHex, "%1A", Chr$(26))
    DecodeHex = Replace$(DecodeHex, "%1B", Chr$(27))
    DecodeHex = Replace$(DecodeHex, "%1C", Chr$(28))
    DecodeHex = Replace$(DecodeHex, "%1D", Chr$(29))
    DecodeHex = Replace$(DecodeHex, "%1E", Chr$(30))
    DecodeHex = Replace$(DecodeHex, "%1F", Chr$(31))
    DecodeHex = Replace$(DecodeHex, "%20", Chr$(32))
    DecodeHex = Replace$(DecodeHex, "%21", Chr$(33))
    DecodeHex = Replace$(DecodeHex, "%22", Chr$(34))
    DecodeHex = Replace$(DecodeHex, "%23", Chr$(35))
    DecodeHex = Replace$(DecodeHex, "%24", Chr$(36))
    DecodeHex = Replace$(DecodeHex, "%25", Chr$(37))
    DecodeHex = Replace$(DecodeHex, "%26", Chr$(38))
    DecodeHex = Replace$(DecodeHex, "%27", Chr$(39))
    DecodeHex = Replace$(DecodeHex, "%28", Chr$(40))
    DecodeHex = Replace$(DecodeHex, "%29", Chr$(41))
    DecodeHex = Replace$(DecodeHex, "%2A", Chr$(42))
    DecodeHex = Replace$(DecodeHex, "%2B", Chr$(43))
    DecodeHex = Replace$(DecodeHex, "%2C", Chr$(44))
    DecodeHex = Replace$(DecodeHex, "%2D", Chr$(45))
    DecodeHex = Replace$(DecodeHex, "%2E", Chr$(46))
    DecodeHex = Replace$(DecodeHex, "%2F", Chr$(47))
    DecodeHex = Replace$(DecodeHex, "%30", Chr$(48))
    DecodeHex = Replace$(DecodeHex, "%31", Chr$(49))
    DecodeHex = Replace$(DecodeHex, "%32", Chr$(50))
    DecodeHex = Replace$(DecodeHex, "%33", Chr$(51))
    DecodeHex = Replace$(DecodeHex, "%34", Chr$(52))
    DecodeHex = Replace$(DecodeHex, "%35", Chr$(53))
    DecodeHex = Replace$(DecodeHex, "%36", Chr$(54))
    DecodeHex = Replace$(DecodeHex, "%37", Chr$(55))
    DecodeHex = Replace$(DecodeHex, "%38", Chr$(56))
    DecodeHex = Replace$(DecodeHex, "%39", Chr$(57))
    DecodeHex = Replace$(DecodeHex, "%3A", Chr$(58))
    DecodeHex = Replace$(DecodeHex, "%3B", Chr$(59))
    DecodeHex = Replace$(DecodeHex, "%3C", Chr$(60))
    DecodeHex = Replace$(DecodeHex, "%3D", Chr$(61))
    DecodeHex = Replace$(DecodeHex, "%3E", Chr$(62))
    DecodeHex = Replace$(DecodeHex, "%3F", Chr$(63))
    DecodeHex = Replace$(DecodeHex, "%40", Chr$(64))
    DecodeHex = Replace$(DecodeHex, "%41", Chr$(65))
    DecodeHex = Replace$(DecodeHex, "%42", Chr$(66))
    DecodeHex = Replace$(DecodeHex, "%43", Chr$(67))
    DecodeHex = Replace$(DecodeHex, "%44", Chr$(68))
    DecodeHex = Replace$(DecodeHex, "%45", Chr$(69))
    DecodeHex = Replace$(DecodeHex, "%46", Chr$(70))
    DecodeHex = Replace$(DecodeHex, "%47", Chr$(71))
    DecodeHex = Replace$(DecodeHex, "%48", Chr$(72))
    DecodeHex = Replace$(DecodeHex, "%49", Chr$(73))
    DecodeHex = Replace$(DecodeHex, "%4A", Chr$(74))
    DecodeHex = Replace$(DecodeHex, "%4B", Chr$(75))
    DecodeHex = Replace$(DecodeHex, "%4C", Chr$(76))
    DecodeHex = Replace$(DecodeHex, "%4D", Chr$(77))
    DecodeHex = Replace$(DecodeHex, "%4E", Chr$(78))
    DecodeHex = Replace$(DecodeHex, "%4F", Chr$(79))
    DecodeHex = Replace$(DecodeHex, "%50", Chr$(80))
    DecodeHex = Replace$(DecodeHex, "%51", Chr$(81))
    DecodeHex = Replace$(DecodeHex, "%52", Chr$(82))
    DecodeHex = Replace$(DecodeHex, "%53", Chr$(83))
    DecodeHex = Replace$(DecodeHex, "%54", Chr$(84))
    DecodeHex = Replace$(DecodeHex, "%55", Chr$(85))
    DecodeHex = Replace$(DecodeHex, "%56", Chr$(86))
    DecodeHex = Replace$(DecodeHex, "%57", Chr$(87))
    DecodeHex = Replace$(DecodeHex, "%58", Chr$(88))
    DecodeHex = Replace$(DecodeHex, "%59", Chr$(89))
    DecodeHex = Replace$(DecodeHex, "%5A", Chr$(90))
    DecodeHex = Replace$(DecodeHex, "%5B", Chr$(91))
    DecodeHex = Replace$(DecodeHex, "%5C", Chr$(92))
    DecodeHex = Replace$(DecodeHex, "%5D", Chr$(93))
    DecodeHex = Replace$(DecodeHex, "%5E", Chr$(94))
    DecodeHex = Replace$(DecodeHex, "%5F", Chr$(95))
    DecodeHex = Replace$(DecodeHex, "%60", Chr$(96))
    DecodeHex = Replace$(DecodeHex, "%61", Chr$(97))
    DecodeHex = Replace$(DecodeHex, "%62", Chr$(98))
    DecodeHex = Replace$(DecodeHex, "%63", Chr$(99))
    DecodeHex = Replace$(DecodeHex, "%64", Chr$(100))
    DecodeHex = Replace$(DecodeHex, "%65", Chr$(101))
    DecodeHex = Replace$(DecodeHex, "%66", Chr$(102))
    DecodeHex = Replace$(DecodeHex, "%67", Chr$(103))
    DecodeHex = Replace$(DecodeHex, "%68", Chr$(104))
    DecodeHex = Replace$(DecodeHex, "%69", Chr$(105))
    DecodeHex = Replace$(DecodeHex, "%6A", Chr$(106))
    DecodeHex = Replace$(DecodeHex, "%6B", Chr$(107))
    DecodeHex = Replace$(DecodeHex, "%6C", Chr$(108))
    DecodeHex = Replace$(DecodeHex, "%6D", Chr$(109))
    DecodeHex = Replace$(DecodeHex, "%6E", Chr$(110))
    DecodeHex = Replace$(DecodeHex, "%6F", Chr$(111))
    DecodeHex = Replace$(DecodeHex, "%70", Chr$(112))
    DecodeHex = Replace$(DecodeHex, "%71", Chr$(113))
    DecodeHex = Replace$(DecodeHex, "%72", Chr$(114))
    DecodeHex = Replace$(DecodeHex, "%73", Chr$(115))
    DecodeHex = Replace$(DecodeHex, "%74", Chr$(116))
    DecodeHex = Replace$(DecodeHex, "%75", Chr$(117))
    DecodeHex = Replace$(DecodeHex, "%76", Chr$(118))
    DecodeHex = Replace$(DecodeHex, "%77", Chr$(119))
    DecodeHex = Replace$(DecodeHex, "%78", Chr$(120))
    DecodeHex = Replace$(DecodeHex, "%79", Chr$(121))
    DecodeHex = Replace$(DecodeHex, "%7A", Chr$(122))
    DecodeHex = Replace$(DecodeHex, "%7B", Chr$(123))
    DecodeHex = Replace$(DecodeHex, "%7C", Chr$(124))
    DecodeHex = Replace$(DecodeHex, "%7D", Chr$(125))
    DecodeHex = Replace$(DecodeHex, "%7E", Chr$(126))
    DecodeHex = Replace$(DecodeHex, "%7F", Chr$(127))
    DecodeHex = Replace$(DecodeHex, "%80", Chr$(128))
    DecodeHex = Replace$(DecodeHex, "%81", Chr$(129))
    DecodeHex = Replace$(DecodeHex, "%82", Chr$(130))
    DecodeHex = Replace$(DecodeHex, "%83", Chr$(131))
    DecodeHex = Replace$(DecodeHex, "%84", Chr$(132))
    DecodeHex = Replace$(DecodeHex, "%85", Chr$(133))
    DecodeHex = Replace$(DecodeHex, "%86", Chr$(134))
    DecodeHex = Replace$(DecodeHex, "%87", Chr$(135))
    DecodeHex = Replace$(DecodeHex, "%88", Chr$(136))
    DecodeHex = Replace$(DecodeHex, "%89", Chr$(137))
    DecodeHex = Replace$(DecodeHex, "%8A", Chr$(138))
    DecodeHex = Replace$(DecodeHex, "%8B", Chr$(139))
    DecodeHex = Replace$(DecodeHex, "%8C", Chr$(140))
    DecodeHex = Replace$(DecodeHex, "%8D", Chr$(141))
    DecodeHex = Replace$(DecodeHex, "%8E", Chr$(142))
    DecodeHex = Replace$(DecodeHex, "%8F", Chr$(143))
    DecodeHex = Replace$(DecodeHex, "%90", Chr$(144))
    DecodeHex = Replace$(DecodeHex, "%91", Chr$(145))
    DecodeHex = Replace$(DecodeHex, "%92", Chr$(146))
    DecodeHex = Replace$(DecodeHex, "%93", Chr$(147))
    DecodeHex = Replace$(DecodeHex, "%94", Chr$(148))
    DecodeHex = Replace$(DecodeHex, "%95", Chr$(149))
    DecodeHex = Replace$(DecodeHex, "%96", Chr$(150))
    DecodeHex = Replace$(DecodeHex, "%97", Chr$(151))
    DecodeHex = Replace$(DecodeHex, "%98", Chr$(152))
    DecodeHex = Replace$(DecodeHex, "%99", Chr$(153))
    DecodeHex = Replace$(DecodeHex, "%9A", Chr$(154))
    DecodeHex = Replace$(DecodeHex, "%9B", Chr$(155))
    DecodeHex = Replace$(DecodeHex, "%9C", Chr$(156))
    DecodeHex = Replace$(DecodeHex, "%9D", Chr$(157))
    DecodeHex = Replace$(DecodeHex, "%9E", Chr$(158))
    DecodeHex = Replace$(DecodeHex, "%9F", Chr$(159))
    DecodeHex = Replace$(DecodeHex, "%A0", Chr$(160))
    DecodeHex = Replace$(DecodeHex, "%A1", Chr$(161))
    DecodeHex = Replace$(DecodeHex, "%A2", Chr$(162))
    DecodeHex = Replace$(DecodeHex, "%A3", Chr$(163))
    DecodeHex = Replace$(DecodeHex, "%A4", Chr$(164))
    DecodeHex = Replace$(DecodeHex, "%A5", Chr$(165))
    DecodeHex = Replace$(DecodeHex, "%A6", Chr$(166))
    DecodeHex = Replace$(DecodeHex, "%A7", Chr$(167))
    DecodeHex = Replace$(DecodeHex, "%A8", Chr$(168))
    DecodeHex = Replace$(DecodeHex, "%A9", Chr$(169))
    DecodeHex = Replace$(DecodeHex, "%AA", Chr$(170))
    DecodeHex = Replace$(DecodeHex, "%AB", Chr$(171))
    DecodeHex = Replace$(DecodeHex, "%AC", Chr$(172))
    DecodeHex = Replace$(DecodeHex, "%AD", Chr$(173))
    DecodeHex = Replace$(DecodeHex, "%AE", Chr$(174))
    DecodeHex = Replace$(DecodeHex, "%AF", Chr$(175))
    DecodeHex = Replace$(DecodeHex, "%B0", Chr$(176))
    DecodeHex = Replace$(DecodeHex, "%B1", Chr$(177))
    DecodeHex = Replace$(DecodeHex, "%B2", Chr$(178))
    DecodeHex = Replace$(DecodeHex, "%B3", Chr$(179))
    DecodeHex = Replace$(DecodeHex, "%B4", Chr$(180))
    DecodeHex = Replace$(DecodeHex, "%B5", Chr$(181))
    DecodeHex = Replace$(DecodeHex, "%B6", Chr$(182))
    DecodeHex = Replace$(DecodeHex, "%B7", Chr$(183))
    DecodeHex = Replace$(DecodeHex, "%B8", Chr$(184))
    DecodeHex = Replace$(DecodeHex, "%B9", Chr$(185))
    DecodeHex = Replace$(DecodeHex, "%BA", Chr$(186))
    DecodeHex = Replace$(DecodeHex, "%BB", Chr$(187))
    DecodeHex = Replace$(DecodeHex, "%BC", Chr$(188))
    DecodeHex = Replace$(DecodeHex, "%BD", Chr$(189))
    DecodeHex = Replace$(DecodeHex, "%BE", Chr$(190))
    DecodeHex = Replace$(DecodeHex, "%BF", Chr$(191))
    DecodeHex = Replace$(DecodeHex, "%C0", Chr$(192))
    DecodeHex = Replace$(DecodeHex, "%C1", Chr$(193))
    DecodeHex = Replace$(DecodeHex, "%C2", Chr$(194))
    DecodeHex = Replace$(DecodeHex, "%C3", Chr$(195))
    DecodeHex = Replace$(DecodeHex, "%C4", Chr$(196))
    DecodeHex = Replace$(DecodeHex, "%C5", Chr$(197))
    DecodeHex = Replace$(DecodeHex, "%C6", Chr$(198))
    DecodeHex = Replace$(DecodeHex, "%C7", Chr$(199))
    DecodeHex = Replace$(DecodeHex, "%C8", Chr$(200))
    DecodeHex = Replace$(DecodeHex, "%C9", Chr$(201))
    DecodeHex = Replace$(DecodeHex, "%CA", Chr$(202))
    DecodeHex = Replace$(DecodeHex, "%CB", Chr$(203))
    DecodeHex = Replace$(DecodeHex, "%CC", Chr$(204))
    DecodeHex = Replace$(DecodeHex, "%CD", Chr$(205))
    DecodeHex = Replace$(DecodeHex, "%CE", Chr$(206))
    DecodeHex = Replace$(DecodeHex, "%CF", Chr$(207))
    DecodeHex = Replace$(DecodeHex, "%D0", Chr$(208))
    DecodeHex = Replace$(DecodeHex, "%D1", Chr$(209))
    DecodeHex = Replace$(DecodeHex, "%D2", Chr$(210))
    DecodeHex = Replace$(DecodeHex, "%D3", Chr$(211))
    DecodeHex = Replace$(DecodeHex, "%D4", Chr$(212))
    DecodeHex = Replace$(DecodeHex, "%D5", Chr$(213))
    DecodeHex = Replace$(DecodeHex, "%D6", Chr$(214))
    DecodeHex = Replace$(DecodeHex, "%D7", Chr$(215))
    DecodeHex = Replace$(DecodeHex, "%D8", Chr$(216))
    DecodeHex = Replace$(DecodeHex, "%D9", Chr$(217))
    DecodeHex = Replace$(DecodeHex, "%DA", Chr$(218))
    DecodeHex = Replace$(DecodeHex, "%DB", Chr$(219))
    DecodeHex = Replace$(DecodeHex, "%DC", Chr$(220))
    DecodeHex = Replace$(DecodeHex, "%DD", Chr$(221))
    DecodeHex = Replace$(DecodeHex, "%DE", Chr$(222))
    DecodeHex = Replace$(DecodeHex, "%DF", Chr$(223))
    DecodeHex = Replace$(DecodeHex, "%E0", Chr$(224))
    DecodeHex = Replace$(DecodeHex, "%E1", Chr$(225))
    DecodeHex = Replace$(DecodeHex, "%E2", Chr$(226))
    DecodeHex = Replace$(DecodeHex, "%E3", Chr$(227))
    DecodeHex = Replace$(DecodeHex, "%E4", Chr$(228))
    DecodeHex = Replace$(DecodeHex, "%E5", Chr$(229))
    DecodeHex = Replace$(DecodeHex, "%E6", Chr$(230))
    DecodeHex = Replace$(DecodeHex, "%E7", Chr$(231))
    DecodeHex = Replace$(DecodeHex, "%E8", Chr$(232))
    DecodeHex = Replace$(DecodeHex, "%E9", Chr$(233))
    DecodeHex = Replace$(DecodeHex, "%EA", Chr$(234))
    DecodeHex = Replace$(DecodeHex, "%EB", Chr$(235))
    DecodeHex = Replace$(DecodeHex, "%EC", Chr$(236))
    DecodeHex = Replace$(DecodeHex, "%ED", Chr$(237))
    DecodeHex = Replace$(DecodeHex, "%EE", Chr$(238))
    DecodeHex = Replace$(DecodeHex, "%EF", Chr$(239))
    DecodeHex = Replace$(DecodeHex, "%F0", Chr$(240))
    DecodeHex = Replace$(DecodeHex, "%F1", Chr$(241))
    DecodeHex = Replace$(DecodeHex, "%F2", Chr$(242))
    DecodeHex = Replace$(DecodeHex, "%F3", Chr$(243))
    DecodeHex = Replace$(DecodeHex, "%F4", Chr$(244))
    DecodeHex = Replace$(DecodeHex, "%F5", Chr$(245))
    DecodeHex = Replace$(DecodeHex, "%F6", Chr$(246))
    DecodeHex = Replace$(DecodeHex, "%F7", Chr$(247))
    DecodeHex = Replace$(DecodeHex, "%F8", Chr$(248))
    DecodeHex = Replace$(DecodeHex, "%F9", Chr$(249))
    DecodeHex = Replace$(DecodeHex, "%FA", Chr$(250))
    DecodeHex = Replace$(DecodeHex, "%FB", Chr$(251))
    DecodeHex = Replace$(DecodeHex, "%FC", Chr$(252))
    DecodeHex = Replace$(DecodeHex, "%FD", Chr$(253))
    DecodeHex = Replace$(DecodeHex, "%FE", Chr$(254))
    DecodeHex = Replace$(DecodeHex, "%FF", Chr$(255))
    
End Function


Private Sub Script_Sink_Error()
    Err.Raise Script.Error.Number, Script.Error.Source, Script.Error.Description
    
End Sub
