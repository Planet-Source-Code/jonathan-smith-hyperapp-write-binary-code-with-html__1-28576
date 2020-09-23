VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLProcLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'HTML Procedure Library
'written by Ultimatum

Public Function DecodeHex(ByVal szString As String) As String
    DeHex = szString
    DeHex = Replace$(DeHex, "%00", Chr$(0))
    DeHex = Replace$(DeHex, "%01", Chr$(1))
    DeHex = Replace$(DeHex, "%02", Chr$(2))
    DeHex = Replace$(DeHex, "%03", Chr$(3))
    DeHex = Replace$(DeHex, "%04", Chr$(4))
    DeHex = Replace$(DeHex, "%05", Chr$(5))
    DeHex = Replace$(DeHex, "%06", Chr$(6))
    DeHex = Replace$(DeHex, "%07", Chr$(7))
    DeHex = Replace$(DeHex, "%08", Chr$(8))
    DeHex = Replace$(DeHex, "%09", Chr$(9))
    DeHex = Replace$(DeHex, "%0A", Chr$(10))
    DeHex = Replace$(DeHex, "%0B", Chr$(11))
    DeHex = Replace$(DeHex, "%0C", Chr$(12))
    DeHex = Replace$(DeHex, "%0D", Chr$(13))
    DeHex = Replace$(DeHex, "%0E", Chr$(14))
    DeHex = Replace$(DeHex, "%0F", Chr$(15))
    DeHex = Replace$(DeHex, "%10", Chr$(16))
    DeHex = Replace$(DeHex, "%11", Chr$(17))
    DeHex = Replace$(DeHex, "%12", Chr$(18))
    DeHex = Replace$(DeHex, "%13", Chr$(19))
    DeHex = Replace$(DeHex, "%14", Chr$(20))
    DeHex = Replace$(DeHex, "%15", Chr$(21))
    DeHex = Replace$(DeHex, "%16", Chr$(22))
    DeHex = Replace$(DeHex, "%17", Chr$(23))
    DeHex = Replace$(DeHex, "%18", Chr$(24))
    DeHex = Replace$(DeHex, "%19", Chr$(25))
    DeHex = Replace$(DeHex, "%1A", Chr$(26))
    DeHex = Replace$(DeHex, "%1B", Chr$(27))
    DeHex = Replace$(DeHex, "%1C", Chr$(28))
    DeHex = Replace$(DeHex, "%1D", Chr$(29))
    DeHex = Replace$(DeHex, "%1E", Chr$(30))
    DeHex = Replace$(DeHex, "%1F", Chr$(31))
    DeHex = Replace$(DeHex, "%20", Chr$(32))
    DeHex = Replace$(DeHex, "%21", Chr$(33))
    DeHex = Replace$(DeHex, "%22", Chr$(34))
    DeHex = Replace$(DeHex, "%23", Chr$(35))
    DeHex = Replace$(DeHex, "%24", Chr$(36))
    DeHex = Replace$(DeHex, "%25", Chr$(37))
    DeHex = Replace$(DeHex, "%26", Chr$(38))
    DeHex = Replace$(DeHex, "%27", Chr$(39))
    DeHex = Replace$(DeHex, "%28", Chr$(40))
    DeHex = Replace$(DeHex, "%29", Chr$(41))
    DeHex = Replace$(DeHex, "%2A", Chr$(42))
    DeHex = Replace$(DeHex, "%2B", Chr$(43))
    DeHex = Replace$(DeHex, "%2C", Chr$(44))
    DeHex = Replace$(DeHex, "%2D", Chr$(45))
    DeHex = Replace$(DeHex, "%2E", Chr$(46))
    DeHex = Replace$(DeHex, "%2F", Chr$(47))
    DeHex = Replace$(DeHex, "%30", Chr$(48))
    DeHex = Replace$(DeHex, "%31", Chr$(49))
    DeHex = Replace$(DeHex, "%32", Chr$(50))
    DeHex = Replace$(DeHex, "%33", Chr$(51))
    DeHex = Replace$(DeHex, "%34", Chr$(52))
    DeHex = Replace$(DeHex, "%35", Chr$(53))
    DeHex = Replace$(DeHex, "%36", Chr$(54))
    DeHex = Replace$(DeHex, "%37", Chr$(55))
    DeHex = Replace$(DeHex, "%38", Chr$(56))
    DeHex = Replace$(DeHex, "%39", Chr$(57))
    DeHex = Replace$(DeHex, "%3A", Chr$(58))
    DeHex = Replace$(DeHex, "%3B", Chr$(59))
    DeHex = Replace$(DeHex, "%3C", Chr$(60))
    DeHex = Replace$(DeHex, "%3D", Chr$(61))
    DeHex = Replace$(DeHex, "%3E", Chr$(62))
    DeHex = Replace$(DeHex, "%3F", Chr$(63))
    DeHex = Replace$(DeHex, "%40", Chr$(64))
    DeHex = Replace$(DeHex, "%41", Chr$(65))
    DeHex = Replace$(DeHex, "%42", Chr$(66))
    DeHex = Replace$(DeHex, "%43", Chr$(67))
    DeHex = Replace$(DeHex, "%44", Chr$(68))
    DeHex = Replace$(DeHex, "%45", Chr$(69))
    DeHex = Replace$(DeHex, "%46", Chr$(70))
    DeHex = Replace$(DeHex, "%47", Chr$(71))
    DeHex = Replace$(DeHex, "%48", Chr$(72))
    DeHex = Replace$(DeHex, "%49", Chr$(73))
    DeHex = Replace$(DeHex, "%4A", Chr$(74))
    DeHex = Replace$(DeHex, "%4B", Chr$(75))
    DeHex = Replace$(DeHex, "%4C", Chr$(76))
    DeHex = Replace$(DeHex, "%4D", Chr$(77))
    DeHex = Replace$(DeHex, "%4E", Chr$(78))
    DeHex = Replace$(DeHex, "%4F", Chr$(79))
    DeHex = Replace$(DeHex, "%50", Chr$(80))
    DeHex = Replace$(DeHex, "%51", Chr$(81))
    DeHex = Replace$(DeHex, "%52", Chr$(82))
    DeHex = Replace$(DeHex, "%53", Chr$(83))
    DeHex = Replace$(DeHex, "%54", Chr$(84))
    DeHex = Replace$(DeHex, "%55", Chr$(85))
    DeHex = Replace$(DeHex, "%56", Chr$(86))
    DeHex = Replace$(DeHex, "%57", Chr$(87))
    DeHex = Replace$(DeHex, "%58", Chr$(88))
    DeHex = Replace$(DeHex, "%59", Chr$(89))
    DeHex = Replace$(DeHex, "%5A", Chr$(90))
    DeHex = Replace$(DeHex, "%5B", Chr$(91))
    DeHex = Replace$(DeHex, "%5C", Chr$(92))
    DeHex = Replace$(DeHex, "%5D", Chr$(93))
    DeHex = Replace$(DeHex, "%5E", Chr$(94))
    DeHex = Replace$(DeHex, "%5F", Chr$(95))
    DeHex = Replace$(DeHex, "%60", Chr$(96))
    DeHex = Replace$(DeHex, "%61", Chr$(97))
    DeHex = Replace$(DeHex, "%62", Chr$(98))
    DeHex = Replace$(DeHex, "%63", Chr$(99))
    DeHex = Replace$(DeHex, "%64", Chr$(100))
    DeHex = Replace$(DeHex, "%65", Chr$(101))
    DeHex = Replace$(DeHex, "%66", Chr$(102))
    DeHex = Replace$(DeHex, "%67", Chr$(103))
    DeHex = Replace$(DeHex, "%68", Chr$(104))
    DeHex = Replace$(DeHex, "%69", Chr$(105))
    DeHex = Replace$(DeHex, "%6A", Chr$(106))
    DeHex = Replace$(DeHex, "%6B", Chr$(107))
    DeHex = Replace$(DeHex, "%6C", Chr$(108))
    DeHex = Replace$(DeHex, "%6D", Chr$(109))
    DeHex = Replace$(DeHex, "%6E", Chr$(110))
    DeHex = Replace$(DeHex, "%6F", Chr$(111))
    DeHex = Replace$(DeHex, "%70", Chr$(112))
    DeHex = Replace$(DeHex, "%71", Chr$(113))
    DeHex = Replace$(DeHex, "%72", Chr$(114))
    DeHex = Replace$(DeHex, "%73", Chr$(115))
    DeHex = Replace$(DeHex, "%74", Chr$(116))
    DeHex = Replace$(DeHex, "%75", Chr$(117))
    DeHex = Replace$(DeHex, "%76", Chr$(118))
    DeHex = Replace$(DeHex, "%77", Chr$(119))
    DeHex = Replace$(DeHex, "%78", Chr$(120))
    DeHex = Replace$(DeHex, "%79", Chr$(121))
    DeHex = Replace$(DeHex, "%7A", Chr$(122))
    DeHex = Replace$(DeHex, "%7B", Chr$(123))
    DeHex = Replace$(DeHex, "%7C", Chr$(124))
    DeHex = Replace$(DeHex, "%7D", Chr$(125))
    DeHex = Replace$(DeHex, "%7E", Chr$(126))
    DeHex = Replace$(DeHex, "%7F", Chr$(127))
    DeHex = Replace$(DeHex, "%80", Chr$(128))
    DeHex = Replace$(DeHex, "%81", Chr$(129))
    DeHex = Replace$(DeHex, "%82", Chr$(130))
    DeHex = Replace$(DeHex, "%83", Chr$(131))
    DeHex = Replace$(DeHex, "%84", Chr$(132))
    DeHex = Replace$(DeHex, "%85", Chr$(133))
    DeHex = Replace$(DeHex, "%86", Chr$(134))
    DeHex = Replace$(DeHex, "%87", Chr$(135))
    DeHex = Replace$(DeHex, "%88", Chr$(136))
    DeHex = Replace$(DeHex, "%89", Chr$(137))
    DeHex = Replace$(DeHex, "%8A", Chr$(138))
    DeHex = Replace$(DeHex, "%8B", Chr$(139))
    DeHex = Replace$(DeHex, "%8C", Chr$(140))
    DeHex = Replace$(DeHex, "%8D", Chr$(141))
    DeHex = Replace$(DeHex, "%8E", Chr$(142))
    DeHex = Replace$(DeHex, "%8F", Chr$(143))
    DeHex = Replace$(DeHex, "%90", Chr$(144))
    DeHex = Replace$(DeHex, "%91", Chr$(145))
    DeHex = Replace$(DeHex, "%92", Chr$(146))
    DeHex = Replace$(DeHex, "%93", Chr$(147))
    DeHex = Replace$(DeHex, "%94", Chr$(148))
    DeHex = Replace$(DeHex, "%95", Chr$(149))
    DeHex = Replace$(DeHex, "%96", Chr$(150))
    DeHex = Replace$(DeHex, "%97", Chr$(151))
    DeHex = Replace$(DeHex, "%98", Chr$(152))
    DeHex = Replace$(DeHex, "%99", Chr$(153))
    DeHex = Replace$(DeHex, "%9A", Chr$(154))
    DeHex = Replace$(DeHex, "%9B", Chr$(155))
    DeHex = Replace$(DeHex, "%9C", Chr$(156))
    DeHex = Replace$(DeHex, "%9D", Chr$(157))
    DeHex = Replace$(DeHex, "%9E", Chr$(158))
    DeHex = Replace$(DeHex, "%9F", Chr$(159))
    DeHex = Replace$(DeHex, "%A0", Chr$(160))
    DeHex = Replace$(DeHex, "%A1", Chr$(161))
    DeHex = Replace$(DeHex, "%A2", Chr$(162))
    DeHex = Replace$(DeHex, "%A3", Chr$(163))
    DeHex = Replace$(DeHex, "%A4", Chr$(164))
    DeHex = Replace$(DeHex, "%A5", Chr$(165))
    DeHex = Replace$(DeHex, "%A6", Chr$(166))
    DeHex = Replace$(DeHex, "%A7", Chr$(167))
    DeHex = Replace$(DeHex, "%A8", Chr$(168))
    DeHex = Replace$(DeHex, "%A9", Chr$(169))
    DeHex = Replace$(DeHex, "%AA", Chr$(170))
    DeHex = Replace$(DeHex, "%AB", Chr$(171))
    DeHex = Replace$(DeHex, "%AC", Chr$(172))
    DeHex = Replace$(DeHex, "%AD", Chr$(173))
    DeHex = Replace$(DeHex, "%AE", Chr$(174))
    DeHex = Replace$(DeHex, "%AF", Chr$(175))
    DeHex = Replace$(DeHex, "%B0", Chr$(176))
    DeHex = Replace$(DeHex, "%B1", Chr$(177))
    DeHex = Replace$(DeHex, "%B2", Chr$(178))
    DeHex = Replace$(DeHex, "%B3", Chr$(179))
    DeHex = Replace$(DeHex, "%B4", Chr$(180))
    DeHex = Replace$(DeHex, "%B5", Chr$(181))
    DeHex = Replace$(DeHex, "%B6", Chr$(182))
    DeHex = Replace$(DeHex, "%B7", Chr$(183))
    DeHex = Replace$(DeHex, "%B8", Chr$(184))
    DeHex = Replace$(DeHex, "%B9", Chr$(185))
    DeHex = Replace$(DeHex, "%BA", Chr$(186))
    DeHex = Replace$(DeHex, "%BB", Chr$(187))
    DeHex = Replace$(DeHex, "%BC", Chr$(188))
    DeHex = Replace$(DeHex, "%BD", Chr$(189))
    DeHex = Replace$(DeHex, "%BE", Chr$(190))
    DeHex = Replace$(DeHex, "%BF", Chr$(191))
    DeHex = Replace$(DeHex, "%C0", Chr$(192))
    DeHex = Replace$(DeHex, "%C1", Chr$(193))
    DeHex = Replace$(DeHex, "%C2", Chr$(194))
    DeHex = Replace$(DeHex, "%C3", Chr$(195))
    DeHex = Replace$(DeHex, "%C4", Chr$(196))
    DeHex = Replace$(DeHex, "%C5", Chr$(197))
    DeHex = Replace$(DeHex, "%C6", Chr$(198))
    DeHex = Replace$(DeHex, "%C7", Chr$(199))
    DeHex = Replace$(DeHex, "%C8", Chr$(200))
    DeHex = Replace$(DeHex, "%C9", Chr$(201))
    DeHex = Replace$(DeHex, "%CA", Chr$(202))
    DeHex = Replace$(DeHex, "%CB", Chr$(203))
    DeHex = Replace$(DeHex, "%CC", Chr$(204))
    DeHex = Replace$(DeHex, "%CD", Chr$(205))
    DeHex = Replace$(DeHex, "%CE", Chr$(206))
    DeHex = Replace$(DeHex, "%CF", Chr$(207))
    DeHex = Replace$(DeHex, "%D0", Chr$(208))
    DeHex = Replace$(DeHex, "%D1", Chr$(209))
    DeHex = Replace$(DeHex, "%D2", Chr$(210))
    DeHex = Replace$(DeHex, "%D3", Chr$(211))
    DeHex = Replace$(DeHex, "%D4", Chr$(212))
    DeHex = Replace$(DeHex, "%D5", Chr$(213))
    DeHex = Replace$(DeHex, "%D6", Chr$(214))
    DeHex = Replace$(DeHex, "%D7", Chr$(215))
    DeHex = Replace$(DeHex, "%D8", Chr$(216))
    DeHex = Replace$(DeHex, "%D9", Chr$(217))
    DeHex = Replace$(DeHex, "%DA", Chr$(218))
    DeHex = Replace$(DeHex, "%DB", Chr$(219))
    DeHex = Replace$(DeHex, "%DC", Chr$(220))
    DeHex = Replace$(DeHex, "%DD", Chr$(221))
    DeHex = Replace$(DeHex, "%DE", Chr$(222))
    DeHex = Replace$(DeHex, "%DF", Chr$(223))
    DeHex = Replace$(DeHex, "%E0", Chr$(224))
    DeHex = Replace$(DeHex, "%E1", Chr$(225))
    DeHex = Replace$(DeHex, "%E2", Chr$(226))
    DeHex = Replace$(DeHex, "%E3", Chr$(227))
    DeHex = Replace$(DeHex, "%E4", Chr$(228))
    DeHex = Replace$(DeHex, "%E5", Chr$(229))
    DeHex = Replace$(DeHex, "%E6", Chr$(230))
    DeHex = Replace$(DeHex, "%E7", Chr$(231))
    DeHex = Replace$(DeHex, "%E8", Chr$(232))
    DeHex = Replace$(DeHex, "%E9", Chr$(233))
    DeHex = Replace$(DeHex, "%EA", Chr$(234))
    DeHex = Replace$(DeHex, "%EB", Chr$(235))
    DeHex = Replace$(DeHex, "%EC", Chr$(236))
    DeHex = Replace$(DeHex, "%ED", Chr$(237))
    DeHex = Replace$(DeHex, "%EE", Chr$(238))
    DeHex = Replace$(DeHex, "%EF", Chr$(239))
    DeHex = Replace$(DeHex, "%F0", Chr$(240))
    DeHex = Replace$(DeHex, "%F1", Chr$(241))
    DeHex = Replace$(DeHex, "%F2", Chr$(242))
    DeHex = Replace$(DeHex, "%F3", Chr$(243))
    DeHex = Replace$(DeHex, "%F4", Chr$(244))
    DeHex = Replace$(DeHex, "%F5", Chr$(245))
    DeHex = Replace$(DeHex, "%F6", Chr$(246))
    DeHex = Replace$(DeHex, "%F7", Chr$(247))
    DeHex = Replace$(DeHex, "%F8", Chr$(248))
    DeHex = Replace$(DeHex, "%F9", Chr$(249))
    DeHex = Replace$(DeHex, "%FA", Chr$(250))
    DeHex = Replace$(DeHex, "%FB", Chr$(251))
    DeHex = Replace$(DeHex, "%FC", Chr$(252))
    DeHex = Replace$(DeHex, "%FD", Chr$(253))
    DeHex = Replace$(DeHex, "%FE", Chr$(254))
    DeHex = Replace$(DeHex, "%FF", Chr$(255))
    
End Function