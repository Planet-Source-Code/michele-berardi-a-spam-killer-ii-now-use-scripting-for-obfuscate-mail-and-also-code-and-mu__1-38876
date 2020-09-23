VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHtmlObfuscator 
   Caption         =   "Spam Killer (C) 2002 Berardi Michele   -  http://web.tiscali.it/mberardi"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "SpamKiller.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8385
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopyScriptObfuscatorToClipBoard 
      Caption         =   "Copy Script To ClipBoard And Paste On Your Web Page!"
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   6000
      Width           =   4815
   End
   Begin VB.CommandButton cmdSaveToFIleScriptObfuscator 
      Caption         =   "SaveToFile Script"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8520
      TabIndex        =   23
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtScriptObfuscator 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      Top             =   5040
      Width           =   8295
   End
   Begin VB.CheckBox chkIsBinary 
      Caption         =   "BinaryCode"
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox chkGenerateScript 
      Caption         =   "Scripting"
      Height          =   255
      Left            =   8280
      TabIndex        =   19
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearResults 
      Caption         =   "Clear Results"
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   7560
      TabIndex        =   17
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdCleartxtMailClearForm 
      Caption         =   "Clear Input"
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   2640
      Width           =   615
   End
   Begin ComctlLib.ProgressBar objProgressBar 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   7920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLoadTextFile 
      Caption         =   "Load File (txt,Html,etc..)"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveToFileHtmlPage 
      Caption         =   "SaveToFile HtmlPage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8520
      TabIndex        =   13
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveToFIleObfuscatedText 
      Caption         =   "SaveToFile Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8520
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgFileSelector 
      Left            =   9120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkConsiderAsText 
      Caption         =   "PlainText"
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdMailObfuscatedCopyToClipBoard 
      Caption         =   "Copy Html Page To ClipBoard and Paste On Text, Save As .Htm"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   7920
      Width           =   4935
   End
   Begin VB.CommandButton cmdCopyMailObfuscatedToClipBoard 
      Caption         =   "Copy Html To ClipBoard And Paste On Your Web Page!"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox txtHtmlPage 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   6480
      Width           =   8295
   End
   Begin VB.TextBox txtMailObfuscated 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3240
      Width           =   8295
   End
   Begin VB.CommandButton cmdMailObfuscator 
      Caption         =   "Obfuscate"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   6135
   End
   Begin VB.TextBox txtMailClearForm 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "SpamKiller.frx":0E42
      Top             =   1320
      Width           =   9375
   End
   Begin VB.Label lblScriptObfuscator 
      Caption         =   "The Obfuscator Script (Pay Attention Scripting, but more efficient than Pure Html solution!):"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   6495
   End
   Begin VB.Label lblHtmlExamplePage 
      Caption         =   "Example of How To use Them  in Html Pages:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Spam Killer II (C) 2002 Berardi Michele"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label lblMailObfuscated 
      Caption         =   "The Obfuscated Mail Adress / Text / Code (Pure Html):"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label lblMailInClearForm 
      Caption         =   "Insert Below The Mail Address or , Text or Code to Obfuscate:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9375
   End
   Begin VB.Label lblHelp 
      Caption         =   $"SpamKiller.frx":0E53
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   8535
   End
End
Attribute VB_Name = "frmHtmlObfuscator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                                                             '
' Spam Killer II (C) 2002 Berardi Michele                     '
'                                                             '
' How to Reduce The Spam changing the codify of the html page '
' you can obfuscate mailto html field or text field!          '
'                                                             '
' New  feauture are:                                          '
'                                                             '
' JavaScript code generator for a better mail obfuscation..   '
'                                                             '
' Future Enanchemet are :                                     '
'                                                             '
' parse the entire page an choose the text part of the html   '
' page to be obfusched!                                       '
'                                                             '
' more script support for permit to many browser to support   '
' this kind of antispam purpose!                              '
'                                                             '
'                                                             '
' Berardi Michele                                             '
' Senior Developer                                            '
'                                                             '
' E-mail(s):                                                  '
' mfxaub@tin.it                                               '
' 03473192000@vizzavi.it                                      '
'                                                             '
' Web:                                                        '
'                                                             '
' Http://web.tiscali.it/mberardi                              '
'                                                             '
Public strMailClear As String, strAsciiMailObfuscator As String, strHexMailObfuscator As String, strLoadedIniScript As String, strLoadedIniMeta As String, strLoadedIniBody As String
Public chrConv, hexConv, strGreetings
Public strAsciiMailObfuscatorPreFix As String, strAsciiMailObfuscatorMidFix As String, strAsciiMailObfuscatorPstFix As String

    '                                                                                                                            '
    ' NOTE FOR VB OLDER THAN RELEASE 6.0:                                                                                        '
    '                                                                                                                            '
    ' the built in Replace() function                                                                                            '
    ' is not supported in version older than                                                                                     '
    ' release 6.0                                                                                                                '
    ' use the commented part of code instead                                                                                     '
    ' of the vb Replace() function.                                                                                              '
    '                                                                                                                            '
    ' Public Function Replace(strWhereReplace As String, strWithReplace As String, strWhatReplace As String) As String           '
    ' Dim i As Integer                                                                                                           '
    ' i = 1                                                                                                                      '
    ' Do While InStr(i, strWhereReplace, strWhatReplace, vbTextCompare) <> 0                                                     '
    ' Replace = Replace & Mid(strWhereReplace, i, InStr(i, strWhereReplace, strWhatReplace, vbTextCompare) - i) & strWithReplace '
    ' i = InStr(i, strWhereReplace, strWhatReplace, vbTextCompare) + Len(strWhatReplace)                                         '
    ' Loop                                                                                                                       '
    ' Replace = Replace & Right(strWhereReplace, Len(strWhereReplace) - i + 1)                                                   '
    ' End Function                                                                                                               '
    '                                                                                                                            '
    ' some newer control as "progress bar"                                                                                       '
    ' or these class of controls:                                                                                                '
    '                                                                                                                            '
    ' Microsoft Common Dialog Control 6.0                                                                                        '
    ' Microsoft Windows Common Controls 5.0                                                                                      '
    '                                                                                                                            '
    ' may not be in your programming                                                                                             '
    ' environment!                                                                                                               '
    '                                                                                                                            '
    ' Don't worry they don't affect                                                                                              '
    ' the main functionality                                                                                                     '
    ' of the program.                                                                                                            '
    '                                                                                                                            '

Public Function fnStrLoadFromFile(strFileName As String) As String

    Dim intNewFile As Integer
    Dim strBuffer As String

    On Error Resume Next

    intNewFile = FreeFile
    Open strFileName For Binary Access Read As #intNewFile
    strBuffer = Space(LOF(intNewFile))
    Get intNewFile, , strBuffer
    Close #intNewFile
    fnStrLoadFromFile = strBuffer

End Function

Public Function fnStrSaveToFile(strFileName As String, strBuffer As String) As Integer

    Dim intNewFile As Integer
    On Error Resume Next

    intNewFile = FreeFile
    Open strFileName For Binary Access Write As #intNewFile
    Put intNewFile, , strBuffer
    Close #intNewFile

End Function


Private Sub cmdClearAll_Click()

    objProgressBar.Value = objProgressBar.Min
    txtMailClearForm.Text = ""
    txtMailObfuscated.Text = ""
    txtScriptObfuscator.Text = ""
    txtHtmlPage = ""

End Sub

Private Sub cmdClearResults_Click()

    objProgressBar.Value = objProgressBar.Min
    txtMailObfuscated.Text = ""
    txtScriptObfuscator.Text = ""
    txtHtmlPage = ""

End Sub

Private Sub cmdCleartxtMailClearForm_Click()

    objProgressBar.Value = objProgressBar.Min
    txtMailClearForm.Text = ""

End Sub

Private Sub cmdCopyMailObfuscatedToClipBoard_Click()

    Clipboard.SetText txtMailObfuscated.Text

End Sub

Private Sub cmdCopyScriptObfuscatorToClipBoard_Click()
    Clipboard.SetText txtScriptObfuscator.Text
End Sub

Private Sub cmdLoadTextFile_Click()

    dlgFileSelector.Filter = "Text Files(*.txt;*.doc)|*.txt;*.doc|All Files(*.*)|*.*"
    dlgFileSelector.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    dlgFileSelector.DialogTitle = " [" & frmHtmlObfuscator.Caption & "] " & "Load File To Process!"
    dlgFileSelector.ShowOpen

    If dlgFileSelector.FileName <> "" Then
        objProgressBar.Value = objProgressBar.Min
        txtMailClearForm.Text = fnStrLoadFromFile(dlgFileSelector.FileName)
    End If

End Sub

Private Sub cmdMailObfuscatedCopyToClipBoard_Click()

    Clipboard.SetText txtHtmlPage.Text

End Sub

Private Sub cmdMailObfuscator_Click()

    txtMailObfuscated.Text = ""
    txtScriptObfuscator.Text = ""
    txtHtmlPage.Text = ""

    strMailClear = txtMailClearForm.Text
    strAsciiMailObfuscator = ""
    strHexMailObfuscator = ""

    '                                                '
    ' find & substitute vbCrLf with a single chr(13) '
    ' for better text formatting!                    '
    '                                                '

    If chkIsBinary.Value = False Then
        strMailClear = Replace(strMailClear, vbCrLf, Chr(13))
    End If

    ' write this in C or Asm and speed the cycle.. '

    objProgressBar.Enabled = True
    objProgressBar.Min = 1

    If Len(strMailClear) Then
        objProgressBar.Max = Len(strMailClear)
    End If

    objProgressBar.Value = objProgressBar.Min

    For X = 1 To Len(strMailClear)

        objProgressBar.Value = X

        chrConv = Asc(Mid(strMailClear, X, 1))

        hexConv = Hex(chrConv)

        If ((chrConv = 10) Or (chrConv = 13)) And (chkConsiderAsText.Value = 1) And (chkIsBinary.Value = 0) Then
            chrConv = "<br>"
        Else
            chrConv = "&#" & CStr(chrConv) & ";"
        End If

        strAsciiMailObfuscator = strAsciiMailObfuscator & chrConv
        strHexMailObfuscator = strHexMailObfuscator & "%" & hexConv

    Next

    If chkConsiderAsText.Value = 0 Then
        strAsciiMailObfuscatorPreFix = "<a href=""mailto:"
        strAsciiMailObfuscatorMidFix = """" & ">" & strAsciiMailObfuscator
        strAsciiMailObfuscatorPstFix = "</a>"
    Else
        strAsciiMailObfuscatorPreFix = "<p>"
        strAsciiMailObfuscatorMidFix = ""
        strAsciiMailObfuscatorPstFix = "</p>"
    End If

    If chkGenerateScript.Value = 1 Then

    End If

    txtMailObfuscated.Text = strAsciiMailObfuscatorPreFix & strAsciiMailObfuscator & strAsciiMailObfuscatorMidFix & strAsciiMailObfuscatorPstFix & vbCrLf


    Call SubGenerateHtmlPage

End Sub


Private Sub SubGenerateHtmlPage()

    txtHtmlPage.Text = ""

    '                            '
    ' Sample Html Page - BEGIN - '
    '                            '
    strGreetings = txtHtmlPage.Text & "<p> <a href=" & """" & "http://web.tiscali.it/mberardi/" & """" & ">Please Visit SpamKiller II Author HomePage!</a> </p>" & vbCrLf & vbCrLf
    '
    txtHtmlPage.Text = "<!DOCTYPE HTML PUBLIC " & """" & "-//W3C//DTD HTML 4.01 Transitional//EN" & """" & ">" & vbCrLf & vbCrLf
    ' <html> '
    txtHtmlPage.Text = txtHtmlPage.Text & "<html>" & vbCrLf & vbCrLf

    ' <head> '
    txtHtmlPage.Text = txtHtmlPage.Text & "<head>" & vbCrLf & vbCrLf

    ' <title> '
    txtHtmlPage.Text = txtHtmlPage.Text & "<title>Spam Killer 2002 Berardi Michele http://web.tiscali.it/mberardi </title>" & vbCrLf & vbCrLf

    '                                          '
    ' Using Scripting for elude the spam bots! '
    '                                          '
    If chkGenerateScript.Value = 1 Then

        strLoadedIniScript = fnStrLoadFromFile(App.Path & "\inc\scripts\default\script.htm")
        strLoadedIniScript = Replace(strLoadedIniScript, "[strSpamKillerCriptString]", strHexMailObfuscator)

        txtScriptObfuscator.Text = strLoadedIniScript

        txtHtmlPage.Text = txtHtmlPage.Text & strLoadedIniScript & vbCrLf & vbCrLf

    ' txtHtmlPage.Text = txtHtmlPage.Text & fnGenerateScript(strHexMailObfuscator) & vbCrLf '

    End If
    '                                          '
    ' Using Scripting for elude the spam bots! '
    '                                          '

    '        '
    ' <meta> '
    '        '

    strLoadedIniMeta = fnStrLoadFromFile(App.Path & "\inc\htm\meta.htm")
    txtHtmlPage.Text = txtHtmlPage.Text & strLoadedIniMeta & vbCrLf

    ' </meta> '

    ' </head> '
    txtHtmlPage.Text = txtHtmlPage.Text & "</head>" & vbCrLf & vbCrLf


    ' <body> '
    strLoadedIniBody = fnStrLoadFromFile(App.Path & "\inc\htm\body.htm")
    strLoadedIniBody = Replace(strLoadedIniBody, "[strSpamKillerBodyHere]", txtMailObfuscated.Text)
    strLoadedIniBody = Replace(strLoadedIniBody, "<body>", "<body>" & vbCrLf & vbCrLf & strGreetings)
    txtHtmlPage.Text = txtHtmlPage.Text & strLoadedIniBody & vbCrLf & vbCrLf
    ' </body> '


    ' </html> '
    txtHtmlPage.Text = txtHtmlPage.Text & "</html>" & vbCrLf & vbCrLf

    '                          '
    ' Sample Html Page - END - '
    '                          '

End Sub


Private Sub cmdSaveToFileHtmlPage_Click()

    dlgFileSelector.Filter = "Html Files(*.htm;*.html)|*.htm;*.html|All Files(*.*)|*.*"
    dlgFileSelector.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    dlgFileSelector.DialogTitle = " [" & frmHtmlObfuscator.Caption & "] " & "Save Dialog"
    dlgFileSelector.ShowSave

    If dlgFileSelector.FileName <> "" Then
        intResult = fnStrSaveToFile(dlgFileSelector.FileName, txtHtmlPage.Text)
    End If


End Sub

Private Sub cmdSaveToFIleObfuscatedText_Click()

    dlgFileSelector.Filter = "Html Files(*.htm;*.html)|*.htm;*.html|All Files(*.*)|*.*"
    dlgFileSelector.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    dlgFileSelector.DialogTitle = " [" & frmHtmlObfuscator.Caption & "] " & "Save Dialog"
    dlgFileSelector.ShowSave

    If dlgFileSelector.FileName <> "" Then
        intResult = fnStrSaveToFile(dlgFileSelector.FileName, txtMailObfuscated.Text)
    End If

End Sub

Private Sub cmdSaveToFIleScriptObfuscator_Click()

    dlgFileSelector.Filter = "Html Files(*.htm;*.html)|*.htm;*.html|All Files(*.*)|*.*"
    dlgFileSelector.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    dlgFileSelector.DialogTitle = " [" & frmHtmlObfuscator.Caption & "] " & "Save Dialog"
    dlgFileSelector.ShowSave

    If dlgFileSelector.FileName <> "" Then
        intResult = fnStrSaveToFile(dlgFileSelector.FileName, txtScriptObfuscator.Text)
    End If


End Sub

Private Sub txtMailClearForm_Change()

    objProgressBar.Value = objProgressBar.Min

End Sub

