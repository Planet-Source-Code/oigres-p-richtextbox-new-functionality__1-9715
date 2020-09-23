VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Rtf - backcolour test"
   ClientHeight    =   3900
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboZoom 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   0
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cDlg1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "View and change  Rtf text here"
      Top             =   2880
      Width           =   4335
   End
   Begin RichTextLib.RichTextBox rtfText1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4260
      _Version        =   393217
      TextRTF         =   $"rtfBackColourTestForm1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
         Begin VB.Menu mnuFontColour 
            Caption         =   "&FontColour"
         End
         Begin VB.Menu mnuFontBackColour 
            Caption         =   "&FontBackColour"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewRtf 
         Caption         =   "&View rtf"
      End
      Begin VB.Menu mnuSetRtf 
         Caption         =   "&setRtf"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuMe 
         Caption         =   "&ME!"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------IMPORTANT INFORMATION-------------------------------
'This sample requires the Riched20.dll (version 3) look at the file properties
'This also requires Riched32.dll (5.00.2008.1)
'and probably the Richx32.ocx control (I've got version 6.00.8418)
'Based in part on the work of Steve McMahon (www.vbaccelarator.com)
'Font used in sample is Times New Roman 8pt (scales better that MS SansSerif)
'Delete/overwrite the riched20 and riched32.dll files
'You need to have vb6/5 not running to replace these files.
'By oigres P (Sergio Perciballi) Email:oigres@postmaster.co.uk
'new richtextbox and dll files can be found at www.vbaccelerator.com
'

Dim charf As CHARFORMAT2 'character format type for extended information

Private Sub Form_Load()
    'MsgBox LenB(charf)
    'MsgBox VarPtr(charf) & ":"
    Dim lIdx
    ''charf.crBackColor = &HFF& 'initially red background
    charf.dwMask = CFM_BACKCOLOR
    charf.cbSize = LenB(charf) 'setup the size of the character format
    'setup zoom combo box - most from Steve McMahon
    cboZoom.AddItem "10%" '1:10
    cboZoom.ItemData(cboZoom.NewIndex) = 1 * &H10000 + 10
    cboZoom.AddItem "25%" '1:4
    cboZoom.ItemData(cboZoom.NewIndex) = 1 * &H10000 + 4
    cboZoom.AddItem "50%" '1:2
    cboZoom.ItemData(cboZoom.NewIndex) = 1 * &H10000 + 2
    cboZoom.AddItem "75%" '3:4'
    cboZoom.ItemData(cboZoom.NewIndex) = 3 * &H10000 + 4
    cboZoom.AddItem "80%" '4:5
    cboZoom.ItemData(cboZoom.NewIndex) = 4 * &H10000 + 5
    cboZoom.AddItem "90%" '9:10
    cboZoom.ItemData(cboZoom.NewIndex) = 9 * &H10000 + 10
    cboZoom.AddItem "100%"
    lIdx = cboZoom.NewIndex '1:1
    cboZoom.ItemData(cboZoom.NewIndex) = 1 * &H10000 + 1
    cboZoom.AddItem "150%" '3:2
    cboZoom.ItemData(cboZoom.NewIndex) = 3 * &H10000 + 2
    cboZoom.AddItem "200%" '2:1
    cboZoom.ItemData(cboZoom.NewIndex) = 2 * &H10000 + 1
    cboZoom.AddItem "250%" '5:2
    cboZoom.ItemData(cboZoom.NewIndex) = 5 * &H10000 + 2
    cboZoom.AddItem "300%"  '3:1'
    cboZoom.ItemData(cboZoom.NewIndex) = 3 * &H10000 + 1
    cboZoom.AddItem "350%"  '7:2'
    cboZoom.ItemData(cboZoom.NewIndex) = 7 * &H10000 + 2
    cboZoom.AddItem "400%"  '4:1'
    cboZoom.ItemData(cboZoom.NewIndex) = 4 * &H10000 + 1
    cboZoom.AddItem "450%"  '9:2'
    cboZoom.ItemData(cboZoom.NewIndex) = 9 * &H10000 + 2
    cboZoom.AddItem "500%"  '5:1'
    cboZoom.ItemData(cboZoom.NewIndex) = 5 * &H10000 + 1
    cboZoom.ListIndex = lIdx
    Form1.Show
    'had to show the form first or else invalid procedure call or argument
    cboZoom.SetFocus

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFontBackColour_Click()
    Dim ret As Long
    On Error GoTo error_cancel
    cDlg1.Flags = cdlCCRGBInit '+ cdlCancel
    cDlg1.Color = vbRed
    cDlg1.CancelError = True
    cDlg1.ShowColor
    'MsgBox "Colour = " & cDlg1.Color
    'set the font colour
    '''rtfText1.SelColor = cDlg1.Color
    charf.crBackColor = cDlg1.Color
    ret = SendMessageLong(rtfText1.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(charf))
    If ret = 0 Then
    MsgBox "You probably have the wrong  files ->" & vbCrLf _
    & "Riched20.dll Version 3.0 (File version 5.30.22.2300)" & vbCrLf _
    & "Riched32.dll (file version 5.00.2008.1)" & vbCrLf _
    & "Richtx32.ocx not so critical (my version is 6.00.8418)" & vbCrLf _
    & "Get the files from www.vbacelerator.com , delete/backup old system files and replace"
    
    ''MsgBox "ret= " & ret
    End If
    Exit Sub
error_cancel:
    ''MsgBox "cancel "
End Sub

Private Sub mnuFontColour_Click()
    On Error GoTo error_cancel
    cDlg1.Flags = cdlCCRGBInit '+ cdlCancel
    cDlg1.Color = vbRed
    cDlg1.CancelError = True
    cDlg1.ShowColor
    'MsgBox "Colour = " & cDlg1.Color
    'set the font colour
    rtfText1.SelColor = cDlg1.Color

    Exit Sub
error_cancel:
    ''MsgBox "cancel "
End Sub

Private Sub mnuMe_Click()
    MsgBox "Sample by oigres P " & vbCrLf & "Email: oigres@postmaster.co.uk", , "            About"
End Sub

Private Sub mnuSetRtf_Click()
    rtfText1.TextRTF = Text1.Text
    'reset zoom ratio
    cboZoom_Click
End Sub

Private Sub mnuViewRtf_Click()
    Text1.Text = rtfText1.TextRTF
End Sub
Private Sub cboZoom_Click()
Dim lND As Long
Dim lNum As Long
Dim lDen As Long
Dim dummyinum As Long, dummydenom As Long
   If cboZoom.ListIndex > -1 Then
      lND = cboZoom.ItemData(cboZoom.ListIndex)
      lNum = lND \ &H10000
      lDen = lND And &H7FFF&
      SetZoom rtfText1.hWnd, lNum, lDen
    ''GetZoom rtfText1.hWnd, dummyinum, dummydenom
   End If
End Sub

Public Sub SetZoom(ByVal hWndRtf As Long, ByVal lNumerator As Long, ByVal lDenominator As Long)
Dim lR As Long
   If lNumerator > 64 Or lDenominator > 64 Or lNumerator < 0 Or lDenominator < 0 Then
      Err.Raise 27110, App.EXEName & ".mRichEdit30", "Numerator and Denominator must be between 1 and 64"
   End If
   lR = SendMessageLong(hWndRtf, EM_SETZOOM, lNumerator, lDenominator)
End Sub
Public Sub GetZoom(ByVal hWndRtf As Long, ByRef lNumerator As Long, ByRef lDenominator As Long)
   SendMessageRef hWndRtf, EM_GETZOOM, lNumerator, lDenominator
    ''Label1.Caption = "Zoom Ratio " & lNumerator & ":" & lDenominator
End Sub
