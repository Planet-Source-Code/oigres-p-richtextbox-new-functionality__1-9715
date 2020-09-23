Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Public Const WM_USER = &H400

' /* New masks and effects -- a parenthesized asterisk indicates that
'   the data is stored by RichEdit2.0, but not displayed */

Public Const CFM_SMALLCAPS = &H40&                 ' /* (*)  */
Public Const CFM_ALLCAPS = &H80&                   ' /* (*)  */
Public Const CFM_HIDDEN = &H100&                   ' /* (*)  */
Public Const CFM_OUTLINE = &H200&                  ' /* (*)  */
Public Const CFM_SHADOW = &H400&                   ' /* (*)  */
Public Const CFM_EMBOSS = &H800&                   ' /* (*)  */
Public Const CFM_IMPRINT = &H1000&                 ' /* (*)  */
Public Const CFM_DISABLED = &H2000&
Public Const CFM_REVISED = &H4000&

Public Const CFM_BACKCOLOR = &H4000000
Public Const CFM_LCID = &H2000000
Public Const CFM_UNDERLINETYPE = &H800000         ' /* (*)  */
Public Const CFM_WEIGHT = &H400000
Public Const CFM_SPACING = &H200000               ' /* (*)  */
Public Const CFM_KERNING = &H100000               ' /* (*)  */
Public Const CFM_STYLE = &H80000                  ' /* (*)  */
Public Const CFM_ANIMATION = &H40000              ' /* (*)  */
Public Const CFM_REVAUTHOR = &H8000&

Public Const CFE_SUBSCRIPT = &H10000               ' /* Superscript and subscript are */
Public Const CFE_SUPERSCRIPT = &H20000            ' /*  mutually exclusive           */

Public Const CFM_SUBSCRIPT = CFE_SUBSCRIPT Or CFE_SUPERSCRIPT
Public Const CFM_SUPERSCRIPT = CFM_SUBSCRIPT

'Public Const CFM_EFFECTS2 = (CFM_EFFECTS Or CFM_DISABLED Or CFM_SMALLCAPS Or CFM_ALLCAPS _
'                    Or CFM_HIDDEN Or CFM_OUTLINE Or CFM_SHADOW Or CFM_EMBOSS _
'                    Or CFM_IMPRINT Or CFM_DISABLED Or CFM_REVISED _
'                    Or CFM_SUBSCRIPT Or CFM_SUPERSCRIPT Or CFM_BACKCOLOR)

'Public Const CFM_ALL2 = (CFM_ALL Or CFM_EFFECTS2 Or CFM_BACKCOLOR Or CFM_LCID _
'                    Or CFM_UNDERLINETYPE Or CFM_WEIGHT Or CFM_REVAUTHOR _
'                    Or CFM_SPACING Or CFM_KERNING Or CFM_STYLE Or CFM_ANIMATION)

Public Const CFE_SMALLCAPS = CFM_SMALLCAPS
Public Const CFE_ALLCAPS = CFM_ALLCAPS
Public Const CFE_HIDDEN = CFM_HIDDEN
Public Const CFE_OUTLINE = CFM_OUTLINE
Public Const CFE_SHADOW = CFM_SHADOW
Public Const CFE_EMBOSS = CFM_EMBOSS
Public Const CFE_IMPRINT = CFM_IMPRINT
Public Const CFE_DISABLED = CFM_DISABLED
Public Const CFE_REVISED = CFM_REVISED

' /* NOTE: CFE_AUTOCOLOR and CFE_AUTOBACKCOLOR correspond to CFM_COLOR and
'   CFM_BACKCOLOR, respectively, which control them */
Public Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR

' /* Underline types */
Public Const CFU_CF1UNDERLINE = &HFF&      ' /* map charformat's bit underline to CF2.*/
Public Const CFU_INVERT = &HFE&            ' /* For IME composition fake a selection.*/
Public Const CFU_UNDERLINEDOTTED = &H4&    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEDOUBLE = &H3&    ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEWORD = &H2&      ' /* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINE = &H1&
Public Const CFU_UNDERLINENONE = 0&

' /* CHARFORMAT masks */
Public Const CFM_BOLD = &H1
Public Const CFM_ITALIC = &H2
Public Const CFM_UNDERLINE = &H4
Public Const CFM_STRIKEOUT = &H8
Public Const CFM_PROTECTED = &H10
Public Const CFM_LINK = &H20&                  ' /* Exchange hyperlink extension */
Public Const CFM_SIZE = &H80000000
Public Const CFM_COLOR = &H40000000
Public Const CFM_FACE = &H20000000
Public Const CFM_OFFSET = &H10000000
Public Const CFM_CHARSET = &H8000000

' /* CHARFORMAT effects */
Public Const CFE_BOLD = &H1&
Public Const CFE_ITALIC = &H2&
Public Const CFE_UNDERLINE = &H4&
Public Const CFE_STRIKEOUT = &H8&
Public Const CFE_PROTECTED = &H10&
Public Const CFE_LINK = &H20&
Public Const CFE_AUTOCOLOR = &H40000000       ' /* NOTE: this corresponds to */
                                        ' /* CFM_COLOR, which controls it */
Public Const yHeightCharPtsMost = 1638&

' /* RichEdit messages */

' #ifndef WM_CONTEXTMENU
Public Const WM_CONTEXTMENU = &H7B&
' #End If

' #ifndef WM_PRINTCLIENT
Public Const WM_PRINTCLIENT = &H318&
' #End If

' #ifndef EM_GETLIMITTEXT
'public Const EM_GETLIMITTEXT = (WM_USER + 37)
' #End If

' #ifndef EM_POSFROMCHAR
'public Const EM_POSFROMCHAR = (WM_USER + 38)
'public Const EM_CHARFROMPOS = (WM_USER + 39)
' #End If

' #ifndef EM_SCROLLCARET
'public Const EM_SCROLLCARET = (WM_USER + 49)
' #End If
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_DISPLAYBAND = (WM_USER + 51)
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXLIMITTEXT = (WM_USER + 53)
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_FINDTEXT = (WM_USER + 56)
Public Const EM_FORMATRANGE = (WM_USER + 57)
Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_GETEVENTMASK = (WM_USER + 59)
Public Const EM_GETOLEINTERFACE = (WM_USER + 60)
Public Const EM_GETPARAFORMAT = (WM_USER + 61)
Public Const EM_GETSELTEXT = (WM_USER + 62)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_PASTESPECIAL = (WM_USER + 64)
Public Const EM_REQUESTRESIZE = (WM_USER + 65)
Public Const EM_SELECTIONTYPE = (WM_USER + 66)
Public Const EM_SETBKGNDCOLOR = (WM_USER + 67)
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const EM_SETOLECALLBACK = (WM_USER + 70)
Public Const EM_SETPARAFORMAT = (WM_USER + 71)
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_STREAMOUT = (WM_USER + 74)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_FINDWORDBREAK = (WM_USER + 76)
Public Const EM_SETOPTIONS = (WM_USER + 77)
Public Const EM_GETOPTIONS = (WM_USER + 78)
Public Const EM_FINDTEXTEX = (WM_USER + 79)
' #ifdef _WIN32
Public Const EM_GETWORDBREAKPROCEX = (WM_USER + 80)
Public Const EM_SETWORDBREAKPROCEX = (WM_USER + 81)
' #End If
'// message constants
Public Const EM_GETZOOM = (WM_USER + 224)
Public Const EM_SETZOOM = (WM_USER + 225)

' /* Richedit v2.0 messages */
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)

Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETTEXTMODE = (WM_USER + 90)

Public Const EM_FINDTEXTW = (WM_USER + 123)
Public Const EM_FINDTEXTEXW = (WM_USER + 124)

' /* enum for use with EM_GET/SETTEXTMODE */
' /* EM_SETCHARFORMAT wParam masks */
Public Const SCF_SELECTION = &H1&
Public Const SCF_WORD = &H2&
Public Const SCF_DEFAULT = &H0&            '// set the default charformat or paraformat
Public Const SCF_ALL = &H4&                '// not valid with SCF_SELECTION or SCF_WORD
Public Const SCF_USEUIRULES = &H8&         '// modifier for SCF_SELECTION; says that
                                   ' // the format came from a toolbar, etc. and
                                   ' // therefore UI formatting rules should be
                                   ' // used instead of strictly formatting the
                                   ' // selection.


Public Const LF_FACESIZE = 32
Public Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
    
    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type

