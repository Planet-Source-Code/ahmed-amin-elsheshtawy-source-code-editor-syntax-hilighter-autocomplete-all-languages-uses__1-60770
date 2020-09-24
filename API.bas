Attribute VB_Name = "API"

Option Explicit

Global sci As Long

Public scint As ScintillaX

Public Const WM_NOTIFY = &H4E

'Estructuras que contienen la información del mensaje de Scintilla
Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Public Type SCNotification
    NotifyHeader As NMHDR
    position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    Message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    x As Long
    y As Long
End Type

' Tipos para búsqueda y recuperación de fragmentos de texto
Public Type CharacterRange
    cpMin As Long
    cpMax As Long
End Type

Public Type TextRange
    chrg As CharacterRange
    lpstrText As String
End Type

Type TextToFind
    chrg As CharacterRange
    lpstrText As String
    chrgText As CharacterRange
End Type

Private OldWindowProc As Long

Private Const GWL_WNDPROC = (-4)

'-------------------COLORES-------------------------------------------
Public Const BLACK = &H0
Public Const WHITE = &HFFFFFF
Public Const BLUE = &HC00000
Public Const RED = &HFF&
Public Const GREEN = &HC000&


Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_EX_CLIENTEDGE = &H200
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40


Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal m As Long, ByVal left As Long, ByVal top As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Any) As Long

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

' Cambia las características de un estilo
Public Sub Style(ByVal sty As Long, Optional ByVal ForeColor As Long = BLACK, _
                 Optional ByVal BackColor As Long = WHITE, Optional fnt As StdFont = Nothing, _
                 Optional ByVal EolFilled As Boolean, Optional ByVal SetVisible As Boolean = True, _
                 Optional ByVal SetCase As SC_CASE = SC_CASE_MIXED, Optional ByVal SetCharset As SC_CHARSET = SC_CHARSET_ANSI)
    Call Message(SCI_STYLESETFORE, sty, ForeColor)
    Call Message(SCI_STYLESETBACK, sty, BackColor)
    If Not fnt Is Nothing Then
        Call Message(SCI_STYLESETSIZE, sty, fnt.Size)
        Call Message(SCI_STYLESETFONT, sty, fnt.Name)
        Call Message(SCI_STYLESETBOLD, sty, fnt.Bold)
        Call Message(SCI_STYLESETITALIC, sty, fnt.Italic)
        Call Message(SCI_STYLESETUNDERLINE, sty, fnt.Underline)
    End If
    Call Message(SCI_STYLESETEOLFILLED, sty, EolFilled)
    Call Message(SCI_STYLESETVISIBLE, sty, SetVisible)
    Call Message(SCI_STYLESETCASE, sty, CLng(SetCase))
    Call Message(SCI_STYLESETCHARACTERSET, sty, CLng(SetCharset))
End Sub

' Envía un mensaje ('msg') a Scintilla
Public Sub Message(ByVal Msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal lParam = 0)
    If VarType(lParam) = vbString Then
        SendMessageString sci, Msg, IIf(wParam = 0, CLng(wParam), wParam), CStr(lParam)
    Else
        SendMessage sci, Msg, IIf(wParam = 0, CLng(wParam), wParam), IIf(lParam = 0, CLng(lParam), lParam)
    End If
End Sub

' Captura de mensajes
Private Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Hex$(Msg) = "4E" Then  ' El mensaje WM_NOTIFY
        Dim Notif As SCNotification
        Call CopyMemory(Notif, ByVal lParam, Len(Notif))
        scint.Raise_Events Notif
    Else
        NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
    End If
End Function

' Comineza la subclasificación de la ventana (ventana, control, etc.) 'hwnd'
Public Function Capture(hwnd As Long) As Boolean
Capture = True
OldWindowProc = GetWindowLong(hwnd, GWL_WNDPROC)
If OldWindowProc = 0 Then
    Capture = False
    Exit Function
End If
If SetWindowLong(hwnd, GWL_WNDPROC, AddressOf NewWindowProc) = 0 Then
    Capture = False
End If
End Function

' Termina la subclasificación
Public Function Release(hwnd As Long) As Boolean
Release = True
If SetWindowLong(hwnd, GWL_WNDPROC, OldWindowProc) = 0 Then
    Release = False
End If
End Function

' Reemplaza las ocurrencias correlativas 'find' + número (p.e.: %1, %2...) dentro de
' 'str' por los valores de args()
Public Function Replace2(ByVal str As String, ByVal find As String, ParamArray args()) As String
Dim n As Integer
    'Replace2 = str
    For n = 0 To UBound(args)
        Replace2 = Replace1(str, find & n, CStr(args(n)))
    Next n
End Function

' Replacement of the builtin 'Replace' VB6 function.
Public Function Replace1(str As String, ByVal find As String, ByVal rep As String)
Dim i As Integer, sBuf As String
    If InStr(str, find) = 0 Then
        Replace1 = str
    Else
        For i = 1 To Len(str)
            If Mid$(str, i, Len(find)) = find Then
                sBuf = sBuf & rep
                i = i + (Len(find) - 1)
            Else
                sBuf = sBuf & Mid$(str, i, 1)
            End If
        Next i
        str = sBuf
        Replace1 = sBuf
    End If
End Function

' Replacement of 'Split' VB6 function.
Public Function Split1(ByVal str As String, Optional separator As String = " ") As Variant
Dim i As Integer, n As Integer, l As Long, ls As Long
Dim sBuf As String
Dim vTmp() As Variant
    If InStr(str, separator) = 0 Then
        ReDim vTmp(0)
        vTmp(0) = str
    Else
        n = 0
        l = Len(separator)
        If Right(str, l) <> separator Then
            str = str & separator
        End If
        ls = Len(str)
        For i = 1 To ls
            If Mid$(str, i, l) = separator Then
                If i <= ls Then
                    ReDim Preserve vTmp(n)
                    vTmp(n) = sBuf
                    sBuf = ""
                    i = i + (l - 1)
                    n = n + 1
                End If
            Else
                sBuf = sBuf & Mid$(str, i, 1)
            End If
        Next i
    End If
    Split1 = vTmp
End Function
