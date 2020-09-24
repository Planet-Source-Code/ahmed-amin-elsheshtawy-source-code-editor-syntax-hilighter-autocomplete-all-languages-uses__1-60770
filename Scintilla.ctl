VERSION 5.00
Begin VB.UserControl ScintillaX 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "Scintilla.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Scintilla.ctx":0011
End
Attribute VB_Name = "ScintillaX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Enum EOL                                ' Distintos finales de línea
    SC_EOL_CRLF = 0                     ' CR + LF
    SC_EOL_CR = 1                       ' CR
    SC_EOL_LF = 2                       ' LF
End Enum

Enum SCWS                               ' Visualización de caracteres invisibles
    SCWS_INVISIBLE = 0                  ' No se ven
    SCWS_VISIBLEALWAYS = 1              ' Siempre se ven
    SCWS_VISIBLEAFTERINDENT = 2         ' Se ven después de la indentación
End Enum

Private lastTotal As Long               ' Almacenamiento temporal del número de líneas

' Propiedades por defecto
Const m_def_LineNumbers = False
Const m_def_Language = "Python"
Const m_def_CallTipBack = WHITE
Const m_def_EOL = SC_EOL_CRLF
Const m_def_SCWS = SCWS_INVISIBLE
Const m_def_SepChar = " "
Const m_def_ConfFile = "lexer.xml"

' Licencia
Const m_def_Text = "This control is a wrapper of the Scintilla control. " & vbCrLf & _
                   "Scintilla source code, news and updates at http://www.scintilla.org." & vbCrLf & vbCrLf & _
                   "VB Control placed in public domain by J. M. Rodríguez, 2002.  Share and enjoy!"
                

' Almacenamiento de propiedades
Private m_LineNumbers As Boolean        ' Mostrar o no los números de línea y su margen
Private m_Text As String                ' El texto que muestra Scintilla
Private m_Language As String            ' El lenguaje vigente
Private m_CallTipBack As OLE_COLOR      ' El color de fondo de los globos de ayuda
Private m_EOL As EOL                    ' El final de línea vigente
Private EOfL As Integer                 ' Final de línea devuelto por Scintilla
Private m_SMCase As Boolean             ' Búsqueda; buscar coincidiendo mays/mins
Private m_SWWord As Boolean             ' Búsqueda; buscar palabras completas
Private m_SWStart As Boolean            ' Búsqueda; buscar al comienzo de palabra
Private m_SRegExp As Boolean            ' Búsqueda; el patrón es una expresión regular
Private SMCase As Long                  ' "Flags" de búsqueda para pasar a Scintilla.
Private SWWord As Long                  ' "Flags" de búsqueda para pasar a Scintilla.
Private SWStart As Long                 ' "Flags" de búsqueda para pasar a Scintilla.
Private SRegExp As Long                 ' "Flags" de búsqueda para pasar a Scintilla.
Private m_SCWS As SCWS                  ' Espacios visibles o no
Private m_ViewEOL As Boolean            ' Finales de línea visibles o no
Private m_SepChar As String             ' El carácter de separación para las listas automáticas
Private m_AutoChide As Boolean          ' La lista automática se oculta si no hay coincidencias
Private m_MatchBraces As Boolean        ' Se realzan los paréntesis o no
Private m_ConfFile As String            ' Fichero de configuración
Private m_HscrollBar As Boolean         ' Visualización de la barra de desplazamiento horizontal
Private m_IndGuides As Boolean          ' Visualización de las guías de indentación

' Eventos
Public Event CharAdded(character As String, word As String)
Public Event Modified(tx As String)
Public Event UpdateUI()

' Si hay un rectángulo de ayuda contextual visible -*
Public Property Get CallTipActive() As Boolean
    CallTipActive = SendMessage(sci, SCI_CALLTIPACTIVE, CLng(0), CLng(0))
End Property

' Ocultar la ayuda contextual -*
Public Sub CallTipCancel()
    Call Message(SCI_CALLTIPCANCEL)
End Sub

' Búsqueda en el texto -*
Public Function find(txttofind As String, Optional ByVal findinrng As Boolean) As Long
Dim targetstart As Long, targetend As Long, pos As Long
    ' Propiedades de la búsqueda
    Call Message(SCI_SETSEARCHFLAGS, SMCase Or SWWord Or SWStart Or SRegExp)
    If findinrng Then                   ' Buscamos en la selección
        targetstart = SendMessage(sci, SCI_GETSELECTIONSTART, CLng(0), CLng(0))
        targetend = SendMessage(sci, SCI_GETSELECTIONEND, CLng(0), CLng(0))
    Else
        targetstart = 0
        targetend = Len(Text)
    End If
    ' Creamos una región de búsqueda (que puede ser el texto completo)
    Call Message(SCI_SETTARGETSTART, targetstart)
    Call Message(SCI_SETTARGETEND, targetend)
    find = SendMessageString(sci, SCI_SEARCHINTARGET, Len(txttofind), txttofind)
    ' Seleccionamos lo que se ha encontrado
    If find > -1 Then
        targetstart = SendMessage(sci, SCI_GETTARGETSTART, CLng(0), CLng(0))
        targetend = SendMessage(sci, SCI_GETTARGETEND, CLng(0), CLng(0))
        Call Message(SCI_SETSEL, targetstart, targetend)
    End If
End Function

' Búsqueda mayúsculas/minúsculas -*
Public Property Get SearchMatchCase() As Boolean
Attribute SearchMatchCase.VB_MemberFlags = "400"
    SearchMatchCase = m_SMCase
End Property

Public Property Let SearchMatchCase(vNewValue As Boolean)
    m_SMCase = vNewValue
    If vNewValue Then SMCase = SCFIND_MATCHCASE Else SMCase = 0
    PropertyChanged "SearchMatchCase"
End Property

' Búsqueda por palabras completas -*
Public Property Get SearchWholeWord() As Boolean
Attribute SearchWholeWord.VB_MemberFlags = "400"
    SearchWholeWord = m_SWWord
End Property

Public Property Let SearchWholeWord(vNewValue As Boolean)
    m_SWWord = vNewValue
    If vNewValue Then SWWord = SCFIND_WHOLEWORD Else SWWord = 0
    PropertyChanged "SearchWholeWord"
End Property

' Búsqueda al inicio de palabra -*
Public Property Get SearchWordStart() As Boolean
Attribute SearchWordStart.VB_MemberFlags = "400"
    SearchWordStart = m_SWStart
End Property

Public Property Let SearchWordStart(vNewValue As Boolean)
    m_SWStart = vNewValue
    If vNewValue Then SWStart = SCFIND_WORDSTART Else SWStart = 0
    PropertyChanged "SearchWordStart"
End Property

' Patrón de búsqueda -> expresión regular -*
Public Property Get SearchRegExp() As Boolean
Attribute SearchRegExp.VB_MemberFlags = "400"
    SearchRegExp = m_SRegExp
End Property

Public Property Let SearchRegExp(vNewValue As Boolean)
    m_SRegExp = vNewValue
    If vNewValue Then SRegExp = SCFIND_REGEXP Else SRegExp = 0
    PropertyChanged "SearchRegExp"
End Property

' Espacios blancos visibles o no -*
Public Property Get WhiteSpaceVisible() As SCWS
Attribute WhiteSpaceVisible.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    WhiteSpaceVisible = m_SCWS
End Property

Public Property Let WhiteSpaceVisible(vNewValue As SCWS)
    m_SCWS = vNewValue
    Call Message(SCI_SETVIEWWS, m_SCWS)
    PropertyChanged "WhiteSpcVisible"
End Property
 
' Finales de línea visibles o no -*
Public Property Get EOLVisible() As Boolean
Attribute EOLVisible.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    EOLVisible = m_ViewEOL
End Property

Public Property Let EOLVisible(vNewValue As Boolean)
    m_ViewEOL = vNewValue
    Call Message(SCI_SETVIEWEOL, m_ViewEOL)
    PropertyChanged "EOLVisible"
End Property

Private Sub UserControl_Initialize()
    ' Cargamos Scintilla
    LoadLibrary ("SciLexer.DLL")
    sci = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "Scint.ocx", WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, UserControl.hwnd, 0, App.hInstance, 0)
    ' Referenciamos el control para llevar a cabo la subclasificación
    Set scint = Me
    ' Subclasificamos el control
    If Capture(UserControl.hwnd) = False Then
        Err.Raise 516, "Control_Initialize", LoadResString(516)
    End If
    ' Sólo nos notificará la inserción de texto. TODO: notificación a elegir
    Call Message(SCI_SETMODEVENTMASK, SC_MOD_INSERTTEXT)
End Sub

Private Sub UserControl_Resize()
    '
    SetWindowPos sci, 0, 0, 0, UserControl.width / 15, UserControl.height / 15, 0
End Sub

' El color de fondo -*
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call ChangeDefault
    PropertyChanged "BackColor"
End Property

' El color de la letra -*
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Call ChangeDefault
    PropertyChanged "ForeColor"
End Property

' El lenguaje vigente
Public Property Get language() As String
Attribute language.VB_MemberFlags = "400"
    language = m_Language
End Property

Public Property Let language(New_Language As String)
    m_Language = New_Language
    Call ChangeLang(m_Language)
    PropertyChanged "Language"
End Property

' Iniciar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Text = m_def_Text
    m_LineNumbers = m_def_LineNumbers
    m_Language = m_def_Language
    m_SCWS = m_def_SCWS
    m_AutoChide = True
    m_SepChar = m_def_SepChar
    m_ConfFile = m_def_ConfFile
    Call ChangeDefault
End Sub

' Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    UserControl.BackColor = PropBag.ReadProperty("BackColor", WHITE)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", BLACK)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_LineNumbers = PropBag.ReadProperty("LineNumbers", m_def_LineNumbers)
    m_Language = PropBag.ReadProperty("Language", m_def_Language)
    m_CallTipBack = PropBag.ReadProperty("CallTipBackcolor", m_def_CallTipBack)
    m_EOL = PropBag.ReadProperty("EndOfLine", m_def_EOL)
    m_SCWS = PropBag.ReadProperty("WhiteSpcVisible", m_def_EOL)
    m_SepChar = PropBag.ReadProperty("Separator", " ")
    m_ViewEOL = PropBag.ReadProperty("EOLVisible", False)
    m_SMCase = PropBag.ReadProperty("SearchMatchCase", False)
    m_AutoChide = PropBag.ReadProperty("AutoChide", True)
    m_MatchBraces = PropBag.ReadProperty("MatchBraces", False)
    m_ConfFile = PropBag.ReadProperty("ConfFile", m_def_ConfFile)
    m_HscrollBar = PropBag.ReadProperty("HScroll", False)
    m_IndGuides = PropBag.ReadProperty("IndGuides", False)
    If m_SMCase Then SMCase = SCFIND_MATCHCASE Else SMCase = 0
    m_SWWord = PropBag.ReadProperty("SearchWholeWord", False)
    If m_SWWord Then SWWord = SCFIND_WHOLEWORD Else SWWord = 0
    m_SWStart = PropBag.ReadProperty("SearchWordStart", False)
    If m_SWStart Then SWStart = SCFIND_WORDSTART Else SWStart = 0
    m_SRegExp = PropBag.ReadProperty("SearchRegExp", False)
    If m_SRegExp Then SRegExp = SCFIND_MATCHCASE Else SRegExp = 0
    Select Case m_EOL
        Case SC_EOL_CRLF
            EOfL = 13
        Case SC_EOL_CR
            EOfL = 13
        Case SC_EOL_LF
            EOfL = 10
    End Select
On Error GoTo 0
    Call ChangeDefault
End Sub

Private Sub UserControl_Terminate()
    Call Release(UserControl.hwnd)
End Sub

' Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, WHITE)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, BLACK)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Language", m_Language, m_def_Language)
    Call PropBag.WriteProperty("LineNumbers", m_LineNumbers, m_def_LineNumbers)
    Call PropBag.WriteProperty("CallTipBackColor", m_CallTipBack, m_def_CallTipBack)
    Call PropBag.WriteProperty("EndOfLine", m_EOL, m_def_EOL)
    Call PropBag.WriteProperty("SearchMatchCase", m_SMCase, False)
    Call PropBag.WriteProperty("SearchWholeWord", m_SWWord, False)
    Call PropBag.WriteProperty("SearchWordStart", m_SWStart, False)
    Call PropBag.WriteProperty("SearchRegExp", m_SRegExp, False)
    Call PropBag.WriteProperty("WhiteSpcVisible", m_SCWS, m_def_SCWS)
    Call PropBag.WriteProperty("EOLVisible", m_ViewEOL, False)
    Call PropBag.WriteProperty("Separator", m_SepChar, " ")
    Call PropBag.WriteProperty("AutoChide", m_AutoChide, True)
    Call PropBag.WriteProperty("MatchBraces", m_MatchBraces, False)
    Call PropBag.WriteProperty("ConfFile", m_ConfFile, m_def_ConfFile)
    Call PropBag.WriteProperty("HScroll", m_HscrollBar, False)
    Call PropBag.WriteProperty("IndGuides", m_IndGuides, False)
End Sub

' El texto que muestra Scintilla -*
Public Property Get Text() As String
Attribute Text.VB_Description = "Devuelve o establece el texto del control / Gets or sets control's text"
Dim numChar As Long
Dim Txt As String
    numChar = SendMessage(sci, SCI_GETLENGTH, 0, 0) + 1
    Txt = String(numChar, "0")  ' Tenemos que iniciar la cadena
    SendMessageString sci, SCI_GETTEXT, numChar, Txt
    Text = Txt
End Property

Public Property Let Text(ByVal New_Text As String)
    Call Message(SCI_SETTEXT, 0, New_Text)
    m_Text = New_Text
    PropertyChanged "Text"
End Property

' Mostrar números de línea o no -*
Public Property Get LineNumbers() As Boolean
Attribute LineNumbers.VB_Description = "Muestra u oculta el margen con números de línea / Show or hides line numbers margin"
Attribute LineNumbers.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    LineNumbers = m_LineNumbers
End Property

Public Property Let LineNumbers(ByVal New_LineNumbers As Boolean)
    m_LineNumbers = New_LineNumbers
    If New_LineNumbers Then
        Call ChangeMarginWidth(0)
        Call Message(SCI_SETMARGINTYPEN, 0, SC_MARGIN_NUMBER)
    Else
        Call Message(SCI_SETMARGINWIDTHN, 0, 0)    ' Poner el ancho del margen a cero lo elimina
    End If
    PropertyChanged "LineNumbers"
End Property

' Cambiar el ancho de un margen
Private Sub ChangeMarginWidth(ByVal marg As Long)
Dim margin_width As Long
    ' El ancho del margen
    margin_width = Len(CStr(Me.TotalLines)) * 10
    Call Message(SCI_SETMARGINWIDTHN, marg, IIf(margin_width < 15, 15, margin_width))
End Sub

' Cambiar las propiedades del control
Private Sub ChangeDefault()
    Me.Text = m_Text
    Me.CallTipBackColor = m_CallTipBack
    Me.LineNumbers = m_LineNumbers
    Me.EOLVisible = m_ViewEOL
    Me.WhiteSpaceVisible = m_SCWS
    Me.AutoCSeparatorChar = m_SepChar
    Me.AutoCHideIfNoMatch = m_AutoChide
    Me.HScrollBarVisible = m_HscrollBar
    Call ChangeDefaultStyle
End Sub

' Rutina de cambio de estilos predefinidos:  números de línea, concordancia de paréntesis,
' caracteres de control y líneas de indentación; excepto estilo por defecto
Private Sub ChangeDefStyles(Optional xmldoc As DOMDocument)
Dim nd As IXMLDOMElement
Dim n As Integer, m As Integer
Dim code As Long, fc As Long, bc As Long
Dim ef As Long, ev As Long, ca As Long, ch As Long
Dim words As String
Dim fnt As StdFont
Dim cleanup As Boolean
    If xmldoc Is Nothing Then
        Set xmldoc = New DOMDocument
        If Not xmldoc.Load(Me.ConfFile) Then
            Set xmldoc = Nothing
            Err.Raise 517, "ChangeDefStyles", LoadResString(517) & ": " & xmldoc.parseError.reason _
                                & " " & xmldoc.parseError.srcText
            Exit Sub
        Else
            cleanup = True
            xmldoc.validateOnParse = True
        End If
    End If
    Set fnt = New StdFont
    Set nd = SearchNode(xmldoc, "defstyles")
    For n = 0 To nd.childNodes.length - 1
        code = nd.childNodes(n).Attributes.getNamedItem("code").nodeValue
        If code <> 32 Then      ' Exceptuando el estilo global por defecto
            For m = 0 To nd.childNodes(n).childNodes.length - 1
                With nd.childNodes(n).childNodes(m)
                    Select Case .nodeName
                        Case "font"
                            fnt.Name = .Attributes.getNamedItem("name").nodeValue
                            fnt.Size = val(.Attributes.getNamedItem("size").nodeValue)
                            fnt.Bold = .Attributes.getNamedItem("bold").nodeValue
                            fnt.Italic = .Attributes.getNamedItem("italic").nodeValue
                            fnt.Underline = .Attributes.getNamedItem("underline").nodeValue
                        Case "forecolor"
                            fc = .Attributes.getNamedItem("value").nodeValue
                        Case "backcolor"
                            bc = .Attributes.getNamedItem("value").nodeValue
                        Case "eolfilled"
                            ef = IIf(.Attributes.getNamedItem("value").nodeValue = "true", 1, 0)
                        Case "visible"
                            ev = IIf(.Attributes.getNamedItem("value").nodeValue = "true", 1, 0)
                        Case "case"
                            ca = CLng(.Attributes.getNamedItem("value").nodeValue)
                        Case "charset"
                            ch = CLng(.Attributes.getNamedItem("value").nodeValue)
                    End Select
                End With
            Next m
            Style code, fc, bc, fnt, ef, ev, ca, ch
        End If
    Next n
    Set nd = Nothing
    If cleanup Then Set xmldoc = Nothing
    Set fnt = Nothing
End Sub

' Cambia el estilo por defecto
Sub ChangeDefaultStyle(Optional xmldoc As DOMDocument)
Attribute ChangeDefaultStyle.VB_MemberFlags = "40"
Dim nd As IXMLDOMElement
Dim m As Integer
Dim code As Long, fc As Long, bc As Long
Dim ef As Long, ev As Long, ca As Long, ch As Long
Dim fnt As StdFont
Dim cleanup As Boolean
    If xmldoc Is Nothing Then
        Set xmldoc = New DOMDocument
        If Not xmldoc.Load(Me.ConfFile) Then
            Set xmldoc = Nothing
            Err.Raise 517, "ChangeDefStyles", LoadResString(517) & ": " & xmldoc.parseError.reason _
                                & " " & xmldoc.parseError.srcText
            Exit Sub
        Else
            cleanup = True
            xmldoc.validateOnParse = True
        End If
    End If
    Call Message(SCI_STYLERESETDEFAULT)
    Set fnt = New StdFont
    Set nd = SearchNode(xmldoc, DEFCONST & "32" & SQLEND)
    code = 32
    For m = 0 To nd.childNodes.length - 1
        With nd.childNodes(m)
            Select Case .nodeName
                Case "font"
                    fnt.Name = .Attributes.getNamedItem("name").nodeValue
                    fnt.Size = val(.Attributes.getNamedItem("size").nodeValue)
                    fnt.Bold = .Attributes.getNamedItem("bold").nodeValue
                    fnt.Italic = .Attributes.getNamedItem("italic").nodeValue
                    fnt.Underline = .Attributes.getNamedItem("underline").nodeValue
                Case "forecolor"
                    fc = .Attributes.getNamedItem("value").nodeValue
                Case "backcolor"
                    bc = .Attributes.getNamedItem("value").nodeValue
                Case "eolfilled"
                    ef = IIf(.Attributes.getNamedItem("value").nodeValue = "true", 1, 0)
                Case "visible"
                    ev = IIf(.Attributes.getNamedItem("value").nodeValue = "true", 1, 0)
                Case "case"
                    ca = CLng(.Attributes.getNamedItem("value").nodeValue)
                Case "charset"
                    ch = CLng(.Attributes.getNamedItem("value").nodeValue)
            End Select
        End With
    Next m
    Style code, fc, bc, fnt, ef, ev, ca, ch
    Set nd = Nothing
    If cleanup Then Set xmldoc = Nothing
    Set fnt = Nothing
End Sub

' Mostrar ayuda contextual:
'   - tip: cadena a mostrar
'   - hltstart: el comienzo de la parte de 'tip' que se ha de mostrar en negrita.
'   - hltend: el final de la parte de 'tip' que se muestra en negrita. -*
Public Sub ShowCallTip(tip As String, Optional hltstart As Long, Optional hltend As Long)
    Dim pos As Long
    pos = SendMessage(sci, SCI_GETCURRENTPOS, 0, 0) 'Obtenemos la posición del cursor
    Call Message(SCI_CALLTIPSHOW, pos, tip)
    If hltstart And hltend Then
        Call Message(SCI_CALLTIPSETHLT, hltstart, hltend)
    End If
End Sub

' Color de fondo de la ayuda contextual -*
Public Property Get CallTipBackColor() As OLE_COLOR
Attribute CallTipBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    CallTipBackColor = m_CallTipBack
End Property

Public Property Let CallTipBackColor(vNewBack As OLE_COLOR)
    m_CallTipBack = vNewBack
    PropertyChanged "CallTipBackColor"
End Property

' Total de líneas del control (sólo lectura). -*
Public Property Get TotalLines() As Long
Attribute TotalLines.VB_MemberFlags = "400"
    TotalLines = SendMessage(sci, SCI_GETLINECOUNT, 0, 0)
End Property

' El tipo de final de línea (CRLF, LF, CR) -*
Property Get EndOfLine() As EOL
Attribute EndOfLine.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    EndOfLine = m_EOL
End Property

Property Let EndOfLine(vNewValue As EOL)
    Call Message(SCI_SETEOLMODE, vNewValue)
    m_EOL = vNewValue
    PropertyChanged "EndOfLine"
End Property

' Gestor de eventos
Friend Sub Raise_Events(Notif As SCNotification)
    Select Case Notif.NotifyHeader.code
        Case SCN_STYLENEEDED
             'TODO
        Case SCN_CHARADDED
            RaiseEvent CharAdded(Chr(Notif.ch), GetWord())
        Case SCN_SAVEPOINTREACHED
            'TODO
        Case SCN_SAVEPOINTLEFT
            'TODO
        Case SCN_MODIFYATTEMPTRO
            'TODO
        Case SCN_DOUBLECLICK
            'TODO
        Case SCN_UPDATEUI
            If m_MatchBraces Then
                Dim pos As Long, pos2 As Long
                pos2 = INVALID_POSITION
                ' Miramos la posición actual
                If IsBrace(CharAtPos(Me.CurPosition)) Then
                    pos2 = Me.CurPosition
                ' Y también la anterior
                ElseIf IsBrace(CharAtPos(Me.CurPosition - 1)) Then
                    pos2 = Me.CurPosition - 1
                End If
                If pos2 <> INVALID_POSITION Then
                    pos = SendMessage(sci, SCI_BRACEMATCH, pos2, CLng(0))
                    If pos = INVALID_POSITION Then
                        ' No hay paréntesis correspondiente
                        Call Message(SCI_BRACEBADLIGHT, pos2)
                    Else
                        ' Realzamos el paréntesis correspondiente
                        Call Message(SCI_BRACEHIGHLIGHT, pos, pos2)
                        ' También la guía si es necesario
                        If m_IndGuides Then
                            Call Message(SCI_SETHIGHLIGHTGUIDE, Me.Column)
                            Debug.Print "Fací neró sooos"
                        End If
                    End If
                Else
                    ' Esto quita cualquier realce de paréntesis
                    Call Message(SCI_BRACEHIGHLIGHT, INVALID_POSITION, INVALID_POSITION)
                End If
            End If
            ' Cambiamos el ancho del margen de números si es necesario
            If (Len(CStr(lastTotal)) <> Len(CStr(Me.TotalLines))) And m_LineNumbers Then ChangeMarginWidth 0
            lastTotal = Me.TotalLines
            RaiseEvent UpdateUI
        Case SCN_MODIFIED
            'int position, int modificationType, string text, int length, int linesAdded, int line, int foldLevelNow, int foldLevelPrev
            RaiseEvent Modified("word") 'TODO
        Case SCN_MACRORECORD
            'TODO
        Case SCN_MARGINCLICK
            'TODO
        Case SCN_NEEDSHOWN
            'TODO
        Case SCN_PAINTED
            'TODO
        Case SCN_USERLISTSELECTION
            'TODO
        Case SCN_DWELLSTART
            'TODO
        Case SCN_DWELLEND
            'TODO
    End Select
End Sub

' Devuelve una cadena que se extiende desde la posición actual del cursor
' hasta el primer espacio encontrado (hacia atrás)
Private Function GetWord() As String
Dim linebuf As String, str As String * 1000, c As String
Dim current As Long, startword As Long
    current = SendMessageString(sci, SCI_GETCURLINE, Len(str), str)
    startword = current
    linebuf = left$(str, startword)
    Do While (startword > 0)
        c = Mid$(linebuf, startword, 1)
        If Not isAlpha(c) Or startword = 1 Then Exit Do
        startword = startword - 1
    Loop
    GetWord = Trim(Mid$(linebuf, IIf(startword = 0, 1, startword), current))
End Function

' Devuelve verdadero si el argumento es un carácter alfanumérico
Private Function isAlpha(ch As String) As Boolean
    isAlpha = ((ch Like "[a-z]") Or (ch Like "[A-Z]") Or (ch Like "[0-9]")) And (ch <> " ")
End Function

' Pegar, ¿se puede en este momento? -*
Public Property Get CanPaste() As Boolean
    CanPaste = SendMessage(sci, SCI_CANPASTE, CLng(0), CLng(0))
End Property

' Pegar -*
Public Sub paste()
    Call Message(SCI_PASTE)
End Sub

' Copiar -*
Public Sub copy()
    Call Message(SCI_COPY)
End Sub

' Cortar -*
Public Sub cut()
    Call Message(SCI_CUT)
End Sub

' Deshacer -*
Public Sub undo()
    Call Message(SCI_UNDO)
End Sub

' Rehacer (?) -*
Public Sub redo()
    Call Message(SCI_REDO)
End Sub

' Devuelve la línea en la que se encuentra el cursor -*
Public Property Get CurLine() As Long
Attribute CurLine.VB_MemberFlags = "400"
    CurLine = SendMessage(sci, SCI_LINEFROMPOSITION, Me.CurPosition, CLng(0)) + 1
End Property

' Fija la línea actual. -*
Public Property Let CurLine(vNewValue As Long)
    Call Message(SCI_GOTOLINE, vNewValue)
End Property

' Devuelve la posición actual del cursor -*
Public Property Get CurPosition() As Long
Attribute CurPosition.VB_MemberFlags = "400"
    CurPosition = SendMessage(sci, SCI_GETCURRENTPOS, CLng(0), CLng(0))
End Property

' Fija la posición actual del cursor (mueve la zona visible si es necesario) -*
Public Property Let CurPosition(vNewValue As Long)
    Call Message(SCI_GOTOPOS, vNewValue)
End Property

' ¿Hay alguna acción para deshacer? -*
Public Property Get CanUndo() As Boolean
    CanUndo = SendMessage(sci, SCI_CANUNDO, CLng(0), CLng(0))
End Property

' Convertir los finales de línea al final especificado por 'newEOL' -*
Public Sub ConvertEOLs(newEOL As EOL)
    Call Message(SCI_CONVERTEOLS, newEOL)
End Sub

' Mostrar la lista automática
'   - prevchars: indica cuántos caractéres antes del cursor proporcionan el
'               contexto para buscar en la lista.
'   - autolist: lista separada por 'AutoCSeparatorChar' (espacios, por defecto) -*
Public Sub ShowAutoCList(prevchars As Long, autolist As String)
    Call Message(SCI_AUTOCSHOW, prevchars, autolist)
End Sub

' Oculta la lista -*
Public Sub HideAutoCList()
    Call Message(SCI_AUTOCCANCEL)
End Sub

' Carácter de separación de los elementos de la lista -*
Public Property Get AutoCSeparatorChar() As String
Attribute AutoCSeparatorChar.VB_MemberFlags = "400"
    AutoCSeparatorChar = m_SepChar
End Property

Public Property Let AutoCSeparatorChar(vNewValue As String)
    m_SepChar = vNewValue
    'Por seguridad, nos quedamos con el primer carácter
    Call Message(SCI_AUTOCSETSEPARATOR, Asc(left(m_SepChar, 1)))
    PropertyChanged "Separator"
End Property

' Si se oculta la lista al no encontrase la cadena tecleada -*
Public Property Get AutoCHideIfNoMatch() As Boolean
Attribute AutoCHideIfNoMatch.VB_MemberFlags = "400"
    AutoCHideIfNoMatch = m_AutoChide
End Property

Public Property Let AutoCHideIfNoMatch(vNewValue As Boolean)
    m_AutoChide = vNewValue
    Call Message(SCI_AUTOCSETAUTOHIDE, m_AutoChide)
    PropertyChanged "AutoChide"
End Property

' Cambio de lenguaje
Private Sub ChangeLang(language As String)
Dim nd As IXMLDOMElement
Dim n As Integer, m As Integer
Dim code As Long, fc As Long, bc As Long, ca As Long, ch As Long
Dim ef As Long, ev As Long, stylebits As Long
Dim words As String
Dim fnt As StdFont
Dim xmldoc As DOMDocument
    Set xmldoc = New DOMDocument
    If xmldoc.Load(Me.ConfFile) Then
        Set fnt = New StdFont
        xmldoc.validateOnParse = True
        Call ChangeDefaultStyle(xmldoc)
        Call Message(SCI_STYLECLEARALL)               ' Limpiamos los estilos
        ' Obtenemos el código del lenguaje para Scintilla
        Set nd = SearchNode(xmldoc, LNGCONST & language & SQLEND)
        code = nd.Attributes.getNamedItem("code").nodeValue
        ' Número de bits de estilo
        On Error Resume Next
        stylebits = nd.Attributes.getNamedItem("stylebits").nodeValue
        If Err Then stylebits = 5: Err.Clear
        On Error GoTo 0
        Call Message(SCI_SETLEXER, code)
        Call Message(SCI_SETSTYLEBITS, stylebits)
        ' Pasamos las palabras reservadas
        Set nd = SearchNode(xmldoc, LNGCONST & language & """]/keywordlists")
        For n = 0 To nd.childNodes.length - 1
            code = nd.childNodes(n).Attributes.getNamedItem("code").nodeValue
            words = nd.childNodes(n).Text
            Call Message(SCI_SETKEYWORDS, code, words)
        Next n
        ' Configuramos los estilos. Seguimos el orden tal y como se hace en SciTE:
        ' 1º Estilo global por defecto. 2º Estilo por defecto del lenguaje.
        ' 3º Resto de estilos globales. 4º Resto de estilos del lenguaje.
        ' De momento, en el fichero de configuración tiene que figurar primero el
        ' estilo por defecto del lenguaje.
        Set nd = SearchNode(xmldoc, LNGCONST & language & """]/styles")
        For n = 0 To nd.childNodes.length - 1
            code = nd.childNodes(n).Attributes.getNamedItem("code").nodeValue
            For m = 0 To nd.childNodes(n).childNodes.length - 1
                With nd.childNodes(n).childNodes(m)
                    Select Case .nodeName
                        Case "font"
                            fnt.Name = .Attributes.getNamedItem("name").nodeValue
                            fnt.Size = val(.Attributes.getNamedItem("size").nodeValue)
                            fnt.Bold = .Attributes.getNamedItem("bold").nodeValue
                            fnt.Italic = .Attributes.getNamedItem("italic").nodeValue
                            fnt.Underline = .Attributes.getNamedItem("underline").nodeValue
                        Case "forecolor"
                            fc = .Attributes.getNamedItem("value").nodeValue
                        Case "backcolor"
                            bc = .Attributes.getNamedItem("value").nodeValue
                        Case "eolfilled"
                            ef = IIf(.Attributes.getNamedItem("value").nodeValue = "true", 1, 0)
                        Case "visible"
                            ev = IIf(.Attributes.getNamedItem("value").nodeValue = "true", 1, 0)
                        Case "case"
                            ca = CLng(.Attributes.getNamedItem("value").nodeValue)
                        Case "charset"
                            ch = CLng(.Attributes.getNamedItem("value").nodeValue)
                    End Select
                End With
            Next m
            Style code, fc, bc, fnt, ef, ev, ca, ch
            If code = 0 Then Call ChangeDefStyles(xmldoc) ' Cambiamos los estilos globales tras cambiar el estilo por defecto
        Next n
        Set nd = Nothing
        Set xmldoc = Nothing
    Else
        Err.Raise 518, "ChangeLang", LoadResString(518) & ": " & xmldoc.parseError.reason _
                            & " " & xmldoc.parseError.srcText
    End If
End Sub

' Buscar un nodo según una cadena "XPath"
Private Function SearchNode(xmldoc As DOMDocument, ByVal sql As String) As IXMLDOMNode
Dim lst As IXMLDOMNodeList
    Set lst = xmldoc.documentElement.selectNodes(sql)
    If lst.length = 1 Then
        Set SearchNode = lst.Item(0)
    ElseIf lst.length = 0 Then
        Err.Raise 514, "SearchNode", LoadResString(514)
    Else
        Err.Raise 513, "SearchNode", LoadResString(513)
    End If
End Function

'Carga un fichero en el control. Reemplaza el texto que hubiera.
Public Sub LoadFile(file As String)
Dim fs As FileSystemObject
Dim ts As TextStream
    Set fs = New FileSystemObject
    Set ts = fs.OpenTextFile(file)
    Me.Text = ts.ReadAll
    Set ts = Nothing
    Set fs = Nothing
End Sub

Public Sub SaveFile(file As String)
'TODO
End Sub

' Realza los brazos de paréntesis... -*
Public Property Get MatchBraces() As Boolean
Attribute MatchBraces.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    MatchBraces = m_MatchBraces
End Property

Public Property Let MatchBraces(vNewValue As Boolean)
    m_MatchBraces = vNewValue
    PropertyChanged "MatchBraces"
End Property

' Devuelve el carácter situado en 'position' -*
Public Function CharAtPos(position As Long) As Long
    CharAtPos = SendMessage(sci, SCI_GETCHARAT, position, CLng(0))
End Function

' Devuelve verdadero si el código 'ch' corresponde a ( ó ) ó [ ó ] ó < ó >
Private Function IsBrace(ch As Long) As Boolean
    IsBrace = (ch = 40 Or ch = 41 Or ch = 60 Or ch = 62 Or ch = 91 Or ch = 93 Or ch = 123 Or ch = 125)
End Function

' Establece el nombre del fichero de configuración -*
Public Property Get ConfFile() As String
ConfFile = m_ConfFile
If InStr(1, m_ConfFile, "\", vbTextCompare) = 0 Or InStr(1, m_ConfFile, "/", vbTextCompare) = 0 Then
    ConfFile = App.Path & "\" & m_ConfFile
End If
End Property

' El fichero de configuración -*
Public Property Let ConfFile(vNewValue As String)
Dim fs As FileSystemObject
    If InStr(1, vNewValue, "\", vbTextCompare) = 0 Or InStr(1, vNewValue, "/", vbTextCompare) = 0 Then
        vNewValue = App.Path & "\" & vNewValue
    End If
    Set fs = New FileSystemObject
    ' Comprobamos si existe el fichero
    If Not fs.FileExists(vNewValue) Then
        Err.Raise 519, "ConfFile", Replace1(LoadResString(519), "%", vNewValue)
    Else
        m_ConfFile = vNewValue
        PropertyChanged "ConfFile"
    End If
    Set fs = Nothing
End Property

' Visualización de la barra de desplazamiento horizontal -*
Public Property Get HScrollBarVisible() As Boolean
Attribute HScrollBarVisible.VB_Description = "Muestra /oculta la barra de desplazamiento horizontal. Show / hides horizontal scroll bar"
Attribute HScrollBarVisible.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    HScrollBarVisible = m_HscrollBar
End Property

Public Property Let HScrollBarVisible(vNewValue As Boolean)
    m_HscrollBar = vNewValue
    Call Message(SCI_SETHSCROLLBAR, vNewValue)
    PropertyChanged "HScroll"
End Property

' Visualización de las guías de indentación -*
Public Property Get IndGuidesVisible() As Boolean
    IndGuidesVisible = m_IndGuides
End Property

Public Property Let IndGuidesVisible(vNewValue As Boolean)
    m_IndGuides = vNewValue
    Call Message(SCI_SETINDENTATIONGUIDES, m_IndGuides)
    PropertyChanged "IndGuides"
End Property

' Devuelve la columna de la posición actual -*
Public Property Get Column() As Long
Attribute Column.VB_MemberFlags = "400"
    Column = SendMessage(sci, SCI_GETCOLUMN, Me.CurPosition, CLng(0))
End Property
