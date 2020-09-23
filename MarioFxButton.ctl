VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl UserControl1 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   EditAtDesignTime=   -1  'True
   MaskColor       =   &H000000FF&
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   213
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1320
      Top             =   0
   End
   Begin PicClip.PictureClip ButtonPicture 
      Left            =   0
      Top             =   0
      _ExtentX        =   1984
      _ExtentY        =   10557
      _Version        =   393216
      Rows            =   19
      Picture         =   "MarioFxButton.ctx":0000
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Api's para dibujar el caption de el boton
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
'Focus Line
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


'Constantes de las propiedades de usercontrol Default
Private Const m_def_Caption = "MarioButton"
Private Const m_NumeroClips = 19
Private Const m_Efect = 0
Private Const m_VelAccelerator = 2
Private Const m_def_TransparentColor = vbWhite
Private Const m_def_HighlightColor = vbButtonText
Private Const m_def_ForeColor = vbButtonText


Private m_Font As Font
Private Num As Single
Private NumeroClips As Single
Private Accelerador As Single
Private tPrevEvent As String
Private m_Caption As String
Private m_TransparentColor As OLE_COLOR
Private m_HighlightColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_ButtonStyle As Picture
Private m_RightToLeft As Boolean
Private focus As Boolean
Private efect As Boolean
Private IsIn As Boolean
Private IsOut As Boolean
Private IsDown As Boolean





Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum EfectoBoton
    Always = 0
    MouseMove = 1
End Enum


Private Enum EstadoBoton
    Down = 0
    Over = 1
End Enum

Private Estado As EstadoBoton
Private Efecto As EfectoBoton
'//---------------------------------------------------------------------------------------
'// Eventos de el Control
'//---------------------------------------------------------------------------------------

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseExit()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)



Private Sub MouseOut()
    efect = True
    IsOut = True
    IsIn = False
End Sub



Private Sub UserControl_Click()
    Call RaiseEventEx("Click")
End Sub

Private Sub UserControl_DblClick()
 Call RaiseEventEx("DblClick")
End Sub

Private Sub UserControl_GotFocus()
focus = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RaiseEventEx("KeyDown", KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Call RaiseEventEx("KeyPress", KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Call RaiseEventEx("KeyUp", KeyCode, Shift)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call RaiseEventEx("Click")
End Sub
Private Function RaiseEventEx(ByVal Name As String, ParamArray Params() As Variant)
        
    Select Case Name
        Case "Click"
            'click event occurred
            RaiseEvent Click
        
        Case "KeyDown"
            'key down event occurred
            RaiseEvent KeyDown(CInt(Params(0)), CInt(Params(1)))
        
        Case "KeyPress"
            'key press event occurred
            RaiseEvent KeyPress(CInt(Params(0)))
        
        Case "KeyUp"
            'key up event occurred
            RaiseEvent KeyUp(CInt(Params(0)), CInt(Params(1)))
        
        Case "MouseDown"
            'mouse down event occurred
            RaiseEvent MouseDown(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        
        Case "MouseMove"
            'mouse move event occurred
            RaiseEvent MouseMove(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        
        Case "MouseUp"
            'mouse up event occurred
            RaiseEvent MouseUp(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        
        Case "MouseExit"
            'mouse exit event occurred
            If tPrevEvent <> "MouseExit" Then
                RaiseEvent MouseExit
            End If
    
            'save previous event (for MouseEnter and MouseExit events)
            tPrevEvent = Name
        
        Case "MouseEnter"
            'mouse enter event occurred
            If tPrevEvent <> "MouseEnter" Then
                RaiseEvent MouseEnter
            End If
    
            'save previous event (for MouseEnter and MouseExit events)
            tPrevEvent = Name
       
    End Select
End Function

Private Sub UserControl_LostFocus()
focus = False
MouseOut
End Sub

Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Error
Estado = Down
IsDown = True
IsIn = False
IsOut = False
If NumeroClips = 0 Then Exit Sub
Num = NumeroClips - 1
Picture = ButtonPicture.GraphicCell(Num)
DrawCaption (Estado)
Refresh
Exit Sub
Error: MsgBox "Error: Numero de Clips Incorrecto!", vbInformation, "MarioFXCommands"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsIn = True Then Exit Sub
If IsDown = True Then Exit Sub
IsOut = False
IsIn = True

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Estado = Over
IsDown = False
UserControl_MouseMove Button, Shift, X, Y
End Sub


Private Sub Timer1_Timer()

    Dim pnt As POINTAPI
    Dim a
    On Error Resume Next
    
    DrawCaption (Estado)
    If IsOut = True Then GoTo Skip
    If IsIn = False Then Exit Sub
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt
    
   
    'Funcion para saber si el moude se salio de el control
    If pnt.X < UserControl.ScaleLeft Or _
            pnt.Y < UserControl.ScaleTop Or _
            pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
            pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
       MouseOut
       Exit Sub
    End If
    
   
    
    If efect = False Then     'Accendente
    
            For a = 1 To Accelerador
               
                 UserControl.Picture = ButtonPicture.GraphicCell(Num)
                 DrawCaption (Estado)
                 UserControl.Refresh 'Para Evitar Flickerness
                
                 Num = Num + 1
            Next a
    
            If Num > NumeroClips Then
                Num = NumeroClips - 1
                If Efecto = Always Then efect = True
                Exit Sub
            End If
    End If
Skip:
   
   If efect = True Then       'Descendente

            For a = 1 To Accelerador
          
            UserControl.Picture = ButtonPicture.GraphicCell(Num)
            DrawCaption (Estado)
            UserControl.Refresh

            If Num > 0 Then Num = Num - 1
            Next a
    
            If Num < 1 Then
                Num = 1
                efect = False
                Exit Sub
            End If
   End If
   



End Sub

Private Sub UserControl_Initialize()
    
    Num = 0
    IsDown = False
    IsOut = False
    IsIn = False
    efect = False
    ForeColor = m_def_ForeColor
    Accelerador = 2
    Estado = Over
    DrawCaption (Estado) 'Dibujar Caption Al iniciar

    
    
End Sub

Private Sub UserControl_InitProperties()
 HighlightColor = m_def_HighlightColor
 TransparentColor = m_def_TransparentColor
 Effect = Always
 NumClips = m_NumeroClips
 Set Font = Ambient.Font
 Caption = m_def_Caption
 Set ButtonStyle = LoadPicture("")
End Sub




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 NumClips = PropBag.ReadProperty("NumeroClips", m_NumeroClips)
 Aceleration = PropBag.ReadProperty("Aceleration", m_VelAccelerator)
 TransparentColor = PropBag.ReadProperty("TransparentColor", m_def_TransparentColor)
 HighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)
 ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
 Caption = PropBag.ReadProperty("Caption", m_def_Caption)
 Effect = PropBag.ReadProperty("Efect", m_Efect)
 Set ButtonStyle = PropBag.ReadProperty("ButtonStyle", Nothing)
 Set Font = PropBag.ReadProperty("Font", Ambient.Font)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("NumeroClips", NumeroClips, m_NumeroClips)
 Call PropBag.WriteProperty("Efect", Efecto, m_Efect)
 Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, m_def_TransparentColor)
 Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, m_def_HighlightColor)
 Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
 Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, Nothing)
 Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
 Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
 Call PropBag.WriteProperty("Aceleration", Accelerador, m_VelAccelerator)
End Sub

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    UserControl.ForeColor = NewValue
    PropertyChanged "ForeColor"
End Property
Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal NewValue As Font)
    Set m_Font = NewValue
    Set UserControl.Font = NewValue
    PropertyChanged "Font"
End Property
Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal NewValue As OLE_COLOR)
    m_HighlightColor = NewValue
    PropertyChanged "HighlightColor"
End Property
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = m_TransparentColor
    UserControl.MaskColor = TransparentColor
End Property

Public Property Let TransparentColor(ByVal NewValue As OLE_COLOR)
    m_TransparentColor = NewValue
    PropertyChanged "TransparentColor"
    UserControl.MaskColor = TransparentColor
End Property

Public Property Get ButtonStyle() As Picture
    Set ButtonStyle = m_ButtonStyle
    UserControl.MaskPicture = ButtonPicture.GraphicCell(0)
    UserControl.Picture = ButtonPicture.GraphicCell(0)
    Resize
End Property

Public Property Set ButtonStyle(ByVal NewValue As Picture)
    Set m_ButtonStyle = NewValue
    PropertyChanged "ButtonStyle"
    If ButtonStyle Is Nothing Then Exit Property
    If ButtonStyle = 0 Then Exit Property     'Asegurarse de que el control nunca se quede sin foto
        ButtonPicture.Picture = ButtonStyle
        UserControl.MaskPicture = ButtonPicture.GraphicCell(0)
        UserControl.Picture = ButtonPicture.GraphicCell(0)
        Resize
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
    Dim lPlace As Long
    m_Caption = NewValue
    lPlace = 0
    lPlace = InStr(lPlace + 1, NewValue, "&", vbTextCompare)
    Do While lPlace <> 0
        If Mid$(NewValue, lPlace + 1, 1) <> "&" Then
            AccessKeys = Mid$(NewValue, lPlace + 1, 1)
            Exit Do
        Else
            lPlace = lPlace + 1
        End If
    
        lPlace = InStr(lPlace + 1, NewValue, "&", vbTextCompare)
    Loop
    PropertyChanged "Caption"
End Property
Public Property Get NumClips() As Single
       NumClips = NumeroClips
       ButtonPicture.Rows = NumClips
 End Property

Public Property Let NumClips(ByVal NewValue As Single)
    NumeroClips = NewValue
    PropertyChanged "NumeroClips"
    ButtonPicture.Rows = NumClips
End Property
Public Property Get Effect() As EfectoBoton
       Effect = Efecto
End Property

Public Property Let Effect(ByVal NewValue As EfectoBoton)
    Efecto = NewValue
    PropertyChanged "Efect"
End Property

Public Property Get Aceleration() As Single
       Aceleration = Accelerador
End Property

Public Property Let Aceleration(ByVal NewValue As Single)
    Accelerador = NewValue
    PropertyChanged "Aceleration"
End Property
Private Sub Resize()
 'resize control
        Width = Picture.Width / 1.76
        Height = Picture.Height / 1.76
End Sub
Private Sub DrawCaption(ByVal Estado As EstadoBoton)
    
    Dim lLeft As Long
    Dim lTop As Long
    
       
   Select Case Estado
         
         Case Over
           UserControl.ForeColor = m_ForeColor
         Case Down
           UserControl.ForeColor = m_HighlightColor
   End Select

    
    lLeft = ((ScaleWidth \ 2) - (TextWidth(m_Caption) \ 2))
    lTop = ((ScaleHeight \ 2) - (TextHeight(m_Caption) \ 2))
    
    
    Call DrawStateText(hdc, 0, 0, m_Caption, Len(m_Caption), lLeft, lTop + 1, 0, 0, &H2)
    'Marcar Linea de Focus en el boton
    If focus = True Then FocusLine lLeft, lTop + TextHeight(m_Caption) + 1, TextWidth(m_Caption), 1, m_ForeColor
    
    
    

End Sub


Private Sub FocusLine(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)

Dim XRECT As RECT
Dim XBrush As Long

XRECT.Left = X
XRECT.Top = Y
XRECT.Right = X + Width
XRECT.Bottom = Y + Height

XBrush = CreateSolidBrush(Color)
'Llamada Api
FrameRect UserControl.hdc, XRECT, XBrush
DeleteObject XBrush
End Sub

