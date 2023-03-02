VERSION 5.00
Begin VB.Form fDelay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alert"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   ControlBox      =   0   'False
   Icon            =   "frmDelay.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   180
      Top             =   3090
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   2993
      TabIndex        =   1
      Top             =   3090
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Si ves este mensaje, es que se han olvidado de indicar el verdadero mensaje a mostrar...   En fin..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6885
   End
End
Attribute VB_Name = "fDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Delayed Form                                                      (25/Ago/98)
'
' Se mostrará este formulario hasta que se pulse en el botón o
' transcurra el tiempo especificado.
'
' El mensaje a mostrar y el tiempo de espera, se indicarán por medio de las
' propiedades: Message y Delay respectivamente.
'
' Para usarla:
'        With fDelay
'            .Delay = 4000 ' Tiempo de espera en milisegundos
'            .Message = "Mensaje a mostrar"
'            .Show vbModal, Me
'        End With
'
' Revisión del 30/Nov/2000: Nuevas propiedades para el tipo de fuente
'
' ©Guillermo 'guille' Som, 1998-2000
'------------------------------------------------------------------------------
Option Explicit

Public Property Get Delay() As Long
    Delay = Timer1.Interval
End Property

Public Property Let Delay(ByVal NewDelay As Long)
    On Error Resume Next
    '
    Timer1.Enabled = False
    Timer1.Interval = NewDelay
    If Err Then
        Err = 0
        Timer1.Interval = 10000
    End If
    Timer1.Enabled = True
End Property

Public Property Get Message() As String
    Message = Label1.Caption
End Property

Public Property Let Message(ByVal sNewMsg As String)
    With Label1
        .Caption = sNewMsg
        .Refresh
    End With
End Property
Private Sub Command1_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fDelay = Nothing
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Unload Me
End Sub

Public Property Get msgFontSize() As Long
    msgFontSize = Label1.FontSize
End Property

Public Property Let msgFontSize(ByVal NewValue As Long)
    On Error Resume Next
    '
    Label1.FontSize = NewValue
    '
    Err = 0
End Property

Public Property Get msgFontBold() As Boolean
    msgFontBold = Label1.FontBold
End Property

Public Property Let msgFontBold(ByVal NewValue As Boolean)
    Label1.FontBold = NewValue
End Property

Public Property Get msgFontName() As String
    msgFontName = Label1.FontName
End Property

Public Property Let msgFontName(ByVal NewValue As String)
    On Error Resume Next
    '
    Label1.FontName = NewValue
    '
    Err = 0
End Property
