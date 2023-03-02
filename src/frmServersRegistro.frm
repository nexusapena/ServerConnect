VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmServersRegistro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tpdServerConnect v22.09"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "frmServersRegistro.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdBorrarValor 
      Caption         =   "Borrar el Nombre"
      Height          =   375
      Left            =   18600
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   " Borra el nombre y el valor "
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Buscar / Reemplazar valores:"
      Height          =   2775
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   4905
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "frmServersRegistro.frx":1122A
         Top             =   2295
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.CheckBox chkConfirmar 
         Caption         =   "Solicitar confirmación"
         Height          =   255
         Left            =   2340
         TabIndex        =   39
         ToolTipText     =   " Si se pide confirmación al reemplazar "
         Top             =   1230
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkTipoComparacion 
         Caption         =   "Mayúsculas / Minúsculas"
         Height          =   255
         Left            =   2340
         TabIndex        =   38
         ToolTipText     =   " Comprueba las palabras teniendo en cuenta mayúsculas y minúsculas "
         Top             =   1500
         Width           =   2235
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar en:"
         Height          =   1125
         Index           =   4
         Left            =   150
         TabIndex        =   34
         Top             =   1140
         Width           =   1995
         Begin VB.CheckBox chkBuscarEn 
            Caption         =   "Datos"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   37
            Top             =   780
            Width           =   1665
         End
         Begin VB.CheckBox chkBuscarEn 
            Caption         =   "Valores (nombres)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   510
            Width           =   1665
         End
         Begin VB.CheckBox chkBuscarEn 
            Caption         =   "Claves"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Value           =   1  'Checked
            Width           =   1665
         End
      End
      Begin VB.CheckBox chkCambiar 
         Caption         =   "Cambiar por:"
         Height          =   315
         Left            =   150
         TabIndex        =   33
         ToolTipText     =   " Marca esta casilla para reemplazar "
         Top             =   750
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtPoner 
         Height          =   315
         Left            =   1410
         TabIndex        =   32
         Text            =   "txtPoner"
         Top             =   750
         Visible         =   0   'False
         Width           =   3345
      End
   End
   Begin VB.CommandButton cmdLeerClave 
      Caption         =   "Mostrar las subclaves"
      Default         =   -1  'True
      Height          =   375
      Left            =   12960
      TabIndex        =   30
      ToolTipText     =   " Mostrar los valores y las subclaves contenidas en la Clave indicada "
      Top             =   2280
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton cmdBorrarClave 
      Caption         =   "Borrar la clave"
      Height          =   375
      Left            =   11400
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   " Borrar la clave indicada y todas las subclaves ¡PRECAUCIÓN! "
      Top             =   2520
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdLeerTodo 
      Caption         =   "Mostrar todas las subclaves"
      Height          =   375
      Left            =   11520
      TabIndex        =   28
      ToolTipText     =   " Mostrar todas las subclaves por debajo de la calve indicada, (Esto puerde tardar un buen rato) "
      Top             =   1800
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      Left            =   11880
      TabIndex        =   27
      Text            =   "HKEY_USERS\.Default\Software\elGuille\Pruebas Registro"
      Top             =   3600
      Visible         =   0   'False
      Width           =   9765
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   14040
      TabIndex        =   26
      Text            =   "Valor"
      Top             =   3960
      Visible         =   0   'False
      Width           =   5565
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   12120
      TabIndex        =   25
      Text            =   "Valor"
      Top             =   6240
      Visible         =   0   'False
      Width           =   5565
   End
   Begin VB.ComboBox cboNombre 
      Height          =   315
      Left            =   13320
      TabIndex        =   24
      Text            =   "cboNombre"
      Top             =   5400
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.TextBox txtPath 
      Height          =   405
      Left            =   14040
      TabIndex        =   23
      Text            =   "D:\IRISHealth\Bin"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de valor a asignar/borrar:"
      Height          =   1575
      Index           =   1
      Left            =   15960
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton optTipo 
         Caption         =   "String (cadena ampliada) "
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   22
         ToolTipText     =   " REG_MULTI_SZ  (no implementada) "
         Top             =   1170
         Width           =   2355
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "DWORD (numérico) "
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   21
         ToolTipText     =   " REG_DWORD "
         Top             =   870
         Width           =   2355
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Binary (binario) "
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   20
         ToolTipText     =   " REG_BINARY "
         Top             =   570
         Width           =   2355
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar el valor"
         Height          =   375
         Left            =   2610
         TabIndex        =   19
         ToolTipText     =   " Asigna el valor indicado al nombre de la clave "
         Top             =   1050
         Width           =   1695
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "String (cadena) "
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   18
         ToolTipText     =   " REG_SZ, REG_EXPAND_SZ "
         Top             =   270
         Value           =   -1  'True
         Width           =   2355
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Des-registrar servidor ActiveX:"
      Height          =   1515
      Index           =   3
      Left            =   11280
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   9555
      Begin VB.TextBox txtUnregister 
         Height          =   285
         Index           =   2
         Left            =   810
         TabIndex        =   12
         Text            =   "Servidor.Clase"
         Top             =   1080
         Width           =   6795
      End
      Begin VB.CommandButton cmdUnRegister 
         Caption         =   "Eliminar del Registro"
         Height          =   405
         Index           =   1
         Left            =   7710
         TabIndex        =   11
         ToolTipText     =   " Eliminar la clase mostrada del registro ¡CUIDADO! "
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdUnRegister 
         Caption         =   "Mostrar info"
         Height          =   645
         Index           =   0
         Left            =   7710
         TabIndex        =   10
         ToolTipText     =   " Mostrar información sobre la clase "
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtUnregister 
         Height          =   285
         Index           =   1
         Left            =   810
         TabIndex        =   9
         Text            =   "Servidor.Clase"
         Top             =   690
         Width           =   6795
      End
      Begin VB.TextBox txtUnregister 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   8
         Text            =   "Servidor.Clase"
         Top             =   330
         Width           =   5025
      End
      Begin VB.Label Label1 
         Caption         =   "TypeLib:"
         Height          =   255
         Index           =   14
         Left            =   180
         TabIndex        =   16
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "CLSID:"
         Height          =   255
         Index           =   13
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Clase, en formato Servidor.Clase:"
         Height          =   255
         Index           =   12
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "&Nombre / Valor:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7800
      Picture         =   "frmServersRegistro.frx":11292
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Terminar el programa "
      Top             =   6120
      Width           =   600
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4935
      Left            =   12240
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Label Label1 
         Caption         =   "Claves:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "999.999.999"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   44
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Valores:"
         Height          =   255
         Index           =   6
         Left            =   1650
         TabIndex        =   43
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "999.999.999"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   42
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "&Clave:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TPD Servers Connect IRIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8580
      Begin VB.CommandButton cmdRDP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4250
         Picture         =   "frmServersRegistro.frx":1171C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Remote Desktop"
         Top             =   6120
         Width           =   600
      End
      Begin VB.CommandButton cmbSTUDIO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4250
         Picture         =   "frmServersRegistro.frx":3999A
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "STUDIO"
         Top             =   5520
         Width           =   600
      End
      Begin VB.CommandButton cmbTERMINAL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4250
         Picture         =   "frmServersRegistro.frx":4014F
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "TERMINAL"
         Top             =   4920
         Width           =   600
      End
      Begin VB.CommandButton cmbPORTAL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4250
         Picture         =   "frmServersRegistro.frx":40591
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Portal IRIS"
         Top             =   4320
         Width           =   600
      End
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3600
         Picture         =   "frmServersRegistro.frx":4165B
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Buscar en Servidores"
         Top             =   6720
         Width           =   400
      End
      Begin VB.TextBox txtBuscar 
         Height          =   315
         Left            =   220
         TabIndex        =   46
         Text            =   "txtBuscar"
         ToolTipText     =   "Texto a buscar"
         Top             =   6800
         Width           =   3105
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6075
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Docle click abre Portal, Terminal y Studio"
         Top             =   500
         Width           =   4100
         _ExtentX        =   7223
         _ExtentY        =   10716
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   847
         LabelEdit       =   1
         Style           =   1
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3800
         Left            =   4250
         TabIndex        =   49
         Top             =   500
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   6694
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Servidores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   " Lista de Claves y subclaves "
         Top             =   260
         Width           =   4815
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Toni Peña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   47
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Path IRIS"
      Height          =   255
      Left            =   12120
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmServersRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Prueba de manipulación del Registro                               (13/Ago/98)
'
' Con pruebas de creación/modificación/borrado                (01:50 18/Ago/98)
'
' Nueva versión para leer y modificar entradas del registro         (31/Ene/99)
' Nueva versión, usando Treeview y ListView                         (23/Sep/00)
' Mostrando un formulario de AVISO                                  (30/Nov/00)
'
' Nueva versión                                                     (28/Dic/01)
'   permite cambiar valores (Buscar/Reemplazar)
'   Algunas confirmaciones extras antes de borrar...
'
' ©Guillermo 'guille' Som, 1998-2001
'------------------------------------------------------------------------------
Option Explicit
'Option Compare Text
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_NORMAL = 1
    Const SWPATH = 1
    Const SWADDRESS = 2
    Const SWAUTHENTICATIONMETHOD = 3
    Const SWCOMMENT = 4
    Const SWCONNECTIONSECURITYLEVEL = 5
    Const SWHTTPS = 6
    Const SWPORT = 7
    Const SWSERVERTYPE = 8
    Const SWSERVICEPRINCIPALNAME = 9
    Const SWTELNET = 10
    Const SWWEBSERVERADDRESS = 11
    Const SWWEBSERVERINSTANCENAME = 12
    Const SWWEBSERVERPORT = 13





' Pongo estas variables a nivel de módulo para mayor rapidez
' (o por necesidad...)
Private TipoComparacion As VbCompareMethod
Private nClavesHalladas As Long
Private nValoresHallados As Long
Private nClavesExaminadas As Currency
Private nValoresExaminados As Currency
Private sBuscar As String, sPoner As String
Private Reemplazar As Boolean, PedirConfirmacion As Boolean
Private EnClaves As Boolean, EnNombres As Boolean, EnDatos As Boolean
'
Private Cancelado As Boolean
Private tQR As cQueryReg
Private ma_scolKeys() As String

    
Private Sub cboNombre_Click()
    ' Leer el valor de este nombre
    Dim sKey As String
    Dim sName As String
    '
    sKey = cboNombre.Tag ' Trim$(txtClave)
    sName = cboNombre.Text
    txtValor = tQR.GetReg(sKey, sName, , True)
End Sub

Private Sub chkCambiar_Click()
    With txtPoner
        If chkCambiar.Value = vbChecked Then
            .Enabled = True
            .BackColor = vbWindowBackground
            chkConfirmar.Enabled = True
            cmdBuscar.Caption = "Reemplazar"
        Else
            .Enabled = False
            .BackColor = vbButtonFace
            chkConfirmar.Enabled = False
            'cmdBuscar.Caption = "Buscar"
        End If
    End With
End Sub




Private Sub cmbSTUDIO_Click()
    Dim lValDev As Long
    Dim studio As String
    Dim params As String
    Me.ListView1.Refresh
    studio = Me.txtPath.Text + "\bin\CStudio.exe"
    params = "/servername=" + TreeView1.SelectedItem.Text
    
    lValDev = ShellExecute(Me.hwnd, "Open", studio, params, "", 1)

End Sub

Private Sub cmdAsignar_Click()
    'Prueba de asignar datos
    Dim sKey As String
    Dim sName As String
    Dim sValue As String
    Dim bValue() As Byte
    Dim asValue() As String
    Dim lValue As Long
    Dim lRet As eHKEYError
    Dim rDT As eHKEYDataType
    Dim sMsg As String
    Dim lMsg As Long
    
    lMsg = vbInformation
    sKey = txtClave
    
    sName = cboNombre.Text
    ' Asignar sólo si el nombre de la clave tiene algo escrito          (31/Ene/99)
    If Len(sName) Then
        If optTipo(0) Then      ' Valor String
            sValue = txtValor
            rDT = REG_SZ
            lRet = tQR.SetReg(sKey, sName, sValue, , rDT, True)
            If lRet = ERROR_SUCCESS Then
                sMsg = "Asignado correctamente el dato String"
            Else
                lMsg = vbExclamation
                sMsg = "Error al asignar el dato String"
            End If
            '
        ElseIf optTipo(1) Then      ' Valor Binary
            bValue = txtValor
            rDT = REG_BINARY
            lRet = tQR.SetReg(sKey, sName, bValue, , rDT, True)
            If lRet = ERROR_SUCCESS Then
                sMsg = "Asignado correctamente el dato Binary"
            Else
                lMsg = vbExclamation
                sMsg = "Error al asignar el dato Binary"
            End If
            '
        ElseIf optTipo(2) Then      ' Valor DWORD
            ' Si no se ha escrito nada es un valor CERO
            If Len(txtValor) = 0 Then
                txtValor = "0"
            End If
            lValue = CLng(txtValor)
            rDT = REG_DWORD
            lRet = tQR.SetReg(sKey, sName, lValue, , rDT, True)
            If lRet = ERROR_SUCCESS Then
                sMsg = "Asignado correctamente el dato DWORD"
            Else
                lMsg = vbExclamation
                sMsg = "Error al asignar el dato DWORD"
            End If
            '
        ElseIf optTipo(3) Then      ' Valor REG_MULTI_SZ
            ' Este tipo será un array de cadenas,                   (22/Nov/00)
            ' en el valor se indicará cada cadena separada por punto y coma
            '
        End If
    End If
    '
    MsgBox sMsg, lMsg, "Asignar valores a una clave del registro"
End Sub

Private Sub cmdBorrarClave_Click()
    'Borrar la clave completa
    Dim sKey As String
    Dim lRet As eHKEYError
    '
    sKey = txtClave
    ' Preguntar si se quiere borrar
    If MsgBox("¿Quieres borrar esta clave:" & vbCrLf & sKey & "?", vbCritical + vbYesNo + vbDefaultButton2, "Borrar clave del registro del sistema") = vbYes Then
        ' Pedir una segunda confirmación... vayamos a p***cas       (28/Dic/01)
        If MsgBox("¿Seguro que quieres borrar la clave" & vbCrLf & sKey & "?" & vbCrLf & vbCrLf & "¡Que pesado soy! ¿verdad?" & vbCrLf & "Pero es que no es plan de jugar con estas cosas..." & vbCrLf & vbCrLf & "Aún estás a tiempo de pulsar en NO..." & vbCrLf & "pero si aún quieres borrar esa clave, pulsa en SI..." & vbCrLf & "¡tu sabrás lo que haces!", vbCritical + vbYesNo + vbDefaultButton2, "Borrar clave del registro del sistema") = vbYes Then
            If InStr(1, sKey, "Software\Microsoft\Windows", vbTextCompare) Then
                ' Esta clave es vital...
                If MsgBox("PSSST!!!" & vbCrLf & vbCrLf & "¿Sabes que esa clave es importante?" & vbCrLf & vbCrLf & "Ya no te advierto más..." & vbCrLf & vbCrLf & "¿Quieres borrarla de todas formas?", vbCritical + vbYesNo + vbDefaultButton2, "Borrar clave del registro del sistema") = vbNo Then
                    MsgBox "Te has hecho de rogar... pero... ¡¡¡al fin has visto la luz!!!", vbInformation + vbOKOnly, "Menos mal que no la has borrado"
                    Exit Sub
                End If
            End If
            lRet = tQR.DeleteKey(sKey)
            If lRet = ERROR_SUCCESS Then
                MsgBox "Se ha borrado la clave: " & sKey & ", con éxito.", vbInformation, "Borrar clave del registro"
            Else
                MsgBox "No se ha podido borrar la clave: " & sKey, vbExclamation, "Borrar clave del registro"
            End If
        End If
    End If
End Sub

Private Sub cmdBorrarValor_Click()
    'Borrar el valor seleccionado en Option1
    'Borrar la clave completa
    Dim sKey As String
    Dim sName As String
    Dim lRet As eHKEYError
    Dim sMsg As String
    Dim lMsg As Long
    
    lMsg = vbInformation
    
    sKey = Trim$(txtClave)
    sName = cboNombre.Text
    If Len(sName) = 0 Then
        MsgBox "No puedes dejar el valor del Nombre en blanco..." & vbCrLf & "ya que se borraría la clave completa..." & vbCrLf & "Si quieres borrar la clave, usa el botón correspondiente... pero... ¡cuidado con lo que haces!", vbExclamation + vbOKOnly, "Borrar un valor"
        txtClave.SetFocus
        Exit Sub
    End If
    ' Pedir confirmación antes de borrar
    If MsgBox("¿Quieres borrar la entrada: " & sName & " de la clave:" & vbCrLf & sKey & "?", vbQuestion + vbDefaultButton2 + vbYesNo, "Borrar una entrada del registro") = vbYes Then
        lRet = tQR.DeleteKey(sKey, sName)
        If lRet = ERROR_SUCCESS Then
            sMsg = "Se ha borrado el nombre: '" & sName & "' de la clave con éxito."
        Else
            lMsg = vbExclamation
            sMsg = "No se ha podido borrar el nombre: '" & sName & "' de la clave."
        End If
        MsgBox sMsg, lMsg, "Borrar nombre de una clave del registro"
    End If
End Sub

Private Sub cmdBuscar_Click()
    ' Buscar o Reemplazar                                           (28/Dic/01)
    ' Avisar de que se debe tener cuidado con esto de reemplazar automáticamente
    Dim i As Long, j As Long
    Dim sKey As String
    Dim s As String
    Static YaEstoy As Boolean
    '
    ' Si estamos dentro, es que se ha pulsado en cancelar
    If YaEstoy Then
        Cancelado = True
        DoEvents
        Exit Sub
    End If
    '
    sBuscar = txtBuscar
    sPoner = txtPoner
    EnClaves = (Me.chkBuscarEn(0).Value = vbChecked)
    EnNombres = (Me.chkBuscarEn(1).Value = vbChecked)
    EnDatos = (Me.chkBuscarEn(2).Value = vbChecked)
    '
    If chkTipoComparacion.Value = vbChecked Then
        TipoComparacion = vbBinaryCompare
    Else
        TipoComparacion = vbTextCompare
    End If
    ' Si no se ha indicado nada para buscar
    If Len(sBuscar) = 0 Then
        txtClave = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\Cache\Servers"
        LeerClaves Trim$(txtClave)
        txtBuscar.Text = ""
        'MsgBox "Debes especificar lo que quieres buscar.", vbInformation + vbOKOnly, "Buscar / Reemplazar"
        'txtBuscar.SetFocus
        Exit Sub
    End If
    '
    If chkCambiar.Value = vbChecked Then
        If Len(sPoner) = 0 Then
            MsgBox "Debes indicar el texto que reemplazará lo buscado.", vbInformation + vbOKOnly, "Buscar / Reemplazar"
            txtPoner.SetFocus
            Exit Sub
        End If
        Reemplazar = True
        PedirConfirmacion = (chkConfirmar.Value = vbChecked)
        If PedirConfirmacion = False Then
            s = "Además no has marcado la opción de pedir confirmación."
        End If
        If MsgBox("¡ATENCIÓN!" & vbCrLf & "Es muy importante que sepas que cambiando los valores de forma automática," & vbCrLf & "se pueden producir efectos no deseados en el registro." & vbCrLf & vbCrLf & "¿Seguro que quieres reemplazar?" & vbCrLf & vbCrLf & s, vbExclamation + vbYesNo + vbDefaultButton2, "Buscar / Reemplazar") = vbNo Then
            Exit Sub
        End If
        If PedirConfirmacion = False Then
            If MsgBox("Perdona que me haga pesado, pero..." & vbCrLf & "¿SABES LO PELIGROSO QUE ES CAMBIAR DE FORMA AUTOMÁTICA EL REGISTRO?" & vbCrLf & vbCrLf & "Aún después de este aviso a la cordura..." & vbCrLf & "¿Quieres continuar con la opción de reemplazar?" & vbCrLf & "(te recuerdo que se hará sin pedir confirmación)", vbExclamation + vbYesNo + vbDefaultButton2, "Buscar / Reemplazar") = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Que conste que te he avisado..." & vbCrLf & "No me hago responsable de lo que pueda pasar con el registro..." & vbCrLf & vbCrLf & "Aún tienes una última oportunidad de cancelar.", vbCritical + vbOKCancel + vbDefaultButton2, "Buscar / Reemplazar") = vbCancel Then
            Exit Sub
        End If
    End If
    '
    ' Buscar a partir de la clave indicada en txtClave y en las subclaves
    '
    ' Guardar los valores actuales en un fichero de configuración
    'TODO:
    '
    YaEstoy = True
    Cancelado = False
    'cmdBuscar.Caption = "Cancelar"
    MousePointer = vbArrowHourglass
    DoEvents
    '
    ' Deshabilitar todos los demás botones
    cmdAsignar.Enabled = Not YaEstoy
    cmdBorrarValor.Enabled = Not YaEstoy
    cmdBorrarClave.Enabled = Not YaEstoy
    cmdLeerTodo.Enabled = Not YaEstoy
    cmdLeerClave.Enabled = Not YaEstoy
    cmdUnRegister(0).Enabled = Not YaEstoy
    cmdUnRegister(1).Enabled = Not YaEstoy
    ' No deshabilitar el Frame1(2), sino ¡no se podrá pulsar en Cancelar!
    ' Deshabilitar el 1, 3 y 4
    For i = 1 To 3 Step 2
        Frame1(i).Enabled = Not YaEstoy
    Next
    Frame1(i - 1).Enabled = Not YaEstoy
    '
    For i = 4 To 7
        Label1(i).Visible = True
    Next
    Label1(5) = ""
    Label1(7) = ""
    '
    Label1(1) = "&SubClaves:"
    DoEvents
    '
    '
    ' Si tiene el separador del final, quitárselo                   (23/Nov/00)
    sKey = Trim$(txtClave.Text)
    If Right$(sKey, 1) = "\" Then
        sKey = Left$(sKey, Len(sKey) - 1)
    End If
    With TreeView1
        .Nodes.Clear
    End With
    With ListView1
        .ListItems.Clear
        .FullRowSelect = True
        .Sorted = True
    End With
    '
    nValoresExaminados = 0
    nClavesExaminadas = 0
    nClavesHalladas = 0
    nValoresHallados = 0
    '
    ' Por si se busca en Nombres y Datos
    If EnNombres = True Or EnDatos = True Then
        BuscarValores sKey, REG_SZ
        BuscarValores sKey, REG_DWORD
    End If
    ' Enumerar las subclaves y buscar en ellas y los Nombres/Datos
    BuscarClaves sKey
    '
    YaEstoy = False
    Cancelado = False
    If Me.chkCambiar.Value = vbChecked Then
        'cmdBuscar.Caption = "Reemplazar"
    Else
        'cmdBuscar.Caption = "Buscar"
    End If
    '
    ' Volver a habilitar los botones
    cmdAsignar.Enabled = Not YaEstoy
    cmdBorrarValor.Enabled = Not YaEstoy
    cmdBorrarClave.Enabled = Not YaEstoy
    cmdLeerTodo.Enabled = Not YaEstoy
    cmdLeerClave.Enabled = Not YaEstoy
    cmdUnRegister(0).Enabled = Not YaEstoy
    cmdUnRegister(1).Enabled = Not YaEstoy
    ' No deshabilitar el Frame1(2), sino ¡no se podrá pulsar en Cancelar!
    For i = 1 To 3 Step 2
        Frame1(i).Enabled = Not YaEstoy
    Next
    Frame1(i - 1).Enabled = Not YaEstoy
    '
    'Label1(1) = "&SubClaves encontradas: " & CStr(TreeView1.Nodes.Count)
    Label1(1) = CStr(nValoresHallados) & " valores hallados en " & CStr(nClavesHalladas) & " claves"
    If nValoresHallados = 0 Then
        txtClave = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\Cache\Servers"
        LeerClaves Trim$(txtClave)
        txtBuscar.Text = ""
        
    End If
    '
    MousePointer = vbDefault
    '
    DoEvents
End Sub

Private Sub cmdLeerClave_Click()
    ' Mostrar los valores de la clave especificada                  (31/Ene/99)
    ' también las subclaves, aunque no las subclaves de las subclaves...
    Static YaEstoy As Boolean
    Dim i As Long
    '
    If YaEstoy Then
        Cancelado = True
        DoEvents
    Else
        YaEstoy = True
        Cancelado = False
        '
        For i = 4 To 7
            Label1(i).Visible = False
        Next
        Label1(1) = "&SubClaves:"
        '
        cmdLeerClave.Tag = cmdLeerClave.Caption
        'cmdLeerClave.Caption = "Cancelar lectura"
        LeerClaves Trim$(txtClave), False
        cmdLeerClave.Caption = cmdLeerClave.Tag
        YaEstoy = False
        Cancelado = False
    End If
End Sub

Private Sub cmdLeerTodo_Click()
    LeerClaves Trim$(txtClave)
End Sub



Private Sub cmdRDP_Click()
    Dim lValDev As Long
    Dim rmd As String
    Dim params As String
    Me.ListView1.Refresh
    rmd = "mstsc"
    params = "/v:" + ListView1.ListItems(SWADDRESS).SubItems(1)
    
    lValDev = ShellExecute(Me.hwnd, "Open", rmd, params, "", 1)
        
        'mstsc /v:computername


End Sub

Private Sub cmdSalir_Click()
    Cancelado = True
    DoEvents
    Unload Me
End Sub

Private Sub cmdUnRegister_Click(Index As Integer)
    ' Info sobre el CLSID de una clase y el TypeLib                 (05/Jul/99)
    If Index = 0 Then
        ' Mostrar el CLSID y el TypeLib de esa clave
        txtUnregister(1) = tQR.ClassCLSID(txtUnregister(0))
        txtUnregister(2) = tQR.ClassTypeLib(txtUnregister(0))
    Else
        ' Eliminar del registro
        Dim tQRError As eHKEYError
        '
        If Len(txtUnregister(0).Text) = 0 Then
            MsgBox "¡¡¡No hay nada que borrar!!!", vbInformation + vbOKOnly, "Borrar clase de servidor ActiveX"
            txtUnregister(0).SetFocus
            Exit Sub
        End If
        ' Pedir confirmación antes de borrar                        (23/Sep/00)
        If MsgBox("¿Seguro que quieres borrar:" & vbCrLf & txtUnregister(0) & "?" & vbCrLf & vbCrLf & "¡Sólo se aviso esta vez!" & vbCrLf & "así que a ver que haces...", vbYesNo + vbDefaultButton2 + vbExclamation, "Borrar clase de servidor ActiveX") = vbYes Then
            tQRError = tQR.UnRegister(txtUnregister(0))
            If tQRError = ERROR_NONE Then
                txtUnregister(1) = "Claves borradas satisfactoriamente"
            Else
                txtUnregister(1) = "Error " & tQRError & " al borrar las claves"
            End If
            txtUnregister(2) = ""
        End If
    End If
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim path As String
    
    ' Posicionar en la parte superior de la pantalla                (23/Sep/00)
    Move 0, -30
    'Move (Screen.Width - Width) \ 2, -30
    '
    Caption = "TPD Server Connenct IRIS v" & CStr(App.Major) & "." & Format$(App.Minor, "00") & " - TPD Server Connect para IRIS"
    '
    ' Asignar el estilo y otras propiedades del TreeView            (25/Sep/00)
    With TreeView1
        .Indentation = 256
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .LabelEdit = tvwManual
    End With
    '
    ' Asignar las cabeceras del ListView1
    With ListView1.ColumnHeaders
        .Add 1, "Nombre", "Nombre", 1500 '1680
        .Add 2, "Valor", "Valor", 2500 '2720
    End With
    '
    ' Crear la clase de manejo del registro
    Set tQR = New cQueryReg
    
    ' Asignar un valor genérico
    'txtClave = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion"
    'txtClave = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\Cache\Servers"
    txtClave = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\IRIS\Configurations\IRISHealth\Directory"
    
    ' Borrar el contenido del resto de los controles                (23/Sep/00)
    txtValor = ""
    cboNombre.Clear
    For i = 0 To 2
        txtUnregister(i) = ""
    Next
    '
    For i = 4 To 7
        Label1(i).Visible = False
    Next
    '
    txtBuscar.Text = ""
    txtPoner.Text = ""
    chkCambiar.Value = vbUnchecked
    chkCambiar_Click
    '
    Show
    '
    ' Mostrar la ventana de aviso,                                  (30/Nov/00)
    ' se mostrará durante 60 segundos o hasta que se pulse en OK
    With fDelay
        .Caption = "TPD Server Connect IRIS"
        .Message = vbCrLf & vbCrLf & vbCrLf & vbCrLf & " TPD Server Connect IRIS" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                "                                       Toni Peña" & vbCrLf
                '" utilizalo sólo si sabes lo que estás haciendo." & vbCrLf & _
                '" Si eliminas o modificas claves, nombres o valores que Windows u otras aplicaciones utilizan, puede que no vuelvan a funcionar como deberían..." & vbCrLf & vbCrLf & _
                '" El que avisa..."
        .Delay = 6000
        .msgFontBold = True
        .msgFontSize = 13
        .Show vbModal
    End With
    'txtClave = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\IRIS\Configurations\IRISHealth\Directory"
    LeerClaves Trim$(txtClave)
    path = ListView1.ListItems(1).SubItems(SWPATH)
    Me.txtPath = path
    
    txtClave = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\Cache\Servers"
    LeerClaves Trim$(txtClave)

    Me.TreeView1.SetFocus

    If TreeView1.SelectedItem.Text = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Intersystems\Cache\Servers" Then
        cmbPORTAL.Enabled = False
        cmbSTUDIO.Enabled = False
        cmbTERMINAL.Enabled = False
    Else
        cmbPORTAL.Enabled = True
        cmbSTUDIO.Enabled = True
        cmbTERMINAL.Enabled = True
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tQR = Nothing
    Set frmServersRegistro = Nothing
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    ' No poder modificar la etiqueta
    Cancel = True
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    ' Mostrar el nombre y el valor, para poder editarlo
    cboNombre.Text = item.Text
    txtValor.Text = item.SubItems(1)
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
    ' No permitir modificar las etiquetas
    Cancel = True
End Sub

Private Sub LeerValores(ByVal sKey As String)
    ' Mostrar los valores de la clave especificada                  (31/Ene/99)
    '
    ' Mostrar en el ListView los nombres y valores
    '
    Dim i As Long
    ' Se utiliza el MSCOMCTL.OCX de la versión 6 de VB
    Dim tListItem As MSComctlLib.ListItem
    
    ' Enumerar los datos de la clave
    If tQR.EnumValues(ma_scolKeys(), sKey) Then
        With ListView1
            .ListItems.Clear
            .FullRowSelect = True
            .Sorted = True
        End With
        '
        txtValor = ""
        cboNombre.Clear
        cboNombre.Tag = sKey
        ' Asignar los valores al combo (nombres y valores)
        For i = 1 To UBound(ma_scolKeys) Step 2
            If Len(ma_scolKeys(i)) Then
                cboNombre.AddItem ma_scolKeys(i)
                ' Mostrar los nombres y valores de la clave actual
                Set tListItem = ListView1.ListItems.Add(, , ma_scolKeys(i))
                tListItem.SubItems(1) = ma_scolKeys(i + 1)
            End If
        Next
        If cboNombre.ListCount > 0 Then
            cboNombre.ListIndex = 0
        End If
        '
    End If
    ' Borrar el contenido del array
    ReDim ma_scolKeys(0)
End Sub

Private Sub LeerClaves(ByVal sKey As String, _
                       Optional ByVal ConSubClaves As Boolean = True)
    ' Leer todas los valores y subclaves de la clave indicada       (22/Nov/00)
    '
    ' Mostrar en el TreeView las claves
    ' y en el ListView los nombres y valores
    '
    Dim i As Long, j As Long, kk As Integer
    Dim sID As String
    Dim sID2 As String
    Dim srvKey As String
    ' Se utiliza el MSCOMCTL.OCX de la versión 6 de VB
    Dim tNode As MSComctlLib.Node
    Dim tNode2 As MSComctlLib.Node
    '
    MousePointer = vbArrowHourglass
    '
    ' Si tiene el separador del final, quitárselo                   (23/Nov/00)
    sKey = Trim$(sKey)
    If Right$(sKey, 1) = "\" Then
        sKey = Left$(sKey, Len(sKey) - 1)
    End If
    
    '---
    
     If InStr(sKey, "Directory") = 0 Then
        kk = InStrRev(sKey, "\")
        srvKey = Right$(sKey, Len(sKey) - kk)
    End If
    
    
    '---
    ' Enumerar los datos de la clave
    If tQR.EnumValues(ma_scolKeys(), sKey) Then
        TreeView1.Nodes.Clear
        '
        sID = sKey
        Set tNode = TreeView1.Nodes.Add(, , sID, sKey)
        ' De esta forma se podrá volver a "releer" esta clave
        tNode.Tag = sKey
        tNode.Sorted = True
        tNode.Expanded = True
        '
        LeerValores sKey
        '
        ' Mostrar las subclaves de la clave indicada
        Call tQR.EnumKeys(ma_scolKeys(), sKey)
        j = UBound(ma_scolKeys)
        TreeView1.Visible = False
        For i = 1 To j
            Label1(1) = "Asignando al árbol: " & CStr(i) & " de " & CStr(j)
            sID2 = sID & "\" & ma_scolKeys(i)
            Set tNode2 = TreeView1.Nodes.Add(sID, tvwChild, sID2, ma_scolKeys(i))
            tNode2.Tag = sID2
            tNode2.Sorted = True
            If ConSubClaves Then
                LeerSubClaves sID2
            End If
            'TreeView1.Refresh
            DoEvents
            If Cancelado Then Exit For
        Next
        TreeView1.Visible = True
        Label1(1) = "&SubClaves: (mostradas: " & CStr(TreeView1.Nodes.Count) & ")"
    End If
    '
    Cancelado = False
    DoEvents
    MousePointer = vbDefault
    ' Borrar el contenido del array
    ReDim ma_scolKeys(0)
End Sub

Private Sub LeerSubClaves(ByVal sKey As String)
    ' Leer las subclaves de la clave indicada,                      (22/Nov/00)
    ' y añadirlas a la rama indicada
    Dim i As Long
    ' Se utiliza el MSCOMCTL.OCX de la versión 6 de VB
    Dim tNode2 As MSComctlLib.Node
    '
    Dim asSubClaves() As String
    Dim sID2 As String
    '
    ' Mostrar las subclaves de la clave indicada
    If tQR.EnumKeys(asSubClaves(), sKey) Then
        For i = 1 To UBound(asSubClaves)
            sID2 = sKey & "\" & asSubClaves(i)
            Set tNode2 = TreeView1.Nodes.Add(sKey, tvwChild, sID2, asSubClaves(i))
            tNode2.Tag = sID2
            tNode2.Sorted = True
            LeerSubClaves sID2
        Next
    End If
End Sub


Private Sub TreeView1_Click()
    If InStr(TreeView1.SelectedItem.Text, "Servers") > 0 Then
        cmbPORTAL.Enabled = False
        cmbSTUDIO.Enabled = False
        cmbTERMINAL.Enabled = False
        cmdRDP.Enabled = False
        
        
    Else
        cmbPORTAL.Enabled = True
        cmbSTUDIO.Enabled = True
        cmbTERMINAL.Enabled = True
        cmdRDP.Enabled = True
    End If
End Sub

Private Sub TreeView1_DblClick()
    If InStr(TreeView1.SelectedItem.Text, "Servers") = 0 Then
        cmbPORTAL_Click
        cmbSTUDIO_Click
        cmbTERMINAL_Click
        cmdRDP_Click
    End If
End Sub

Private Sub TreeView1_GotFocus()

    If InStr(TreeView1.SelectedItem.Text, "Servers") > 0 Then
        cmbPORTAL.Enabled = False
        cmbSTUDIO.Enabled = False
        cmbTERMINAL.Enabled = False
        cmdRDP.Enabled = False
    Else
        cmbPORTAL.Enabled = True
        cmbSTUDIO.Enabled = True
        cmbTERMINAL.Enabled = True
        cmdRDP.Enabled = True
    End If

End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    ' Si se hace click en una rama del árbol
    Dim sKey As String
    '
    sKey = Node.Tag
    ' Mostrar los nombres y valores
    LeerValores sKey
    txtClave = sKey
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    ' Si se hace click en una rama del árbol
    Dim sKey As String
    '
    sKey = Node.Tag
    ' Mostrar los nombres y valores
    LeerValores sKey
    txtClave = sKey
End Sub

Private Sub TreeView1_Validate(Cancel As Boolean)
    Me.ListView1.Refresh
End Sub

Private Sub txtBuscar_GotFocus()
    With txtBuscar
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        
        KeyAscii = 0        ' Para que no "pite"
        Me.cmdBuscar.SetFocus
        'SendKeys "{tab}"    ' Envía una pulsación TAB
        cmdBuscar_Click
        Me.txtBuscar.SetFocus
        
    End If
End Sub

Private Sub txtPoner_GotFocus()
    With txtPoner
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtValor_GotFocus()
    With txtValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub BuscarValores(ByVal sKey As String, _
                          Optional ByVal elDT As eHKEYDataType = REG_SZ)
    '--------------------------------------------------------------------------
    ' Se buscarán los valores en la clave del tipo indicado
    '--------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim aDatos() As String
    Dim Hallado As Boolean
    'Dim sID As String
    Dim CambiarDato As Boolean
    '
    Dim lRet As eHKEYError
    Dim rDT As eHKEYDataType
    Dim k As Long
    Dim s As String
    Dim resp As VbMsgBoxResult
    '
    ReDim aDatos(0)
    '
'    sID = ""
    If EnClaves Then
        If InStr(1, sKey, sBuscar, TipoComparacion) Then
'            sID = sKey
            Hallado = True
        End If
    End If
    ' Enumerar los datos de la clave, (sólo los de cadena)
    If tQR.EnumValuesByType(aDatos(), sKey, elDT) Then
        j = UBound(aDatos)
        nValoresExaminados = nValoresExaminados + j \ 2
        Label1(7).Caption = Format$(nValoresExaminados, "###,###,###")
        For i = 1 To j Step 2
            'Hallado = False
            If EnNombres Then
                If InStr(1, aDatos(i), sBuscar, TipoComparacion) Then
                    Hallado = True
                    nValoresHallados = nValoresHallados + 1
                    '
                    If Reemplazar Then
                        ' se va a cambiar el valor
                        CambiarDato = True
                        If PedirConfirmacion Then
                            resp = MsgBox("¿Quieres cambiar este Dato?" & vbCrLf & "Nombre: " & aDatos(i) & vbCrLf & "Datos: " & aDatos(i + 1), vbQuestion + vbYesNoCancel, "Buscar / Reemplazar")
                            If resp = vbNo Then
                                CambiarDato = False
                            ElseIf resp = vbCancel Then
                                CambiarDato = False
                                Cancelado = True
                                Exit For
                            End If
                        End If
                        rDT = tQR.GetRegType(sKey, aDatos(i))
                        ' Cambiar cadenas y DWORD
                        Select Case rDT
                        Case REG_SZ, REG_EXPAND_SZ, REG_DWORD
                            ' OK
                        Case Else
                            CambiarDato = False
                        End Select
                        '
                        If CambiarDato Then
                            ' Se borrará el nombre y se añadirá el nuevo
                            If tQR.DeleteValue(sKey, aDatos(i)) = ERROR_SUCCESS Then
                                s = aDatos(i)
                                k = InStr(1, s, sBuscar, TipoComparacion)
                                s = Left$(s, k - 1) & sPoner & Mid$(s, k + Len(sBuscar))
                                lRet = tQR.SetReg(sKey, s, aDatos(i + 1), , rDT, True)
                                If lRet <> ERROR_SUCCESS Then
                                    MsgBox "ERROR al cambiar el dato: " & vbCrLf & "Nombre: " & aDatos(i) & vbCrLf & "Datos: " & aDatos(i + 1) & vbCrLf & "Clave: " & sKey, vbExclamation + vbOKOnly, "Buscar / Reemplazar"
                                    s = aDatos(i)
                                    ' Intentar restaurarla
                                    lRet = tQR.SetReg(sKey, s, aDatos(i + 1), , rDT, True)
                                End If
                                ' No cambiar el nombre hasta que se sepa que todo va bien
                                aDatos(i) = s
                            End If
                        End If
                    End If
                    '
                End If
            End If
            If EnDatos = True Then
                If InStr(1, aDatos(i + 1), sBuscar, TipoComparacion) Then
                    Hallado = True
                    nValoresHallados = nValoresHallados + 1
                    '
                    If Reemplazar Then
                        ' se va a cambiar el valor
                        CambiarDato = True
                        If PedirConfirmacion Then
                            resp = MsgBox("¿Quieres cambiar este Dato?" & vbCrLf & "Nombre: " & aDatos(i) & vbCrLf & "Datos: " & aDatos(i + 1), vbQuestion + vbYesNoCancel, "Buscar / Reemplazar")
                            If resp = vbNo Then
                                CambiarDato = False
                            ElseIf resp = vbCancel Then
                                CambiarDato = False
                                Cancelado = True
                                Exit For
                            End If
                        End If
                        '
                        rDT = tQR.GetRegType(sKey, aDatos(i))
                        ' Cambiar cadenas y DWORD
                        Select Case rDT
                        Case REG_SZ, REG_EXPAND_SZ, REG_DWORD
                            ' OK
                        Case Else
                            CambiarDato = False
                        End Select
                        '
                        If CambiarDato Then
                            s = aDatos(i + 1)
                            k = InStr(1, s, sBuscar, TipoComparacion)
                            s = Left$(s, k - 1) & sPoner & Mid$(s, k + Len(sBuscar))
                            lRet = tQR.SetReg(sKey, aDatos(i), s, , rDT)
                            If lRet <> ERROR_SUCCESS Then
                                MsgBox "ERROR al cambiar el dato: " & vbCrLf & "Nombre: " & aDatos(i) & vbCrLf & "Datos: " & aDatos(i + 1) & vbCrLf & "Clave: " & sKey, vbExclamation + vbOKOnly, "Buscar / Reemplazar"
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            If Cancelado Then Exit For
        Next
        ' Si se ha hallado algún dato...
        If Hallado Then
            On Error Resume Next
            Err = 0
            With TreeView1
                With .Nodes.Add(, , sKey, sKey)
                    ' De esta forma se podrá volver a "releer" esta clave
                    .Tag = sKey
                    .Sorted = True
                    .Expanded = True
                End With
            End With
            If Err = 0 Then
                nClavesHalladas = nClavesHalladas + 1
            Else
                Err = 0
            End If
            On Error GoTo 0
        End If
    End If
    '
End Sub

Private Sub BuscarClaves(ByVal sKey As String)
    Dim i As Long, j As Long
    Dim aClaves() As String
    Static PrimeraVez As Boolean
    Static n As Long
    Dim Asignado As Boolean
    Dim srvKey As String
    Dim kk As Integer
    '
    On Error GoTo ErrorBuscarClaves
    '
    ReDim aClaves(0)
    ' Enumerar las subclaves
    If tQR.EnumKeys(aClaves(), sKey) Then
        j = UBound(aClaves)
        If PrimeraVez = False Then
            ' Para una explicación de que va esto, ver más abajo
            PrimeraVez = True
            Asignado = True
            nClavesExaminadas = 0
        End If
        '
        If EnClaves Then
            If InStr(1, sKey, sBuscar, TipoComparacion) Then
                On Error Resume Next
                Err = 0
                If InStr(sKey, "Directory") = 0 Then
                    kk = InStrRev(sKey, "\")
                    srvKey = Right$(sKey, Len(sKey) - kk)
                End If
                With TreeView1
                    With .Nodes.Add(, , sKey, srvKey)
                        ' De esta forma se podrá volver a "releer" esta clave
                        .Tag = sKey
                        .Sorted = True
                        .Expanded = True
                    End With
                End With
                If Err = 0 Then
                    nClavesHalladas = nClavesHalladas + 1
                Else
                    Err = 0
                End If
                On Error GoTo ErrorBuscarClaves
            End If
        End If
        '
        nClavesExaminadas = nClavesExaminadas + j
        Label1(5).Caption = Format$(nClavesExaminadas, "###,###,###")
        Label1(1) = CStr(nValoresHallados) & " valores hallados en " & CStr(nClavesHalladas) & " claves"
        For i = 1 To j
            If Len(aClaves(i)) Then
                'Label1(1) = aClaves(i)
                ' Sólo si se debe buscar en Nombres y Datos
                ' antes se hacía la comprobación dentro de BuscarValores,
                ' de esta forma sólo hacemos una, en lugar de dos,
                ' (una para cada tipo)
                If EnNombres = True Or EnDatos = True Then
                    ' Buscar en los valores y subclaves de tipos Cadena y DWORD
                    BuscarValores sKey & "\" & aClaves(i), REG_SZ
                    BuscarValores sKey & "\" & aClaves(i), REG_DWORD
                End If
                DoEvents
                If Cancelado Then Exit For
                BuscarClaves sKey & "\" & aClaves(i)
            End If
            DoEvents
            If Cancelado Then Exit For
        Next
        If Asignado Then
            ' Esto sólo se cumplirá con la clave que entre por primera vez
            ' Explicación:
            '   Al ser un procedimiento recursivo (se llama a sí mismo),
            '   las variables estáticas está disponibles para todas las veces que se entre,
            '   pero las normales, sólo mantendrán el valor en cada una de las veces
            '   que se entre en el procedimiento, por tanto "Asignado" sólo será TRUE
            '   la primera vez que se entre, que es cuando "PrimeraVez" vale FALSE
            '   Por tanto Asignado será True sólo cuando se haya asignado PrimeraVez,
            '   así que usamos ese valor para volver a poner "PrimeraVez" a FALSE,
            '   con idea de que cuando se vuelvan a buscar nuevos datos se pueda
            '   volver a usar este "truco" para que lo que se muestra en la etiqueta
            '   no se cambie cada vez que se entra en este procedimiento.
            '   De nada, es que esto de la recursividad es un poco "lioso".
            '
            ' Ponemos los valores estáticos a cero cuando sale el primero que entró
            PrimeraVez = False
            nClavesExaminadas = 0
        End If
    End If
    Exit Sub
    '
ErrorBuscarClaves:
    If MsgBox("Se ha producido el error: " & CStr(Err.Number) & " " & Err.Description & vbCrLf & "con " & CStr(nClavesExaminadas) & " y " & CStr(j) & " subclaves en esta ronda." & vbCrLf & vbCrLf & "La clave que se está examinando es: " & vbCrLf & sKey & vbCrLf & vbCrLf & "¿Quieres terminar el programa?", vbInformation + vbYesNo, "Buscar / Reemplazar") = vbYes Then
        Cancelado = True
        DoEvents
        Unload Me
    End If
End Sub


'Private Function ServerVar(index As Integer) As String
'    Dim tItem As ListItem
'    Dim lvwFind As ListFindItemHowConstants
'    Dim lvwWhere As ListFindItemWhereConstants
'    Dim i As Integer
'    Dim col As Integer
'    Dim item As String
'    ' realizamos la búsqueda
'
'    'For i = 2 To ListView1.ListItems.Count
'        i = index
'        item = ListView1.ListItems(i).Text
'        ServerVar = ListView1.ListItems(i).Text
'        Select Case i
'            Case SWADDRESS
'                Adress = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = Adress
'            Case SWAUTHENTICATIONMETHOD
'                AuthenticationMethod = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = AuthenticationMethod
'            Case SWCOMMENT
'                Comment = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = Comment
'            Case SWCONNECTIONSECURITYLEVEL
'                ConnectionSecurityLevel = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = ConnectionSecurityLevel
'            Case SWHTTPS
'                HTTPS = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = AdrHTTPSss
'            Case SWPORT
'                Port = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = Port
'            Case SWSERVERTYPE
'                ServerType = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = ServerType
'            Case SWSERVICEPRINCIPALNAME
'                ServicePrincipalName = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = ServicePrincipalName
'            Case SWTELNET
'                Telnet = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = Telnet
'            Case SWWEBSERVERADDRESS
'                WebServerAddress = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = WebServerAddress
'            Case SWWEBSERVERINSTANCENAME
'                WebServerInstanceName = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = WebServerInstanceName
'            Case SWWEBSERVERPORT
'                WebServerPort = ListView1.ListItems(i).SubItems(1)
'                ServerVariables = WebServerPort
'       End Select
'    Next
'
'    'Set tItem = ListView1.FindItem("1972", lvwWhere, 1, lvwFind)
'    'Adress = ListView1.ListItems(2).SubItems(1)
'
'End Function

Private Sub cmbPORTAL_Click()
    Dim url As String
    Dim X As String
    Me.ListView1.Refresh
    
    url = "http://" + ListView1.ListItems(SWADDRESS).SubItems(1) + ":" + ListView1.ListItems(SWWEBSERVERPORT).SubItems(1) + "/csp/sys/UtilHome.csp"
    X = ShellExecute(Me.hwnd, "Open", url, &O0, &O0, SW_NORMAL)

End Sub


Private Sub cmbTERMINAL_Click()

    Dim lValDev As Long
    Dim terminal As String
    Dim params As String
    Me.ListView1.Refresh
    terminal = Me.txtPath.Text + "\bin\Iristerm.exe"
    params = "/server=" + TreeView1.SelectedItem.Text
    
    lValDev = ShellExecute(Me.hwnd, "Open", terminal, params, "", 1)

End Sub

