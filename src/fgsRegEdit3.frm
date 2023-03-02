VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fgsRegEdit3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "gsRegEdit v3.00 - Utilidad para manipular el Registro de Windows"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "fgsRegEdit3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11250
   Begin VB.Frame Frame1 
      Caption         =   "Des-registrar servidor ActiveX:"
      Height          =   1515
      Index           =   3
      Left            =   120
      TabIndex        =   32
      Top             =   6450
      Width           =   9555
      Begin VB.TextBox txtUnregister 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   34
         Text            =   "Servidor.Clase"
         Top             =   330
         Width           =   5025
      End
      Begin VB.TextBox txtUnregister 
         Height          =   285
         Index           =   1
         Left            =   810
         TabIndex        =   37
         Text            =   "Servidor.Clase"
         Top             =   690
         Width           =   6795
      End
      Begin VB.CommandButton cmdUnRegister 
         Caption         =   "Mostrar info"
         Height          =   645
         Index           =   0
         Left            =   7710
         TabIndex        =   35
         ToolTipText     =   " Mostrar informaci�n sobre la clase "
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdUnRegister 
         Caption         =   "Eliminar del Registro"
         Height          =   405
         Index           =   1
         Left            =   7710
         TabIndex        =   40
         ToolTipText     =   " Eliminar la clase mostrada del registro �CUIDADO! "
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtUnregister 
         Height          =   285
         Index           =   2
         Left            =   810
         TabIndex        =   39
         Text            =   "Servidor.Clase"
         Top             =   1080
         Width           =   6795
      End
      Begin VB.Label Label1 
         Caption         =   "Clase, en formato Servidor.Clase:"
         Height          =   255
         Index           =   12
         Left            =   180
         TabIndex        =   33
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "CLSID:"
         Height          =   255
         Index           =   13
         Left            =   180
         TabIndex        =   36
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "TypeLib:"
         Height          =   255
         Index           =   14
         Left            =   180
         TabIndex        =   38
         Top             =   1110
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   405
      Left            =   9810
      TabIndex        =   41
      ToolTipText     =   " Terminar el programa "
      Top             =   7560
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asignaci�n y lectura de claves y valores de una entrada del registro:"
      Height          =   6135
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   10935
      Begin VB.Frame Frame1 
         Caption         =   "&Buscar / Reemplazar valores:"
         Height          =   2775
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   3210
         Width           =   4905
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "fgsRegEdit3.frx":0442
            Top             =   2295
            Width           =   4665
         End
         Begin VB.CheckBox chkConfirmar 
            Caption         =   "Solicitar confirmaci�n"
            Height          =   255
            Left            =   2340
            TabIndex        =   18
            ToolTipText     =   " Si se pide confirmaci�n al reemplazar "
            Top             =   1230
            Value           =   1  'Checked
            Width           =   2145
         End
         Begin VB.CheckBox chkTipoComparacion 
            Caption         =   "May�sculas / Min�sculas"
            Height          =   255
            Left            =   2340
            TabIndex        =   17
            ToolTipText     =   " Comprueba las palabras teniendo en cuenta may�sculas y min�sculas "
            Top             =   1500
            Width           =   2235
         End
         Begin VB.Frame Frame1 
            Caption         =   "Buscar en:"
            Height          =   1125
            Index           =   4
            Left            =   150
            TabIndex        =   13
            Top             =   1140
            Width           =   1995
            Begin VB.CheckBox chkBuscarEn 
               Caption         =   "Datos"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   16
               Top             =   780
               Value           =   1  'Checked
               Width           =   1665
            End
            Begin VB.CheckBox chkBuscarEn 
               Caption         =   "Valores (nombres)"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   15
               Top             =   510
               Value           =   1  'Checked
               Width           =   1665
            End
            Begin VB.CheckBox chkBuscarEn 
               Caption         =   "Claves"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   1665
            End
         End
         Begin VB.CheckBox chkCambiar 
            Caption         =   "Cambiar por:"
            Height          =   315
            Left            =   150
            TabIndex        =   11
            ToolTipText     =   " Marca esta casilla para reemplazar "
            Top             =   750
            Width           =   1245
         End
         Begin VB.TextBox txtPoner 
            Height          =   315
            Left            =   1410
            TabIndex        =   12
            Text            =   "txtPoner"
            Top             =   750
            Width           =   3345
         End
         Begin VB.TextBox txtBuscar 
            Height          =   315
            Left            =   1410
            TabIndex        =   10
            Text            =   "txtBuscar"
            Top             =   360
            Width           =   3345
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar / Reemplazar"
            Height          =   375
            Left            =   3090
            TabIndex        =   20
            ToolTipText     =   " Buscar/Reemplazar a partir de la clave indicada (y en las subclaves) "
            Top             =   1860
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Buscar:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   390
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdLeerTodo 
         Caption         =   "Mostrar todas las subclaves"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         ToolTipText     =   " Mostrar todas las subclaves por debajo de la calve indicada, (Esto puerde tardar un buen rato) "
         Top             =   720
         Width           =   2505
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1905
         Left            =   5100
         TabIndex        =   21
         Top             =   1230
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   1905
         Left            =   150
         TabIndex        =   7
         Top             =   1230
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   3360
         _Version        =   393217
         Indentation     =   847
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de valor a asignar/borrar:"
         Height          =   1575
         Index           =   1
         Left            =   5220
         TabIndex        =   26
         Top             =   4410
         Width           =   4455
         Begin VB.OptionButton optTipo 
            Caption         =   "String (cadena ampliada) "
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   30
            ToolTipText     =   " REG_MULTI_SZ  (no implementada) "
            Top             =   1170
            Width           =   2355
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "DWORD (num�rico) "
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   29
            ToolTipText     =   " REG_DWORD "
            Top             =   870
            Width           =   2355
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Binary (binario) "
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   28
            ToolTipText     =   " REG_BINARY "
            Top             =   570
            Width           =   2355
         End
         Begin VB.CommandButton cmdAsignar 
            Caption         =   "Asignar el valor"
            Height          =   375
            Left            =   2610
            TabIndex        =   31
            ToolTipText     =   " Asigna el valor indicado al nombre de la clave "
            Top             =   1050
            Width           =   1695
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "String (cadena) "
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   27
            ToolTipText     =   " REG_SZ, REG_EXPAND_SZ "
            Top             =   270
            Value           =   -1  'True
            Width           =   2355
         End
      End
      Begin VB.CommandButton cmdBorrarValor 
         Caption         =   "Borrar el Nombre"
         Height          =   375
         Left            =   9090
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   " Borra el nombre y el valor "
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox cboNombre 
         Height          =   315
         Left            =   6480
         TabIndex        =   23
         Text            =   "cboNombre"
         Top             =   3210
         Width           =   4305
      End
      Begin VB.CommandButton cmdLeerClave 
         Caption         =   "Mostrar las subclaves"
         Default         =   -1  'True
         Height          =   375
         Left            =   8280
         TabIndex        =   3
         ToolTipText     =   " Mostrar los valores y las subclaves contenidas en la Clave indicada "
         Top             =   720
         Width           =   2505
      End
      Begin VB.CommandButton cmdBorrarClave 
         Caption         =   "Borrar la clave"
         Height          =   375
         Left            =   3450
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   " Borrar la clave indicada y todas las subclaves �PRECAUCI�N! "
         Top             =   720
         Width           =   2025
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   5220
         TabIndex        =   25
         Text            =   "Valor"
         Top             =   4050
         Width           =   5565
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         Left            =   1020
         TabIndex        =   2
         Text            =   "HKEY_USERS\.Default\Software\elGuille\Pruebas Registro"
         Top             =   330
         Width           =   9765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "999.999.999"
         Height          =   255
         Index           =   7
         Left            =   2430
         TabIndex        =   45
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Valores:"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   44
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "999.999.999"
         Height          =   255
         Index           =   5
         Left            =   750
         TabIndex        =   43
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Claves:"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   42
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "&SubClaves:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   " Lista de Claves y subclaves "
         Top             =   990
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "&Nombre / Valor:"
         Height          =   255
         Index           =   2
         Left            =   5190
         TabIndex        =   22
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "&Clave:"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   360
         Width           =   645
      End
   End
End
Attribute VB_Name = "fgsRegEdit3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Prueba de manipulaci�n del Registro                               (13/Ago/98)
'
' Con pruebas de creaci�n/modificaci�n/borrado                (01:50 18/Ago/98)
'
' Nueva versi�n para leer y modificar entradas del registro         (31/Ene/99)
' Nueva versi�n, usando Treeview y ListView                         (23/Sep/00)
' Mostrando un formulario de AVISO                                  (30/Nov/00)
'
' Nueva versi�n                                                     (28/Dic/01)
'   permite cambiar valores (Buscar/Reemplazar)
'   Algunas confirmaciones extras antes de borrar...
'
' �Guillermo 'guille' Som, 1998-2001
'------------------------------------------------------------------------------
Option Explicit
'Option Compare Text

' Pongo estas variables a nivel de m�dulo para mayor rapidez
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
            cmdBuscar.Caption = "Buscar"
        End If
    End With
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
    ' Asignar s�lo si el nombre de la clave tiene algo escrito          (31/Ene/99)
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
            ' Este tipo ser� un array de cadenas,                   (22/Nov/00)
            ' en el valor se indicar� cada cadena separada por punto y coma
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
    If MsgBox("�Quieres borrar esta clave:" & vbCrLf & sKey & "?", vbCritical + vbYesNo + vbDefaultButton2, "Borrar clave del registro del sistema") = vbYes Then
        ' Pedir una segunda confirmaci�n... vayamos a p***cas       (28/Dic/01)
        If MsgBox("�Seguro que quieres borrar la clave" & vbCrLf & sKey & "?" & vbCrLf & vbCrLf & "�Que pesado soy! �verdad?" & vbCrLf & "Pero es que no es plan de jugar con estas cosas..." & vbCrLf & vbCrLf & "A�n est�s a tiempo de pulsar en NO..." & vbCrLf & "pero si a�n quieres borrar esa clave, pulsa en SI..." & vbCrLf & "�tu sabr�s lo que haces!", vbCritical + vbYesNo + vbDefaultButton2, "Borrar clave del registro del sistema") = vbYes Then
            If InStr(1, sKey, "Software\Microsoft\Windows", vbTextCompare) Then
                ' Esta clave es vital...
                If MsgBox("PSSST!!!" & vbCrLf & vbCrLf & "�Sabes que esa clave es importante?" & vbCrLf & vbCrLf & "Ya no te advierto m�s..." & vbCrLf & vbCrLf & "�Quieres borrarla de todas formas?", vbCritical + vbYesNo + vbDefaultButton2, "Borrar clave del registro del sistema") = vbNo Then
                    MsgBox "Te has hecho de rogar... pero... ���al fin has visto la luz!!!", vbInformation + vbOKOnly, "Menos mal que no la has borrado"
                    Exit Sub
                End If
            End If
            lRet = tQR.DeleteKey(sKey)
            If lRet = ERROR_SUCCESS Then
                MsgBox "Se ha borrado la clave: " & sKey & ", con �xito.", vbInformation, "Borrar clave del registro"
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
        MsgBox "No puedes dejar el valor del Nombre en blanco..." & vbCrLf & "ya que se borrar�a la clave completa..." & vbCrLf & "Si quieres borrar la clave, usa el bot�n correspondiente... pero... �cuidado con lo que haces!", vbExclamation + vbOKOnly, "Borrar un valor"
        txtClave.SetFocus
        Exit Sub
    End If
    ' Pedir confirmaci�n antes de borrar
    If MsgBox("�Quieres borrar la entrada: " & sName & " de la clave:" & vbCrLf & sKey & "?", vbQuestion + vbDefaultButton2 + vbYesNo, "Borrar una entrada del registro") = vbYes Then
        lRet = tQR.DeleteKey(sKey, sName)
        If lRet = ERROR_SUCCESS Then
            sMsg = "Se ha borrado el nombre: '" & sName & "' de la clave con �xito."
        Else
            lMsg = vbExclamation
            sMsg = "No se ha podido borrar el nombre: '" & sName & "' de la clave."
        End If
        MsgBox sMsg, lMsg, "Borrar nombre de una clave del registro"
    End If
End Sub

Private Sub cmdBuscar_Click()
    ' Buscar o Reemplazar                                           (28/Dic/01)
    ' Avisar de que se debe tener cuidado con esto de reemplazar autom�ticamente
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
        MsgBox "Debes especificar lo que quieres buscar.", vbInformation + vbOKOnly, "Buscar / Reemplazar"
        txtBuscar.SetFocus
        Exit Sub
    End If
    '
    If chkCambiar.Value = vbChecked Then
        If Len(sPoner) = 0 Then
            MsgBox "Debes indicar el texto que reemplazar� lo buscado.", vbInformation + vbOKOnly, "Buscar / Reemplazar"
            txtPoner.SetFocus
            Exit Sub
        End If
        Reemplazar = True
        PedirConfirmacion = (chkConfirmar.Value = vbChecked)
        If PedirConfirmacion = False Then
            s = "Adem�s no has marcado la opci�n de pedir confirmaci�n."
        End If
        If MsgBox("�ATENCI�N!" & vbCrLf & "Es muy importante que sepas que cambiando los valores de forma autom�tica," & vbCrLf & "se pueden producir efectos no deseados en el registro." & vbCrLf & vbCrLf & "�Seguro que quieres reemplazar?" & vbCrLf & vbCrLf & s, vbExclamation + vbYesNo + vbDefaultButton2, "Buscar / Reemplazar") = vbNo Then
            Exit Sub
        End If
        If PedirConfirmacion = False Then
            If MsgBox("Perdona que me haga pesado, pero..." & vbCrLf & "�SABES LO PELIGROSO QUE ES CAMBIAR DE FORMA AUTOM�TICA EL REGISTRO?" & vbCrLf & vbCrLf & "A�n despu�s de este aviso a la cordura..." & vbCrLf & "�Quieres continuar con la opci�n de reemplazar?" & vbCrLf & "(te recuerdo que se har� sin pedir confirmaci�n)", vbExclamation + vbYesNo + vbDefaultButton2, "Buscar / Reemplazar") = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Que conste que te he avisado..." & vbCrLf & "No me hago responsable de lo que pueda pasar con el registro..." & vbCrLf & vbCrLf & "A�n tienes una �ltima oportunidad de cancelar.", vbCritical + vbOKCancel + vbDefaultButton2, "Buscar / Reemplazar") = vbCancel Then
            Exit Sub
        End If
    End If
    '
    ' Buscar a partir de la clave indicada en txtClave y en las subclaves
    '
    ' Guardar los valores actuales en un fichero de configuraci�n
    'TODO:
    '
    YaEstoy = True
    Cancelado = False
    cmdBuscar.Caption = "Cancelar"
    MousePointer = vbArrowHourglass
    DoEvents
    '
    ' Deshabilitar todos los dem�s botones
    cmdAsignar.Enabled = Not YaEstoy
    cmdBorrarValor.Enabled = Not YaEstoy
    cmdBorrarClave.Enabled = Not YaEstoy
    cmdLeerTodo.Enabled = Not YaEstoy
    cmdLeerClave.Enabled = Not YaEstoy
    cmdUnRegister(0).Enabled = Not YaEstoy
    cmdUnRegister(1).Enabled = Not YaEstoy
    ' No deshabilitar el Frame1(2), sino �no se podr� pulsar en Cancelar!
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
    ' Si tiene el separador del final, quit�rselo                   (23/Nov/00)
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
        cmdBuscar.Caption = "Reemplazar"
    Else
        cmdBuscar.Caption = "Buscar"
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
    ' No deshabilitar el Frame1(2), sino �no se podr� pulsar en Cancelar!
    For i = 1 To 3 Step 2
        Frame1(i).Enabled = Not YaEstoy
    Next
    Frame1(i - 1).Enabled = Not YaEstoy
    '
    'Label1(1) = "&SubClaves encontradas: " & CStr(TreeView1.Nodes.Count)
    Label1(1) = CStr(nValoresHallados) & " valores hallados en " & CStr(nClavesHalladas) & " claves"
    '
    MousePointer = vbDefault
    '
    DoEvents
End Sub

Private Sub cmdLeerClave_Click()
    ' Mostrar los valores de la clave especificada                  (31/Ene/99)
    ' tambi�n las subclaves, aunque no las subclaves de las subclaves...
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
        cmdLeerClave.Caption = "Cancelar lectura"
        LeerClaves Trim$(txtClave), False
        cmdLeerClave.Caption = cmdLeerClave.Tag
        YaEstoy = False
        Cancelado = False
    End If
End Sub

Private Sub cmdLeerTodo_Click()
    LeerClaves Trim$(txtClave)
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
            MsgBox "���No hay nada que borrar!!!", vbInformation + vbOKOnly, "Borrar clase de servidor ActiveX"
            txtUnregister(0).SetFocus
            Exit Sub
        End If
        ' Pedir confirmaci�n antes de borrar                        (23/Sep/00)
        If MsgBox("�Seguro que quieres borrar:" & vbCrLf & txtUnregister(0) & "?" & vbCrLf & vbCrLf & "�S�lo se aviso esta vez!" & vbCrLf & "as� que a ver que haces...", vbYesNo + vbDefaultButton2 + vbExclamation, "Borrar clase de servidor ActiveX") = vbYes Then
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

Private Sub Form_Load()
    Dim i As Long
    
    ' Posicionar en la parte superior de la pantalla                (23/Sep/00)
    Move 0, -30
    'Move (Screen.Width - Width) \ 2, -30
    '
    Caption = "gsRegEdit v" & CStr(App.Major) & "." & Format$(App.Minor, "00") & " - Utilidad para manipular el Registro de Windows"
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
        .Add 1, "Nombre", "Nombre", 2130 '1680
        .Add 2, "Valor", "Valor", 3195 '2720
    End With
    '
    ' Crear la clase de manejo del registro
    Set tQR = New cQueryReg
    
    ' Asignar un valor gen�rico
    txtClave = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion"
    
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
    ' se mostrar� durante 60 segundos o hasta que se pulse en OK
    With fDelay
        .Caption = "Aviso importante"
        .Message = " ��� AVISO MUY, MUY, PERO QUE MUY IMPORTANTE !!!" & vbCrLf & vbCrLf & _
                " No juegues con el registro," & vbCrLf & _
                " utilizalo s�lo si sabes lo que est�s haciendo." & vbCrLf & _
                " Si eliminas o modificas claves, nombres o valores que Windows u otras aplicaciones utilizan, puede que no vuelvan a funcionar como deber�an..." & vbCrLf & vbCrLf & _
                " El que avisa..."
        .Delay = 60000
        .msgFontBold = True
        .msgFontSize = 13
        .Show vbModal
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tQR = Nothing
    Set fgsRegEdit3 = Nothing
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    ' No poder modificar la etiqueta
    Cancel = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' Mostrar el nombre y el valor, para poder editarlo
    cboNombre.Text = Item.Text
    txtValor.Text = Item.SubItems(1)
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
    ' Se utiliza el MSCOMCTL.OCX de la versi�n 6 de VB
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
    Dim i As Long, j As Long
    Dim sID As String
    Dim sID2 As String
    ' Se utiliza el MSCOMCTL.OCX de la versi�n 6 de VB
    Dim tNode As MSComctlLib.Node
    Dim tNode2 As MSComctlLib.Node
    '
    MousePointer = vbArrowHourglass
    '
    ' Si tiene el separador del final, quit�rselo                   (23/Nov/00)
    sKey = Trim$(sKey)
    If Right$(sKey, 1) = "\" Then
        sKey = Left$(sKey, Len(sKey) - 1)
    End If
    ' Enumerar los datos de la clave
    If tQR.EnumValues(ma_scolKeys(), sKey) Then
        TreeView1.Nodes.Clear
        '
        sID = sKey
        Set tNode = TreeView1.Nodes.Add(, , sID, sKey)
        ' De esta forma se podr� volver a "releer" esta clave
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
            Label1(1) = "Asignando al �rbol: " & CStr(i) & " de " & CStr(j)
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
    ' y a�adirlas a la rama indicada
    Dim i As Long
    ' Se utiliza el MSCOMCTL.OCX de la versi�n 6 de VB
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

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    ' Si se hace click en una rama del �rbol
    Dim sKey As String
    '
    sKey = Node.Tag
    ' Mostrar los nombres y valores
    LeerValores sKey
    txtClave = sKey
End Sub

Private Sub txtBuscar_GotFocus()
    With txtBuscar
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
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
    ' Se buscar�n los valores en la clave del tipo indicado
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
    ' Enumerar los datos de la clave, (s�lo los de cadena)
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
                            resp = MsgBox("�Quieres cambiar este Dato?" & vbCrLf & "Nombre: " & aDatos(i) & vbCrLf & "Datos: " & aDatos(i + 1), vbQuestion + vbYesNoCancel, "Buscar / Reemplazar")
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
                            ' Se borrar� el nombre y se a�adir� el nuevo
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
                            resp = MsgBox("�Quieres cambiar este Dato?" & vbCrLf & "Nombre: " & aDatos(i) & vbCrLf & "Datos: " & aDatos(i + 1), vbQuestion + vbYesNoCancel, "Buscar / Reemplazar")
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
        ' Si se ha hallado alg�n dato...
        If Hallado Then
            On Error Resume Next
            Err = 0
            With TreeView1
                With .Nodes.Add(, , sKey, sKey)
                    ' De esta forma se podr� volver a "releer" esta clave
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
    '
    On Error GoTo ErrorBuscarClaves
    '
    ReDim aClaves(0)
    ' Enumerar las subclaves
    If tQR.EnumKeys(aClaves(), sKey) Then
        j = UBound(aClaves)
        If PrimeraVez = False Then
            ' Para una explicaci�n de que va esto, ver m�s abajo
            PrimeraVez = True
            Asignado = True
            nClavesExaminadas = 0
        End If
        '
        If EnClaves Then
            If InStr(1, sKey, sBuscar, TipoComparacion) Then
                On Error Resume Next
                Err = 0
                With TreeView1
                    With .Nodes.Add(, , sKey, sKey)
                        ' De esta forma se podr� volver a "releer" esta clave
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
                ' S�lo si se debe buscar en Nombres y Datos
                ' antes se hac�a la comprobaci�n dentro de BuscarValores,
                ' de esta forma s�lo hacemos una, en lugar de dos,
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
            ' Esto s�lo se cumplir� con la clave que entre por primera vez
            ' Explicaci�n:
            '   Al ser un procedimiento recursivo (se llama a s� mismo),
            '   las variables est�ticas est� disponibles para todas las veces que se entre,
            '   pero las normales, s�lo mantendr�n el valor en cada una de las veces
            '   que se entre en el procedimiento, por tanto "Asignado" s�lo ser� TRUE
            '   la primera vez que se entre, que es cuando "PrimeraVez" vale FALSE
            '   Por tanto Asignado ser� True s�lo cuando se haya asignado PrimeraVez,
            '   as� que usamos ese valor para volver a poner "PrimeraVez" a FALSE,
            '   con idea de que cuando se vuelvan a buscar nuevos datos se pueda
            '   volver a usar este "truco" para que lo que se muestra en la etiqueta
            '   no se cambie cada vez que se entra en este procedimiento.
            '   De nada, es que esto de la recursividad es un poco "lioso".
            '
            ' Ponemos los valores est�ticos a cero cuando sale el primero que entr�
            PrimeraVez = False
            nClavesExaminadas = 0
        End If
    End If
    Exit Sub
    '
ErrorBuscarClaves:
    If MsgBox("Se ha producido el error: " & CStr(Err.Number) & " " & Err.Description & vbCrLf & "con " & CStr(nClavesExaminadas) & " y " & CStr(j) & " subclaves en esta ronda." & vbCrLf & vbCrLf & "La clave que se est� examinando es: " & vbCrLf & sKey & vbCrLf & vbCrLf & "�Quieres terminar el programa?", vbInformation + vbYesNo, "Buscar / Reemplazar") = vbYes Then
        Cancelado = True
        DoEvents
        Unload Me
    End If
End Sub
