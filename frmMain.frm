VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER SYSTEM"
   ClientHeight    =   5295
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datChequeAbogado 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cheque_abogado"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datCheque 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cheque_cliente"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datPago 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "factura_cliente"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame frmGerente 
      Caption         =   "Gerente Administrativo"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   10215
      Begin VB.CommandButton cmdDptos 
         Caption         =   "Asignar porcentajes de Departamentos"
         Height          =   735
         Left            =   8520
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ver historial abogados"
         Height          =   735
         Left            =   5160
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cambio de jerarquia a abogado"
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Agregar abogado"
         Height          =   735
         Left            =   3480
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar titulos a abogados"
         Height          =   735
         Left            =   6840
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Titulos de abogados"
         Height          =   735
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Data datAbogado 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame frmCaja 
      Caption         =   "Caja"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   8040
      TabIndex        =   22
      Top             =   480
      Width           =   2295
      Begin VB.CommandButton cmdCobros 
         Caption         =   "Cobros"
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdGestion 
      Caption         =   "Gestion de Usuarios"
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame frmRecepcion 
      Caption         =   "Recepcion"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   10215
      Begin VB.CommandButton cmdVerAbogados 
         Caption         =   "Registro de abogados"
         Height          =   735
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdNuevaCita 
         Caption         =   "Registrar nueva cita"
         Height          =   735
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdNuevoCliente 
         Caption         =   "Registrar nuevo cliente"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdVerCitas 
         Caption         =   "Registro de citas"
         Height          =   735
         Left            =   8280
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdVerCientes 
         Caption         =   "Registro de clientes"
         Height          =   735
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame frmContaduria 
      Caption         =   "Contaduria"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   7815
      Begin VB.CommandButton cmdCheques 
         Caption         =   "Registro de cheques"
         Height          =   735
         Left            =   5880
         TabIndex        =   32
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdPagosAbogados 
         Caption         =   "Pagos a abogados"
         Height          =   735
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdPagos 
         Caption         =   "Registro de pagos"
         Height          =   735
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdDeudores 
         Caption         =   "Deudores"
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmAbogado 
      Caption         =   "Abogado"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   10215
      Begin VB.CommandButton cmdNuevoCaso 
         Caption         =   "Iniciar nuevo caso"
         Height          =   735
         Left            =   7680
         TabIndex        =   25
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdCalificacion 
         Caption         =   "Calificacion de abogados"
         Height          =   735
         Left            =   5160
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdCitasAbogado 
         Caption         =   "Citas"
         Height          =   735
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdCasos 
         Caption         =   "Listado de casos pendientes"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Registrar"
      Height          =   255
      Left            =   7000
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNewPass 
      Caption         =   "¿Olvido su contraseña?"
      Height          =   255
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Data datUsers 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "usuarios"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cerrar Sesion"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Entrar"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblPass 
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblUser 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UserLogin, Deudores, ChequeVencido, ChequeAPunto As Boolean

Private Sub cmdCalificacion_Click()
    If Permiso = "Gerente Adm." Then
        frmCalificacion.Show
        frmMain.Hide
        Exit Sub
    End If
    If Permiso = "Abogado" Then
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If datUsers.Recordset.Fields(5) = datAbogado.Recordset.Fields(2) Then
                If datAbogado.Recordset.Fields(7) = "Senior" Then
                    frmCalificacion.Show
                    frmMain.Hide
                    Exit Sub
                Else
                    MsgBox "Esta es una funcion exclusiva para Abogados Senior", vbExclamation
                    Exit Sub
                End If
            Else
               datAbogado.Recordset.MoveNext
            End If
        Loop
    Else
        MsgBox "Esta es una funcion exclusiva para Abogados Senior", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdCasos_Click()
    frmVerCasosAbogado.Show
    frmMain.Hide
End Sub

Private Sub cmdCheques_Click()
    frmVerCheques.Show
    frmMain.Hide
End Sub

Private Sub cmdCitasAbogado_Click()
    frmVerCitasAbogado.Show
    frmMain.Hide
End Sub

Private Sub cmdCobros_Click()
    frmCobros.Show
    frmMain.Hide
End Sub

Private Sub cmdDeudores_Click()
    frmDeudores.Show
    frmMain.Hide
End Sub

Private Sub cmdDptos_Click()
    frmDptos.Show
    frmMain.Hide
End Sub

Private Sub cmdGestion_Click()
    frmUserGestion.Show
    frmMain.Hide
End Sub

Private Sub cmdNewPass_Click()
    frmNewPass.Show
    frmMain.Hide
End Sub

Private Sub cmdNuevaCita_Click()
    frmNuevaCita.Show
    frmMain.Hide
End Sub

Private Sub cmdNuevoCaso_Click()
    frmMain.Hide
    frmNuevoCaso.Show
End Sub

Private Sub cmdNuevoCliente_Click()
    frmNuevoCliente.Show
    frmMain.Hide
End Sub

Private Sub cmdPagos_Click()
    frmVerPagos.Show
    frmMain.Hide
End Sub

Private Sub cmdPagosAbogados_Click()
    frmPagosAbogados.Show
    frmMain.Hide
End Sub

Private Sub cmdReg_Click()
    frmReg.Show
    frmMain.Hide
End Sub

Private Sub cmdVerAbogados_Click()
    frmVerAbogados.Show
    frmMain.Hide
End Sub

Private Sub cmdVerCientes_Click()
    frmVerClientes.Show
    frmMain.Hide
End Sub

Private Sub cmdVerCitas_Click()
    frmVerCitas.Show
    frmMain.Hide
End Sub


Private Sub Command1_Click()
    frmTitulos.Show
    frmMain.Hide
End Sub

Private Sub Command2_Click()
    frmNuevoTitulo.Show
    frmMain.Hide
End Sub

Private Sub Command3_Click()
    frmNuevoAbogado.Show
    frmMain.Hide
End Sub

Private Sub Command4_Click()
    frmNuevaJerarquia.Show
    frmMain.Hide
End Sub

Private Sub Command5_Click()
    frmVerHistorial.Show
    frmMain.Hide
End Sub

Private Sub Form_Activate()
    txtUser = ""
    txtPass = ""
    If datPago.Recordset.EOF Then
        Exit Sub
    End If
    datPago.Recordset.MoveFirst
    Do While datPago.Recordset.EOF = False
        If DateDiff("d", Now, CDate(datPago.Recordset.Fields(1))) <= -30 And datPago.Recordset.Fields(5) = False Then
            datPago.Recordset.Edit
            datPago.Recordset.Fields(6) = True
            datPago.Recordset.Update
            Deudores = True
        End If
        datPago.Recordset.MoveNext
    Loop
    If datCheque.Recordset.EOF Then
        Exit Sub
    End If
    datCheque.Recordset.MoveFirst
    Do While datCheque.Recordset.EOF = False
        If DateDiff("d", Now, CDate(datCheque.Recordset.Fields(2))) <= -20 And DateDiff("d", Now, CDate(datCheque.Recordset.Fields(2))) >= -30 And datCheque.Recordset.Fields(5) = False Then
            ChequeAPunto = True
        End If
        If DateDiff("d", Now, CDate(datCheque.Recordset.Fields(2))) <= -30 And datCheque.Recordset.Fields(5) = False Then
            datCheque.Recordset.Edit
            datCheque.Recordset.Fields(6) = True
            datCheque.Recordset.Update
            ChequeVencido = True
        End If
        datCheque.Recordset.MoveNext
    Loop
End Sub

Private Sub cmdClose_Click()
    If UserLogin = True Then
        UserLogin = False
        cmdClose.Visible = False
        cmdReg.Visible = True
        cmdNewPass.Visible = True
        cmdLogin.Visible = True
        txtPass.Visible = True
        txtUser.Visible = True
        txtPass = ""
        txtUser = ""
        lblUser.Caption = "Usuario"
        lblPass.Caption = "Contraseña"
        frmContaduria.Enabled = False
        frmRecepcion.Enabled = False
        frmAbogado.Enabled = False
        frmCaja.Enabled = False
        frmGerente.Enabled = False
        nroDNI = 0
        If cmdGestion.Visible Then
            cmdGestion.Visible = False
        End If
        Exit Sub
    Else
        MsgBox "No hay ningun usuario logeado", vbCritical
        Exit Sub
    End If
End Sub


Private Sub cmdLogin_Click()
    If UserLogin = False Then
        If txtUser = "" Then
            MsgBox "No se introdujo nombre de usuario", vbCritical
            Exit Sub
        Else
            If txtPass = "" Then
                MsgBox "No se introdujo contraseña", vbCritical
                Exit Sub
            Else
                datUsers.Recordset.MoveFirst
                Do While datUsers.Recordset.EOF = False
                    If txtUser = datUsers.Recordset.Fields(0) Then
                        If txtPass = datUsers.Recordset.Fields(1) Then
                           MsgBox "Logueo Exitoso", vbInformation
                           UserLogin = True
                           cmdClose.Visible = True
                           cmdReg.Visible = False
                           cmdNewPass.Visible = False
                           cmdLogin.Visible = False
                           lblUser.Caption = "Usuario: " & datUsers.Recordset.Fields(0)
                           lblPass.Caption = "Permiso: " & datUsers.Recordset.Fields(2)
                           txtUser.Visible = False
                           txtPass.Visible = False
                           If datUsers.Recordset.Fields(2) = "Gerente Adm." Then
                                cmdGestion.Visible = True
                                frmRecepcion.Enabled = True
                                frmAbogado.Enabled = True
                                frmContaduria.Enabled = True
                                frmCaja.Enabled = True
                                frmGerente.Enabled = True
                                'Porcentaje = InputBox("Ingresar porcentaje de sueldo basico para abogados")
                           End If
                           If datUsers.Recordset.Fields(2) = "Recepcion" Then
                                frmRecepcion.Enabled = True
                                frmAbogado.Enabled = False
                                frmContaduria.Enabled = False
                                frmCaja.Enabled = False
                           End If
                           If datUsers.Recordset.Fields(2) = "Contaduria" Then
                                frmRecepcion.Enabled = False
                                frmAbogado.Enabled = False
                                frmContaduria.Enabled = True
                                frmCaja.Enabled = False
                                If Deudores = True Then
                                    MsgBox "Se han detectado nuevos deudores", vbExclamation
                                End If
                                If ChequeAPunto = True Then
                                    MsgBox "Hay cheques a punto de vencer", vbExclamation
                                End If
                                If ChequeVencido = True Then
                                    MsgBox "Hay cheques vencidos", vbExclamation
                                End If
                           End If
                           If datUsers.Recordset.Fields(2) = "Abogado" Then
                                frmRecepcion.Enabled = False
                                frmAbogado.Enabled = True
                                frmContaduria.Enabled = False
                                frmCaja.Enabled = False
                           End If
                           If datUsers.Recordset.Fields(2) = "Caja" Then
                                frmRecepcion.Enabled = False
                                frmAbogado.Enabled = False
                                frmContaduria.Enabled = False
                                frmCaja.Enabled = True
                           End If
                           nroDNI = datUsers.Recordset.Fields(5)
                           Permiso = datUsers.Recordset.Fields(2)
                           Exit Sub
                        Else
                           MsgBox "Contraseña Invalida", vbCritical
                           Exit Sub
                        End If
                    Else
                        datUsers.Recordset.MoveNext
                    End If
                Loop
                MsgBox "El usuario especificado no existe", vbCritical
            End If
        End If
    Else
        MsgBox "Ya hay un usuario logeado", vbCritical
    End If
End Sub
