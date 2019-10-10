VERSION 5.00
Begin VB.Form frmUserGestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Gestion de Usuarios"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datAbogado 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data datUsers 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "usuarios"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Frame fraEliminate 
      Caption         =   "Eliminar Usuarios"
      Height          =   3375
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton cmdEliminate 
         Caption         =   "Eliminar Usuario"
         Height          =   855
         Left            =   4800
         TabIndex        =   21
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txtUserEliminate 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   1440
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre de usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame fraUsers 
      Caption         =   "Ver usuarios"
      Height          =   3375
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   9015
      Begin VB.ListBox lstUsers2 
         Height          =   2790
         Left            =   4560
         TabIndex        =   28
         Top             =   480
         Width           =   4335
      End
      Begin VB.ListBox lstUsers 
         Height          =   2790
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre                                                                                      Cargo"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame fraCharge 
      Caption         =   "Cambiar cargos"
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton cmdChargeAccept 
         Caption         =   "Aceptar"
         Height          =   735
         Left            =   7200
         TabIndex        =   27
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cboCharge 
         Height          =   315
         ItemData        =   "frmUserGestion.frx":0000
         Left            =   2160
         List            =   "frmUserGestion.frx":0010
         TabIndex        =   26
         Text            =   "Seleccionar..."
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox txtUserChange 
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label7 
         Caption         =   "Nueva seccion"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre de Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdUserEliminate 
      Caption         =   "Eliminar usuarios"
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame fraUserModify 
      Caption         =   "Modificacion de Usuarios"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9015
      Begin VB.CommandButton cmdAccept2 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   7440
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtNewPass 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   2400
         Width           =   5055
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   2040
         Width           =   5055
      End
      Begin VB.CommandButton cmdUserAccept 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtOldUser 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txtNewUser 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label10 
         Caption         =   "Cambiar nombre de usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Cambiar contraseña de usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Nueva contraseña"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre de usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre nuevo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre actual"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "Ver los usuarios"
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdRankChange 
      Caption         =   "Asignacion de seccion"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdUserModify 
      Caption         =   "Modificacion de usuarios"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
End
Attribute VB_Name = "frmUserGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dniabogado As Integer


Private Sub cmdAccept2_Click()
    If txtUser = "" Or txtNewPass = "" Then
        MsgBox "Dejo alguno de los campos vacio", vbExclamation
        Exit Sub
    Else
        datUsers.Recordset.MoveFirst
        Do While datUsers.Recordset.EOF = False
            If datUsers.Recordset.Fields(0) = txtUser Then
                datUsers.Recordset.Edit
                datUsers.Recordset.Fields(1) = txtNewPass
                MsgBox "Contraseña cambiada", vbInformation
                datUsers.Recordset.Update
                Exit Sub
            End If
            datUsers.Recordset.MoveNext
        Loop
    End If
    
End Sub

Private Sub cmdChargeAccept_Click()
    If txtUserChange = "" Then
        MsgBox "Debe ingresar un nombre de usuario"
        Exit Sub
    Else
        datUsers.Recordset.MoveFirst
        Do While datUsers.Recordset.EOF = False
            If datUsers.Recordset.Fields(0) = txtUserChange Then
                datUsers.Recordset.Edit
                Select Case cboCharge.Text
                    Case "Recepcion"
                        datUsers.Recordset.Fields(2) = "Recepcion"
                    Case "Abogado"
                        dniabogado = InputBox("Ingrese el dni del abogado al que se va a vincular")
                        datAbogado.Recordset.MoveFirst
                        Do While datAbogado.Recordset.EOF = False
                            If datAbogado.Recordset.Fields(2) = dniabogado Then
                                datUsers.Recordset.Fields(5) = dniabogado
                                datUsers.Recordset.Fields(2) = "Abogado"
                                Exit Sub
                            Else
                                datAbogado.Recordset.MoveNext
                            End If
                        Loop
                    Case "Contaduria"
                        datUsers.Recordset.Fields(2) = "Contaduria"
                    Case "Caja"
                        datUsers.Recordset.Fields(2) = "Caja"
                    Case Else
                        cboCharge.Text = "Seleccionar..."
                        MsgBox "Cargo invalido", vbCritical
                    Exit Sub
                End Select
                datUsers.Recordset.Update
            End If
            datUsers.Recordset.MoveNext
        Loop
        MsgBox "No hay usuarios con ese nombre", vbCritical
    End If
End Sub

Private Sub cmdEliminate_Click()
    If txtUserEliminate = "" Then
        MsgBox "Ingresar un nombre de usuario", vbExclamation
        Exit Sub
    Else
        datUsers.Recordset.MoveFirst
        Do While datUsers.Recordset.EOF = False
            If datUsers.Recordset.Fields(0) = txtUserEliminate Then
                datUsers.Recordset.Delete
                MsgBox "Usuario eliminado con exito", vbInformation
                Exit Sub
            End If
            datUsers.Recordset.MoveNext
        Loop
        MsgBox "Ese usuario no existe", vbCritical
        Exit Sub
    End If
End Sub

Private Sub cmdUserAccept_Click()
    If txtOldUser = txtNewUser Then
        MsgBox "No puede poner el mismo nombre que antes", vbCritical
    End If
    If txtOldUser = "" Or txtNewUser = "" Then
        MsgBox "Ha dejado alguno de los campos vacio", vbCritical
    Else
        datUsers.Recordset.MoveFirst
        Do While datUsers.Recordset.EOF = False
            If txtOldUser = datUsers.Recordset.Fields(0) Then
                datUsers.Recordset.Edit
                datUsers.Recordset.Fields(0) = txtNewUser
                datUsers.Recordset.Update
                MsgBox "Nombre cambiado con exito", vbInformation
                Exit Sub
            End If
            datUsers.Recordset.MoveNext
        Loop
        MsgBox "Ese nombre de usuario no existe", vbCritical
    End If
End Sub

Private Sub cmdUserEliminate_Click()
    fraUserModify.Visible = False
    fraEliminate.Visible = True
    fraUsers.Visible = False
    fraCharge.Visible = False
End Sub

Private Sub cmdUserModify_Click()
    fraUserModify.Visible = True
    fraEliminate.Visible = False
    fraUsers.Visible = False
    fraCharge.Visible = False
End Sub

Private Sub cmdUsers_Click()
    fraUserModify.Visible = False
    fraEliminate.Visible = False
    fraUsers.Visible = True
    fraCharge.Visible = False
    lstUsers.Clear
    lstUsers2.Clear
    datUsers.Recordset.MoveFirst
    Do While datUsers.Recordset.EOF = False
        lstUsers.AddItem datUsers.Recordset.Fields(0)
        lstUsers2.AddItem datUsers.Recordset.Fields(2)
        datUsers.Recordset.MoveNext
    Loop
End Sub

Private Sub cmdRankChange_Click()
    fraUserModify.Visible = False
    fraEliminate.Visible = False
    fraUsers.Visible = False
    fraCharge.Visible = True
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub
