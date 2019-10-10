VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Registro"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datAbogado 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtDNI 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox cboDpto 
      Height          =   315
      ItemData        =   "frmReg.frx":0000
      Left            =   1680
      List            =   "frmReg.frx":0010
      TabIndex        =   12
      Text            =   "Seleccionar..."
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtRes 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox cboPreg 
      Height          =   315
      ItemData        =   "frmReg.frx":003A
      Left            =   1680
      List            =   "frmReg.frx":0050
      TabIndex        =   7
      Text            =   "Seleccionar..."
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Data datUsers 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "usuarios"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtContraRep 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtContra 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "DNI Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Puesto"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Respuesta secreta"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Pregunta secreta"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Repetir contraseña"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HayUsuario As Boolean

Private Sub cboDpto_Click()
    If cboDpto.Text = "Abogado" Then
        txtDNI.Enabled = True
    Else
        txtDNI.Enabled = False
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    HayUsuario = False
End Sub

Private Sub cmdAceptar_Click()
    HayUsuario = False
    If cboDpto.Text = "Abogado" Then
        If IsNumeric(txtDNI) = False Then
            MsgBox "El DNI debe ser un numero", vbExclamation
            Exit Sub
        End If
    End If
    If txtUser = "" Or txtContra = "" Or txtContraRep = "" Or txtRes = "" Or cboPreg.Text = "Seleccionar..." Then
        MsgBox "Ha dejado alguno de los campos vacio", vbExclamation
        Exit Sub
    Else
        datUsers.Recordset.MoveFirst
        Do While datUsers.Recordset.EOF = False
            If txtUser = datUsers.Recordset.Fields(0) Then
                MsgBox "Ese nombre de usuario ya existe", vbExclamation
                txtUser = ""
                HayUsuario = True
                Exit Do
            End If
            datUsers.Recordset.MoveNext
        Loop
        If HayUsuario = False Then
            If txtContra = txtContraRep Then
                If cboDpto = "Abogado" Then
                    datAbogado.Recordset.MoveFirst
                    Do While datAbogado.Recordset.EOF = False
                        If txtDNI = datAbogado.Recordset.Fields(2) Then
                            datUsers.Recordset.AddNew
                            datUsers.Recordset.Fields(0) = txtUser
                            datUsers.Recordset.Fields(1) = txtContra
                            datUsers.Recordset.Fields(2) = "Abogado"
                            Select Case cboPreg.Text
                            Case "¿Cual es el nombre de su primera mascota?"
                                datUsers.Recordset.Fields(3) = 1
                            Case "¿Cual es la marca de su primer auto?"
                                datUsers.Recordset.Fields(3) = 2
                            Case "¿Cual es el nombre de su madre?"
                                datUsers.Recordset.Fields(3) = 3
                            Case "¿Cual es el nombre de su padre?"
                                datUsers.Recordset.Fields(3) = 4
                            Case "¿Cual es el nombre de su hermano?"
                                datUsers.Recordset.Fields(3) = 5
                            Case "¿Cual es el nombre de su hermana?"
                                datUsers.Recordset.Fields(3) = 6
                            Case Else
                                cboPreg.Text = "Seleccionar..."
                                MsgBox "Pregunta secreta invalida", vbCritical
                                Exit Sub
                            End Select
                            datUsers.Recordset.Fields(4) = txtRes
                            datUsers.Recordset.Fields(5) = txtDNI
                            datUsers.Recordset.Update
                            MsgBox "Usuario registrado con exito", vbInformation
                            frmMain.Show
                            frmReg.Hide
                            Exit Sub
                        End If
                    datAbogado.Recordset.MoveNext
                    Loop
                    MsgBox "No hay abogados con ese DNI", vbCritical
                Else
                    datUsers.Recordset.AddNew
                    datUsers.Recordset.Fields(0) = txtUser
                    datUsers.Recordset.Fields(1) = txtContra
                    Select Case cboDpto.Text
                    Case "Contaduria"
                        datUsers.Recordset.Fields(2) = "Contaduria"
                    Case "Recepcion"
                        datUsers.Recordset.Fields(2) = "Recepcion"
                    Case "Caja"
                        datUsers.Recordset.Fields(2) = "Caja"
                    Case Else
                        cboDpto.Text = "Seleccionar..."
                        MsgBox "Departamento Invalido.", vbCritical
                        Exit Sub
                    End Select
                    Select Case cboPreg.Text
                    Case "¿Cual es el nombre de su primera mascota?"
                        datUsers.Recordset.Fields(3) = 1
                    Case "¿Cual es la marca de su primer auto?"
                        datUsers.Recordset.Fields(3) = 2
                    Case "¿Cual es el nombre de su madre?"
                        datUsers.Recordset.Fields(3) = 3
                    Case "¿Cual es el nombre de su padre?"
                        datUsers.Recordset.Fields(3) = 4
                    Case "¿Cual es el nombre de su hermano?"
                        datUsers.Recordset.Fields(3) = 5
                    Case "¿Cual es el nombre de su hermana?"
                        datUsers.Recordset.Fields(3) = 6
                    Case Else
                        cboPreg.Text = "Seleccionar..."
                        MsgBox "Pregunta secreta invalida", vbCritical
                        Exit Sub
                    End Select
                    datUsers.Recordset.Fields(4) = txtRes
                    datUsers.Recordset.Fields(5) = 0
                    datUsers.Recordset.Update
                    MsgBox "Usuario registrado con exito", vbInformation
                    frmMain.Show
                    frmReg.Hide
                    Exit Sub
                End If
            Else
                MsgBox "Las contraseñas no coinciden", vbExclamation
            End If
        End If
    End If
End Sub
