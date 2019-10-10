VERSION 5.00
Begin VB.Form frmNewPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Cambio de contraseña"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtPreg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdAceptar3 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtNuevaPassRep 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox txtNuevaPass 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton cmdAceptar2 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtRes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Data datUsers 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "usuarios"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblPreg 
      Caption         =   "Pregunta secreta"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Repetir nueva contraseña"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Ingrese una nueva contraseña"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Respuesta"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblUser 
      Caption         =   "Nombre de usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    If txtUser.Text = "" Then
        MsgBox "Ingrese un nombre de usuario", vbExclamation
    Else
        datUsers.Recordset.MoveFirst
        Do While datUsers.Recordset.EOF = False
            If txtUser = datUsers.Recordset.Fields(0) Then
                Select Case datUsers.Recordset.Fields(3)
                Case 1
                    txtPreg = "¿Cual es el nombre de su primera mascota?"
                Case 2
                    txtPreg = "¿Cual es la marca de su primer auto?"
                Case 3
                    txtPreg = "¿Cual es el nombre de su madre?"
                Case 4
                    txtPreg = "¿Cual es el nombre de su padre?"
                Case 5
                    txtPreg = "¿Cual es el nombre de su hermano?"
                Case 6
                    txtPreg = "¿Cual es el nombre de su hermana?"
                End Select
                txtRes.Enabled = True
                cmdAceptar2.Enabled = True
                Exit Sub
            Else
                datUsers.Recordset.MoveNext
            End If
        Loop
        MsgBox "No existe ningun usuario con ese nombre", vbExclamation
        txtUser = ""
        Exit Sub
    End If
End Sub

Private Sub cmdAceptar2_Click()
    If txtRes.Text = "" Then
        MsgBox "Ingrese una respuesta", vbExclamation
    Else
        If txtRes = datUsers.Recordset.Fields(4) Then
            txtNuevaPass.Enabled = True
            txtNuevaPassRep.Enabled = True
            cmdAceptar3.Enabled = True
            Exit Sub
        Else
            MsgBox "La respuesta es incorrecta", vbExclamation
            txtRes = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdAceptar3_Click()
    If txtNuevaPass.Text = "" Then
        MsgBox "Ingrese una nueva contraseña", vbExclamation
    Else
        If txtNuevaPass.Text = txtNuevaPassRep.Text Then
            datUsers.Recordset.Edit
            datUsers.Recordset.Fields(1) = txtNuevaPass.Text
            datUsers.Recordset.Update
            MsgBox "Nueva contraseña registrada", vbInformation
            cmdAceptar2.Enabled = False
            cmdAceptar3.Enabled = False
            txtUser = ""
            txtPreg = ""
            txtRes = ""
            txtNuevaPass = ""
            txtNuevaPassRep = ""
            frmNewPass.Hide
            frmMain.Show
            Exit Sub
        Else
            MsgBox "Las contraseñas no coinciden", vbExclamation
        End If
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub
