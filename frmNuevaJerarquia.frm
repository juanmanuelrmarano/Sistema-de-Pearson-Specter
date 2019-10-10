VERSION 5.00
Begin VB.Form frmNuevaJerarquia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Cambiar jerarquia"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboJer 
      Height          =   315
      ItemData        =   "frmNuevaJerarquia.frx":0000
      Left            =   1560
      List            =   "frmNuevaJerarquia.frx":000A
      TabIndex        =   10
      Text            =   "Seleccionar..."
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtApellido 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Data datAbogado 
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
      RecordSource    =   "abogado"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nueva Jerarquia"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmNuevaJerarquia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If IsNumeric(txtNombre) Or IsNumeric(txtApellido) Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If txtNombre = "" Or txtApellido = "" Or cboJer = "Seleccionar..." Then
        MsgBox "Ha dejado uno de los campos vacio", vbExclamation
        Exit Sub
    Else
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If txtNombre = datAbogado.Recordset.Fields(0) And txtApellido = datAbogado.Recordset.Fields(1) Then
                datAbogado.Recordset.Edit
                Select Case cboJer.Text
                Case "Junior"
                    datAbogado.Recordset.Fields("puesto") = "Junior"
                Case "Senior"
                    datAbogado.Recordset.Fields("puesto") = "Senior"
                Case Else
                    cboPreg.Text = "Seleccionar..."
                    MsgBox "Puesto invalido", vbCritical
                    Exit Sub
                End Select
                datAbogado.Recordset.Update
                MsgBox "Nueva jerarquia registrada", vbInformation
                Unload Me
                frmMain.Show
                Exit Sub
            Else
                datAbogado.Recordset.MoveNext
            End If
        Loop
        MsgBox "El nombre o el apellido no existen", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    datAbogado.Recordset.MoveFirst
    Do While datAbogado.Recordset.EOF = False
        List1.AddItem datAbogado.Recordset.Fields(0)
        List2.AddItem datAbogado.Recordset.Fields(1)
        List3.AddItem datAbogado.Recordset.Fields("puesto")
        datAbogado.Recordset.MoveNext
    Loop
End Sub
