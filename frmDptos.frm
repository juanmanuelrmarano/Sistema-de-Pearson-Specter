VERSION 5.00
Begin VB.Form frmDptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Departamentos"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Height          =   1620
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.Data datDpto 
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
      RecordSource    =   "departamento"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cboDpto 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Seleccionar..."
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Nuevo porcentaje"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Departamento"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If IsNumeric(Text1) = False Then
        MsgBox "datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text1 = "" Or cboDpto.Text = "Seleccionar..." Then
        MsgBox "Ha dejado uno de los campos vacio", vbExclamation
        Exit Sub
    Else
        datDpto.Recordset.MoveFirst
        Do While datDpto.Recordset.EOF = False
            If cboDpto.Text = datDpto.Recordset.Fields(1) Then
                datDpto.Recordset.Edit
                datDpto.Recordset.Fields(3) = Text1
                datDpto.Recordset.Update
                MsgBox "Cambio registrado", vbInformation
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datDpto.Recordset.MoveNext
        Loop
        MsgBox "Ese departamento no existe", vbExclamation
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    datDpto.Recordset.MoveFirst
    Do While datDpto.Recordset.EOF = False
        cboDpto.AddItem datDpto.Recordset.Fields(1)
        datDpto.Recordset.MoveNext
    Loop
    datDpto.Recordset.MoveFirst
    Do While datDpto.Recordset.EOF = False
        List3.AddItem datDpto.Recordset.Fields(1)
        List2.AddItem datDpto.Recordset.Fields(3)
        datDpto.Recordset.MoveNext
    Loop
End Sub
