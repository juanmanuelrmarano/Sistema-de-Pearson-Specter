VERSION 5.00
Begin VB.Form frmVerClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Nuestros Clientes"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Eliminar cliente"
      Height          =   1215
      Left            =   7680
      TabIndex        =   9
      Top             =   3240
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "DNI cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ListBox List7 
      Height          =   2985
      Left            =   10680
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List6 
      Height          =   2985
      Left            =   9000
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   2985
      Left            =   7200
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox List4 
      Height          =   2985
      Left            =   5400
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox List3 
      Height          =   2985
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Data datCliente 
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
      RecordSource    =   "cliente"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmVerClientes.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "frmVerClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    If VerClientesCita = True Then
        Unload Me
        VerClientesCita = False
        Exit Sub
    End If
    frmMain.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    If IsNumeric(Text1) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text1 = "" Then
        MsgBox "Ingresar un DNI de cliente", vbExclamation
        Exit Sub
    Else
        datCliente.Recordset.MoveFirst
        Do While datCliente.Recordset.EOF = False
            If datCliente.Recordset.Fields(2) = Text1 Then
                datCliente.Recordset.Delete
                MsgBox "Cliente eliminado con exito", vbInformation
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datCliente.Recordset.MoveNext
        Loop
        MsgBox "Ese cliente no existe", vbCritical
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    List7.Clear
    If VerClientesCita = True Then
        Command1.Enabled = False
    End If
    datCliente.Recordset.MoveFirst
    Do While datCliente.Recordset.EOF = False
        List1.AddItem datCliente.Recordset.Fields(2)
        List2.AddItem datCliente.Recordset.Fields(0)
        List3.AddItem datCliente.Recordset.Fields(1)
        List4.AddItem datCliente.Recordset.Fields(4)
        List5.AddItem datCliente.Recordset.Fields(5)
        List6.AddItem datCliente.Recordset.Fields(6)
        List7.AddItem datCliente.Recordset.Fields(3)
        datCliente.Recordset.MoveNext
    Loop
End Sub
