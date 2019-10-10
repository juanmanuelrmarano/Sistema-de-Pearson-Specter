VERSION 5.00
Begin VB.Form frmVerAbogados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Nuestros Abogados"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   15375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Eliminar abogado"
      Height          =   1215
      Left            =   10560
      TabIndex        =   12
      Top             =   4320
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "DNI abogado"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Data datDpto 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "departamento"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datAbogado 
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
      RecordSource    =   "abogado"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   3960
      Left            =   5160
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   3960
      Left            =   6840
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List6 
      Height          =   3960
      Left            =   8520
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List7 
      Height          =   3960
      Left            =   10200
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List8 
      Height          =   3960
      Left            =   11880
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List9 
      Height          =   3960
      Left            =   13560
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Departamento"
      Height          =   255
      Left            =   13560
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmVerAbogados.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "frmVerAbogados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    If VerAbogadosCita = True Then
        Unload Me
        VerAbogadosCita = False
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
        MsgBox "Ingresar un DNI de abogado", vbExclamation
        Exit Sub
    Else
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If datAbogado.Recordset.Fields(2) = Text1 Then
                datAbogado.Recordset.Delete
                MsgBox "Abogado eliminado con exito", vbInformation
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datAbogado.Recordset.MoveNext
        Loop
        MsgBox "Ese abogado no existe", vbCritical
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
    List8.Clear
    List9.Clear
    If VerAbogadosCita = True Then
        Command1.Enabled = False
    End If
    datAbogado.Recordset.MoveFirst
    Do While datAbogado.Recordset.EOF = False
        List1.AddItem datAbogado.Recordset.Fields(2)
        List2.AddItem datAbogado.Recordset.Fields(0)
        List3.AddItem datAbogado.Recordset.Fields(1)
        List4.AddItem datAbogado.Recordset.Fields(3)
        List5.AddItem datAbogado.Recordset.Fields(4)
        List6.AddItem datAbogado.Recordset.Fields(5)
        List7.AddItem datAbogado.Recordset.Fields(7)
        List8.AddItem datAbogado.Recordset.Fields(10)
        datDpto.Recordset.MoveFirst
        Do While datDpto.Recordset.EOF = False
            If datAbogado.Recordset.Fields(11) = datDpto.Recordset.Fields(0) Then
                List9.AddItem datDpto.Recordset.Fields(1)
                Exit Do
            End If
            datDpto.Recordset.MoveNext
        Loop
        datAbogado.Recordset.MoveNext
    Loop
End Sub

