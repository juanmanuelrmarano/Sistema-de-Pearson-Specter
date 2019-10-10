VERSION 5.00
Begin VB.Form frmVerCasosAbogado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Casos del abogado"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   18750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Eliminar caso"
      Height          =   1215
      Left            =   13920
      TabIndex        =   14
      Top             =   4320
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "ID caso"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ListBox List11 
      Height          =   3960
      Left            =   16920
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List10 
      Height          =   3960
      Left            =   15240
      TabIndex        =   10
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List9 
      Height          =   3960
      Left            =   13560
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List8 
      Height          =   3960
      Left            =   11880
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List7 
      Height          =   3960
      Left            =   10200
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List6 
      Height          =   3960
      Left            =   8520
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   3960
      Left            =   6840
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   3960
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Data datCaso 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "casos"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "Monto Pedido                Conciliacion"
      Height          =   255
      Left            =   15240
      TabIndex        =   13
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   $"frmVerCasosAbogado.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "frmVerCasosAbogado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    If IsNumeric(Text1) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text1 = "" Then
        MsgBox "Ingresar un ID de caso", vbExclamation
        Exit Sub
    Else
        datCaso.Recordset.MoveFirst
        Do While datCaso.Recordset.EOF = False
            If datCaso.Recordset.Fields(0) = Text1 Then
                datCaso.Recordset.Delete
                MsgBox "Caso eliminado con exito", vbInformation
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datCaso.Recordset.MoveNext
        Loop
        MsgBox "Ese caso no existe", vbCritical
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    datCaso.Recordset.MoveFirst
    Do While datCaso.Recordset.EOF = False
        If nroDNI = datCaso.Recordset.Fields(6) Then
            List1.AddItem datCaso.Recordset.Fields(0)
            List2.AddItem datCaso.Recordset.Fields(1)
            List3.AddItem datCaso.Recordset.Fields(2)
            List4.AddItem datCaso.Recordset.Fields(3)
            List5.AddItem datCaso.Recordset.Fields(4)
            List6.AddItem datCaso.Recordset.Fields(5)
            List7.AddItem datCaso.Recordset.Fields(7)
            List8.AddItem datCaso.Recordset.Fields(8)
            List9.AddItem datCaso.Recordset.Fields(9)
            List10.AddItem datCaso.Recordset.Fields(10)
            If datCaso.Recordset.Fields(11) = True Then
                List11.AddItem "Si"
            ElseIf datCaso.Recordset.Fields(11) = False Then
                List11.AddItem "No"
            End If
        End If
        datCaso.Recordset.MoveNext
    Loop
End Sub
