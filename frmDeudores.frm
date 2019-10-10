VERSION 5.00
Begin VB.Form frmDeudores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Deudores"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10350
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   4350
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   4350
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   4350
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List6 
      Height          =   4350
      Left            =   8520
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Data datPago 
      Caption         =   "datAbogado"
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
      RecordSource    =   "factura_cliente"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDeudores.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmDeudores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    datPago.Recordset.MoveFirst
    Do While datPago.Recordset.EOF = False
        If datPago.Recordset.Fields(6) = True Then
            List1.AddItem datPago.Recordset.Fields(0)
            List2.AddItem datPago.Recordset.Fields(1)
            List3.AddItem datPago.Recordset.Fields(2)
            List4.AddItem datPago.Recordset.Fields(3)
            List5.AddItem datPago.Recordset.Fields(4)
            If datPago.Recordset.Fields(5) = False Then
                List6.AddItem "No"
            Else
                List6.AddItem "Si"
            End If
        End If
        datPago.Recordset.MoveNext
    Loop
End Sub
