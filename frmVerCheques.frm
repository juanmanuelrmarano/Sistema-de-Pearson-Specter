VERSION 5.00
Begin VB.Form frmVerCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Registro de cheques"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List7 
      Height          =   4350
      Left            =   10200
      TabIndex        =   12
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7800
      TabIndex        =   9
      Top             =   5160
      Width           =   2655
   End
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
      Top             =   5160
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
   Begin VB.Data datCheque 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cheque_cliente"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Nro Cheque"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Cobrar cheque"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmVerCheques.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "frmVerCheques"
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
        MsgBox "Debe ingresar el ID del cheque que se cobro", vbExclamation
        Exit Sub
    Else
        datCheque.Recordset.MoveFirst
        Do While datCheque.Recordset.EOF = False
            If datCheque.Recordset.Fields(0) = Text1 Then
                If datCheque.Recordset.Fields(6) = True Then
                    MsgBox "Ese cheque ha vencido", vbExclamation
                    Exit Sub
                ElseIf datCheque.Recordset.Fields(5) = True Then
                    MsgBox "Ese cheque ya se cobro", vbExclamation
                    Exit Sub
                Else
                    datCheque.Recordset.Edit
                    datCheque.Recordset.Fields(5) = True
                    datCheque.Recordset.Update
                    MsgBox "Cheque cobrado exitosamente", vbInformation
                    frmMain.Show
                    Unload Me
                    Exit Sub
                End If
            End If
            datCheque.Recordset.MoveNext
        Loop
        MsgBox "Ese cheque no existe", vbExclamation
    End If
End Sub

Private Sub Form_Activate()
    If datCheque.Recordset.EOF Then
        Exit Sub
    Else
        datCheque.Recordset.MoveFirst
    End If
    Do While datCheque.Recordset.EOF = False
        List1.AddItem datCheque.Recordset.Fields(0)
        List2.AddItem datCheque.Recordset.Fields(1)
        List3.AddItem datCheque.Recordset.Fields(2)
        List4.AddItem datCheque.Recordset.Fields(3)
        List5.AddItem datCheque.Recordset.Fields(4)
        If datCheque.Recordset.Fields(5) = False Then
            List7.AddItem "No"
        Else
            List7.AddItem "Si"
        End If
        If datCheque.Recordset.Fields(6) = False Then
            List6.AddItem "No"
        Else
            List6.AddItem "Si"
        End If
        datCheque.Recordset.MoveNext
    Loop
End Sub

