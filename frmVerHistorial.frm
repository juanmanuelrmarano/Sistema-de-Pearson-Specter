VERSION 5.00
Begin VB.Form frmVerHistorial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Historial de abogados"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Data datAbogado 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datHistorial 
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
      RecordSource    =   "historial_abogado"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmVerHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    datAbogado.Recordset.MoveFirst
    datHistorial.Recordset.MoveFirst
    Do While datAbogado.Recordset.EOF = False
        List1.AddItem "----Abogado " & datAbogado.Recordset.Fields(0) & " " & datAbogado.Recordset.Fields(1) & "----"
        datHistorial.Recordset.MoveFirst
        Do While datHistorial.Recordset.EOF = False
            If datAbogado.Recordset.Fields(2) = datHistorial.Recordset.Fields(0) Then
                List1.AddItem "Codigo Cita --> " & datHistorial.Recordset.Fields(1)
            End If
            datHistorial.Recordset.MoveNext
        Loop
        datAbogado.Recordset.MoveNext
        List1.AddItem " "
    Loop
End Sub
