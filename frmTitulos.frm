VERSION 5.00
Begin VB.Form frmTitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Titulos de nuestros abogados"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datTitulo 
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
      RecordSource    =   "titulos"
      Top             =   4200
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
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmTitulos"
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
    datTitulo.Recordset.MoveFirst
    Do While datAbogado.Recordset.EOF = False
        List1.AddItem "----Abogado " & datAbogado.Recordset.Fields(0) & " " & datAbogado.Recordset.Fields(1) & "----"
        datTitulo.Recordset.MoveFirst
        Do While datTitulo.Recordset.EOF = False
            If datAbogado.Recordset.Fields(2) = datTitulo.Recordset.Fields(1) Then
                List1.AddItem "Nombre de Titulo   " & datTitulo.Recordset.Fields(4)
                List1.AddItem "Año Recibido        " & datTitulo.Recordset.Fields(2)
                List1.AddItem "Ente emisor           " & datTitulo.Recordset.Fields(3)
                List1.AddItem "-----------------"
            End If
            datTitulo.Recordset.MoveNext
        Loop
        datAbogado.Recordset.MoveNext
        List1.AddItem " "
    Loop
End Sub

