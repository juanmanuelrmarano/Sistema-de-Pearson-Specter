VERSION 5.00
Begin VB.Form frmVerCitasAbogado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Citas de abogado"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datCitas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "citas"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstCliente 
      Columns         =   2
      DataField       =   "dni_cliente"
      DataSource      =   "datCitas"
      Height          =   2985
      ItemData        =   "frmVerCitasAbogado.frx":0000
      Left            =   4440
      List            =   "frmVerCitasAbogado.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstFecha 
      Height          =   2985
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstNum 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Numero Cita                             Fecha                                      DNI Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmVerCitasAbogado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    datCitas.Recordset.MoveFirst
    Do While datCitas.Recordset.EOF = False
        If nroDNI = datCitas.Recordset.Fields(2) Then
            lstNum.AddItem datCitas.Recordset.Fields(0)
            lstFecha.AddItem datCitas.Recordset.Fields(1) & " a las " & datCitas.Recordset.Fields(3)
            lstCliente.AddItem datCitas.Recordset.Fields(4)
        End If
        datCitas.Recordset.MoveNext
    Loop
End Sub
