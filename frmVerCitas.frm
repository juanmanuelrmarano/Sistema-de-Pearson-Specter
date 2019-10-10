VERSION 5.00
Begin VB.Form frmVerCitas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Citas"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Eliminar cita"
      Height          =   1215
      Left            =   4080
      TabIndex        =   6
      Top             =   3360
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "ID cita"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ListBox lstAbogado 
      Height          =   2985
      Left            =   6600
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Data datCitas 
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
      RecordSource    =   "citas"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lstCliente 
      Height          =   2985
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstFecha 
      Height          =   2985
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstNum 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"frmVerCitas.frx":0000
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmVerCitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVolver_Click()
    If VerCitasCita = True Then
        Unload Me
        VerCitasCita = False
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
        MsgBox "Ingresar un ID de cita", vbExclamation
        Exit Sub
    Else
        datCitas.Recordset.MoveFirst
        Do While datCitas.Recordset.EOF = False
            If datCitas.Recordset.Fields(0) = Text1 Then
                datCitas.Recordset.Delete
                MsgBox "Cita eliminada con exito", vbInformation
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datCitas.Recordset.MoveNext
        Loop
        MsgBox "Esa cita no existe", vbCritical
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    lstNum.Clear
    lstFecha.Clear
    lstCliente.Clear
    lstAbogado.Clear
    If VerCitasCita = True Then
        Command1.Enabled = False
    End If
    datCitas.Recordset.MoveFirst
    Do While datCitas.Recordset.EOF = False
        lstNum.AddItem datCitas.Recordset.Fields(0)
        lstFecha.AddItem datCitas.Recordset.Fields(1) & " a las " & datCitas.Recordset.Fields(3)
        lstCliente.AddItem datCitas.Recordset.Fields(4)
        lstAbogado.AddItem datCitas.Recordset.Fields(2)
        datCitas.Recordset.MoveNext
    Loop
End Sub
