VERSION 5.00
Begin VB.Form frmNuevoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Nuevo Cliente"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datCliente 
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
      RecordSource    =   "cliente"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3960
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "@"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Celular"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Telefono Laboral"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Telefono"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "DNI"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If IsNumeric(Text3) Or IsNumeric(Text2) Or IsNumeric(Text4) Or IsNumeric(Text8) Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Text1) = False Or IsNumeric(Text5) = False Or IsNumeric(Text6) = False Or IsNumeric(Text7) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Then
        MsgBox "Ha dejado uno de los campos vacio", vbInformation
        Exit Sub
    Else
        datCliente.Recordset.MoveFirst
        Do While datCliente.Recordset.EOF = False
            If Text1 = datCliente.Recordset.Fields(2) Then
                MsgBox "Ese DNI ya existe", vbCritical
                Exit Sub
            End If
            datCliente.Recordset.MoveNext
        Loop
        datCliente.Recordset.AddNew
        datCliente.Recordset.Fields(0) = Text2
        datCliente.Recordset.Fields(1) = Text3
        datCliente.Recordset.Fields(2) = Text1
        datCliente.Recordset.Fields(3) = Text4 & "@" & Text8
        datCliente.Recordset.Fields(4) = Text5
        datCliente.Recordset.Fields(5) = Text6
        datCliente.Recordset.Fields(6) = Text7
        datCliente.Recordset.Update
        MsgBox "Nuevo cliente registrado", vbInformation
        frmMain.Show
        Unload Me
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub
