VERSION 5.00
Begin VB.Form frmNuevoCaso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Iniciar nuevo caso"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datDpto 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "departamento"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboFuero 
      Height          =   315
      ItemData        =   "frmNuevoCaso.frx":0000
      Left            =   1680
      List            =   "frmNuevoCaso.frx":0002
      TabIndex        =   28
      Text            =   "Seleccionar..."
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1680
      TabIndex        =   26
      Top             =   840
      Width           =   3495
   End
   Begin VB.Data datCita 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "citas"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datCaso 
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
      RecordSource    =   "casos"
      Top             =   4080
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
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.ComboBox cboConc 
      Height          =   315
      ItemData        =   "frmNuevoCaso.frx":0004
      Left            =   1680
      List            =   "frmNuevoCaso.frx":000E
      TabIndex        =   4
      Text            =   "Seleccionar..."
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Data datCliente 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cliente"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "ID cita"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Fuero"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Juez"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Juzgado"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de actor"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "DNI abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "DNI cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Conciliacion"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Monto de demanda"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de inicio"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "AAAA"
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   " MM"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   " DD"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "frmNuevoCaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cod As Integer
Dim CitaVal, ClienVal, AboVal As Boolean

Private Sub cmdAceptar_Click()
      CitaVal = False
      ClienVal = False
      AboVal = False
    If Text11 < 0 Or Text11 > 31 Then
        MsgBox "Dia invalido", vbExclamation
        Exit Sub
    End If
    If Text12 < 0 Or Text12 > 12 Then
        MsgBox "Mes invalido", vbExclamation
        Exit Sub
    End If
    If Text8 < 1000 Or Text8 > 9999 Then
        MsgBox "Año invalido", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Text4) Or IsNumeric(Text3) Or IsNumeric(Text6) Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Text1) = False Or IsNumeric(Text2) = False Or IsNumeric(Text10) = False Or IsNumeric(Text11) = False Or IsNumeric(Text12) = False Or IsNumeric(Text8) = False Or IsNumeric(Text9) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
      If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Then
        MsgBox "Ha dejado uno de los campos vacio", vbExclamation
        Exit Sub
      Else
        datCaso.Recordset.MoveLast
        Cod = datCaso.Recordset.Fields(0) + 1
        datCaso.Recordset.AddNew
        datCaso.Recordset.Fields(0) = Cod
        datCita.Recordset.MoveFirst
        Do While datCita.Recordset.EOF = False
            If Text10 = datCita.Recordset.Fields(0) Then
                datCaso.Recordset.Fields(1) = Text10
                CitaVal = True
                Exit Do
            End If
            datCita.Recordset.MoveNext
        Loop
        datCliente.Recordset.MoveFirst
        Do While datCliente.Recordset.EOF = False
            If Text1 = datCliente.Recordset.Fields(2) Then
                datCaso.Recordset.Fields(2) = Text1
                ClienVal = True
                Exit Do
            End If
            datCliente.Recordset.MoveNext
        Loop
        datCaso.Recordset.Fields(3) = Text3
        If IsDate(Text11 & "/" & Text12 & "/" & Text8) Then
            datCaso.Recordset.Fields(4) = Text11 & "/" & Text12 & "/" & Text8
        Else
            MsgBox "No es una fecha valida", vbExclamation
            Exit Sub
        End If
        datCaso.Recordset.Fields(5) = Text3
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If Text2 = datAbogado.Recordset.Fields(2) Then
                datCaso.Recordset.Fields(6) = Text2
                AboVal = True
                Exit Do
            End If
            datAbogado.Recordset.MoveNext
        Loop
        If ClienVal = False Then
            MsgBox "Ese cliente no existe", vbExclamation
            Exit Sub
        ElseIf AboVal = False Then
            MsgBox "Ese abogado no existe", vbExclamation
            Exit Sub
        ElseIf CitaVal = False Then
            MsgBox "Esa cita no existe", vbExclamation
            Exit Sub
        End If
        datCaso.Recordset.Fields(7) = Text5
        datCaso.Recordset.Fields(8) = Text6
        Select Case cboFuero.Text
        Case "Penal"
            datCaso.Recordset.Fields(9) = "Penal"
        Case "Civil"
            datCaso.Recordset.Fields(9) = "Civil"
        Case "Constitucional"
            datCaso.Recordset.Fields(9) = "Constitucional"
        Case "Administrativo"
            datCaso.Recordset.Fields(9) = "Administrativo"
        Case "Procesal"
            datCaso.Recordset.Fields(9) = "Procesal"
        Case "Financiero"
            datCaso.Recordset.Fields(9) = "Financiero"
        Case "Tributario"
            datCaso.Recordset.Fields(9) = "Tributario"
        Case Else
            cboFuero.Text = "Seleccionar..."
            MsgBox "Fuero invalido", vbCritical
            Exit Sub
        End Select
        Select Case cboConc.Text
        Case "Si"
            datCaso.Recordset.Fields(11) = True
        Case "No"
            datCaso.Recordset.Fields(11) = False
        Case Else
            cboConc.Text = "Seleccionar..."
            MsgBox "Conciliacion invalida", vbCritical
            Exit Sub
        End Select
        datCaso.Recordset.Fields(10) = Text9
        MsgBox "Caso iniciado", vbInformation
        datCaso.Recordset.Update
        frmMain.Show
        Unload Me
      End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    datDpto.Recordset.MoveFirst
    Do While datDpto.Recordset.EOF = False
        cboFuero.AddItem datDpto.Recordset.Fields(1)
        datDpto.Recordset.MoveNext
    Loop
End Sub
    
