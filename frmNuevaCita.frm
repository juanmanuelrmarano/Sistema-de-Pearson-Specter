VERSION 5.00
Begin VB.Form frmNuevaCita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Nueva cita"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Ver Citas"
      Height          =   1335
      Left            =   7200
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver abogados"
      Height          =   615
      Left            =   5160
      TabIndex        =   26
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Clientes"
      Height          =   615
      Left            =   5160
      TabIndex        =   25
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   15
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Data datHist 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "historial_abogado"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datFactura 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "factura_cliente"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datCita 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "citas"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datCliente 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cliente"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtAbogado 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Data datAbogado 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   18
      Top             =   1750
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   5370
      Left            =   120
      Picture         =   "frmNuevaCita.frx":0000
      Top             =   1560
      Width           =   8715
   End
   Begin VB.Label Label5 
      Caption         =   " MM"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   " HH"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Hora"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "AAAA"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   " MM"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   " DD"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "DNI Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "DNI Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmNuevaCita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodCita, CodFact As Integer
Dim AboVal As Boolean
Dim ClienVal As Boolean
Dim con As String


Private Sub cmdAceptar_Click()
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
    If DateDiff("d", DateValue(Now), CDate(Text11 & "/" & Text12 & "/" & Text8)) < 0 Then
        MsgBox "La fecha ingresada es anterior a la actual", vbExclamation
        Exit Sub
    End If
    If Text3 < 0 Or Text3 > 24 Then
        MsgBox "Hora invalida", vbExclamation
        Exit Sub
    End If
    If Text1 < 0 Or Text1 > 59 Then
        MsgBox "Minuto invalido", vbExclamation
        Exit Sub
    End If
    If IsNumeric(txtAbogado) = False Or IsNumeric(txtCliente) = False Or IsNumeric(Text11) = False Or IsNumeric(Text12) = False Or IsNumeric(Text8) = False Or IsNumeric(Text3) = False Or IsNumeric(Text1) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If txtAbogado = "" Or txtCliente = "" Or Text11 = "" Or Text12 = "" Or Text8 = "" Or Text3 = "" Or Text1 = "" Then
        MsgBox "Ha dejado uno de los campos vacios", vbExclamation
        Exit Sub
    Else
        AboVal = False
        ClienVal = False
        datCita.Recordset.MoveFirst
        If datCita.Recordset.EOF Then
            CodCita = 0
        Else
            datCita.Recordset.MoveLast
            CodCita = datCita.Recordset.Fields(0)
            datCita.Recordset.MoveFirst
        End If
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If txtAbogado = datAbogado.Recordset.Fields(2) Then
                AboVal = True
                datCita.Recordset.MoveFirst
                Do While datCita.Recordset.EOF = False
                    If datAbogado.Recordset.Fields(2) = datCita.Recordset.Fields(2) Then
                        If Text11 = Day(datCita.Recordset.Fields(1)) And Text12 = Month(datCita.Recordset.Fields(1)) And Text8 = Year(datCita.Recordset.Fields(1)) Then
                           If DateDiff("h", Text3 & ":" & Text1, datCita.Recordset.Fields(3)) = 0 Then
                                MsgBox "Horario no disponible, tiene una cita a las " & datCita.Recordset.Fields(3), vbExclamation
                                Exit Sub
                           End If
                        End If
                    End If
                    datCita.Recordset.MoveNext
                Loop
                Exit Do
            End If
            datAbogado.Recordset.MoveNext
        Loop
        datCita.Recordset.AddNew
        datHist.Recordset.AddNew
        datCita.Recordset.Fields(0) = CodCita + 1
        datHist.Recordset.Fields(1) = CodCita + 1
        If datFactura.Recordset.EOF Then
            CodFact = 0
        Else
            datFactura.Recordset.MoveLast
            CodFact = datFactura.Recordset.Fields(0)
            datFactura.Recordset.MoveFirst
        End If
        datFactura.Recordset.AddNew
        datFactura.Recordset.Fields(0) = CodFact + 1
        datFactura.Recordset.Fields(4) = CodCita + 1
        If IsDate(Text11 & "/" & Text12 & "/" & Text8) Then
            datCita.Recordset.Fields(1) = Text11 & "/" & Text12 & "/" & Text8
        Else
            MsgBox "No es una fecha valida", vbExclamation
            Exit Sub
        End If
        If IsDate(Text3 & ":" & Text1) Then
            datCita.Recordset.Fields(3) = Text3 & ":" & Text1
        Else
            MsgBox "No es una hora valida", vbExclamation
            Exit Sub
        End If
        datFactura.Recordset.Fields(1) = DateValue(Now)
        datCliente.Recordset.MoveFirst
        Do While datCliente.Recordset.EOF = False
            If txtCliente = datCliente.Recordset.Fields(2) Then
                datCita.Recordset.Fields(4) = txtCliente
                datFactura.Recordset.Fields(2) = txtCliente
                Label8 = datCliente.Recordset.Fields(0) & " " & datCliente.Recordset.Fields(1)
                ClienVal = True
                Exit Do
            End If
            datCliente.Recordset.MoveNext
        Loop
        If AboVal = False Or ClienVal = False Then
            MsgBox "El abogado o el cliente ingresados no existen", vbInformation
            Exit Sub
        Else
            datCita.Recordset.Fields(2) = txtAbogado
            datHist.Recordset.Fields(0) = txtAbogado
            datFactura.Recordset.Fields(3) = datAbogado.Recordset.Fields(10)
        End If
        con = (MsgBox("Los datos de la cita son correctos?", vbYesNo))
        If con = "6" Then
            MsgBox "Nueva cita registrada", vbInformation
        ElseIf con = "7" Then
            Exit Sub
        End If
        Label10 = "Dr. " & datAbogado.Recordset.Fields(0) & " " & datAbogado.Recordset.Fields(1)
        Label9 = DateValue(Now) & " " & TimeValue(Now)
        Label6 = CodCita + 1
        Label12 = Text11 & "/" & Text12 & "/" & Text8
        Label16 = Text3 & ":" & Text1
        Label11 = "Nro Factura: " & CodFact + 1
        datCita.Recordset.Update
        datFactura.Recordset.Update
        datHist.Recordset.Update
        Text11 = ""
        Text12 = ""
        Text8 = ""
        Text3 = ""
        Text1 = ""
        txtCliente = ""
        txtAbogado = ""
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload frmVerCitas
    Unload frmVerAbogados
    Unload frmVerClientes
    Unload Me
End Sub

Private Sub Command1_Click()
    VerClientesCita = True
    frmVerClientes.Show
End Sub

Private Sub Command2_Click()
    VerAbogadosCita = True
    frmVerAbogados.Show
End Sub

Private Sub Command3_Click()
    VerCitasCita = True
    frmVerCitas.Show
End Sub

Private Sub Command4_Click()
    frmMain.Show
    Unload frmVerCitas
    Unload frmVerAbogados
    Unload frmVerClientes
    MsgBox "Cita concretada", vbInformation
    Unload Me
End Sub
