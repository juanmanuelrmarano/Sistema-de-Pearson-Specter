VERSION 5.00
Begin VB.Form frmCobros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Cobros"
   ClientHeight    =   585
   ClientLeft      =   5520
   ClientTop       =   2340
   ClientWidth     =   8025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   8025
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Facturas"
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Data datCheque 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cheque_cliente"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmCobrar 
      Caption         =   "Cobrar"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Data datDpto 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "departamento"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datAbogado 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datCita 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "citas"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datCliente 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cliente"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Generar Factura"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtFact 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Data datPago 
      Caption         =   "datAbogado"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "factura_cliente"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Monto"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Banco"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Informacion de cheque"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "$"
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
      TabIndex        =   9
      Top             =   3470
      Width           =   135
   End
   Begin VB.Label Label7 
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
      Left            =   6120
      TabIndex        =   8
      Top             =   3465
      Width           =   1695
   End
   Begin VB.Label Label6 
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
      Left            =   5640
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
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
      Left            =   5640
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      TabIndex        =   5
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label3 
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1950
      Width           =   4095
   End
   Begin VB.Label Label2 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1675
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   4785
      Left            =   120
      Picture         =   "frmCobros.frx":0000
      Top             =   600
      Visible         =   0   'False
      Width           =   7755
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Encontrado, Pagada As Boolean
Dim dniabogado, nrodpto, Cod As Integer

Private Sub cmCobrar_Click()
        Pagada = False
        datPago.Recordset.MoveFirst
        Do While datPago.Recordset.EOF = False
            If txtFact = datPago.Recordset.Fields(0) Then
                If datPago.Recordset.Fields(5) = True Then
                    Pagada = True
                End If
                Exit Do
            End If
            datPago.Recordset.MoveNext
        Loop
        If Pagada = True Then
            MsgBox "Esta factura ya fue cobrada!", vbCritical
            Exit Sub
        Else
            If datCheque.Recordset.EOF Then
                Cod = 0
            Else
            datCheque.Recordset.MoveLast
            Cod = datCheque.Recordset.Fields(0) + 1
            End If
            datCheque.Recordset.AddNew
            datCheque.Recordset.Fields(0) = Cod
            datCheque.Recordset.Fields(1) = txtFact
            datCheque.Recordset.Fields(2) = DateValue(Now)
            datCheque.Recordset.Fields(3) = Text1
            If Text2 < Label7 Then
                MsgBox "El monto a pagar es mayor", vbExclamation
                Exit Sub
            Else
                datCheque.Recordset.Fields(4) = Text2
            End If
            datCheque.Recordset.Update
            MsgBox "La factura Nro " & datPago.Recordset.Fields(0) & " ha sido cobrada por un monto de " & "$" & datPago.Recordset.Fields(3), vbInformation
            datPago.Recordset.Edit
            datPago.Recordset.Fields(5) = True
            datPago.Recordset.Fields(6) = False
            datPago.Recordset.Update
            Unload Me
            frmMain.Show
            Exit Sub
        End If
End Sub

Private Sub cmdAceptar_Click()
    Encontrado = False
    If IsNumeric(txtFact) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If txtFact = "" Then
        MsgBox "Debe ingresar un numero de factura", vbExclamation
    Else
        frmCobros.Height = 7515
        cmdVolver.Top = 6600
        cmdVolver.Left = 120
        Image1.Visible = True
        datPago.Recordset.MoveFirst
        Do While datPago.Recordset.EOF = False
            If txtFact = datPago.Recordset.Fields(0) Then
                Encontrado = True
                Exit Do
            End If
            datPago.Recordset.MoveNext
        Loop
        If Encontrado = False Then
            MsgBox "Factura invalida", vbExclamation
            frmMain.Show
            Unload Me
            Exit Sub
        Else
            Label4 = txtFact
            Label3 = DateValue(Now) & " " & TimeValue(Now)
            Label7 = datPago.Recordset.Fields(3)
            datCliente.Recordset.MoveFirst
            Do While datCliente.Recordset.EOF = False
                If datPago.Recordset.Fields(2) = datCliente.Recordset.Fields(2) Then
                    Label2 = datCliente.Recordset.Fields(0) & " " & datCliente.Recordset.Fields(1)
                    Exit Do
                End If
                datCliente.Recordset.MoveNext
            Loop
            datCita.Recordset.MoveFirst
            Do While datCita.Recordset.EOF = False
                If datPago.Recordset.Fields(4) = datCita.Recordset.Fields(0) Then
                    dniabogado = datCita.Recordset.Fields(2)
                    Exit Do
                End If
                datCita.Recordset.MoveNext
            Loop
            datAbogado.Recordset.MoveFirst
            Do While datAbogado.Recordset.EOF = False
                If datAbogado.Recordset.Fields(2) = dniabogado Then
                    Label6 = "Dr. " & datAbogado.Recordset.Fields(0) & " " & datAbogado.Recordset.Fields(1)
                    nrodpto = datAbogado.Recordset.Fields("departamento")
                    Exit Do
                End If
                datAbogado.Recordset.MoveNext
            Loop
            datDpto.Recordset.MoveFirst
            Do While datDpto.Recordset.EOF = False
                If datDpto.Recordset.Fields(0) = nrodpto Then
                    Label5 = "Dpto. " & datDpto.Recordset.Fields(1)
                    Exit Do
                End If
                datDpto.Recordset.MoveNext
            Loop
        End If
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload frmVerPagos
    Unload Me
End Sub

Private Sub Command1_Click()
    VerFacturas = True
    frmVerPagos.Show
End Sub
