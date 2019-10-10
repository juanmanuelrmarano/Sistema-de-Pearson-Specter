VERSION 5.00
Begin VB.Form frmPagosAbogados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Pagos a abogados"
   ClientHeight    =   4935
   ClientLeft      =   4020
   ClientTop       =   3840
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datCita 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "citas"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datDpto 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "departamento"
      Top             =   4440
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
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "abogado"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "Liquidaciones extra"
      Height          =   1215
      Left            =   7080
      TabIndex        =   7
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir bonificacion de Diciembre"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "DNI Abogado"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Liquidacion de sueldo basico"
      Height          =   1215
      Left            =   7080
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "DNI Abogado"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ListBox List4 
      Height          =   3960
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Data datChequeAbogado 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "cheque_abogado"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Cheque                   Emision                         Monto                           DNI Abogado"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmPagosAbogados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cod As Integer
Dim monto, montodpto As Double

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    If datChequeAbogado.Recordset.EOF Then
        Cod = 0
    Else
        datChequeAbogado.Recordset.MoveLast
        Cod = datChequeAbogado.Recordset.Fields(0) + 1
    End If
    If IsNumeric(Text7) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text7 = "" Then
        MsgBox "Ingrese un DNI", vbExclamation
        Exit Sub
    Else
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If datAbogado.Recordset.Fields(2) = Text7 Then
                datChequeAbogado.Recordset.AddNew
                datChequeAbogado.Recordset.Fields(0) = Cod
                datChequeAbogado.Recordset.Fields(1) = DateValue(Now)
                datDpto.Recordset.MoveFirst
                Do While datDpto.Recordset.EOF = False
                    If datAbogado.Recordset.Fields(11) = datDpto.Recordset.Fields(0) Then
                        datChequeAbogado.Recordset.Fields(2) = datDpto.Recordset.Fields(2)
                        Exit Do
                    End If
                    datDpto.Recordset.MoveNext
                Loop
                datChequeAbogado.Recordset.Fields(3) = datAbogado.Recordset.Fields(2)
                MsgBox "Cheque de salario basico generado", vbInformation
                datChequeAbogado.Recordset.Update
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datAbogado.Recordset.MoveNext
        Loop
        MsgBox "Ese abogado no existe", vbExclamation
    End If
End Sub

Private Sub Command2_Click()
    datChequeAbogado.Recordset.MoveLast
    Cod = datChequeAbogado.Recordset.Fields(0) + 1
    If IsNumeric(Text1) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text1 = "" Then
        MsgBox "Ingrese un DNI", vbExclamation
        Exit Sub
    Else
        monto = 0
        datAbogado.Recordset.MoveFirst
        Do While datAbogado.Recordset.EOF = False
            If datAbogado.Recordset.Fields(2) = Text1 Then
                datChequeAbogado.Recordset.AddNew
                datChequeAbogado.Recordset.Fields(0) = Cod
                datChequeAbogado.Recordset.Fields(1) = DateValue(Now)
                datChequeAbogado.Recordset.Fields(3) = datAbogado.Recordset.Fields(2)
                datDpto.Recordset.MoveFirst
                If Check1.Value = 1 Then
                    monto = monto + 2000
                End If
                Do While datDpto.Recordset.EOF = False
                    If datAbogado.Recordset.Fields(11) = datDpto.Recordset.Fields(0) Then
                        montodpto = datDpto.Recordset.Fields(2)
                        Exit Do
                    End If
                    datDpto.Recordset.MoveNext
                Loop
                monto = monto + (((datAbogado.Recordset.Fields(12) * 10) * montodpto) / 100)
                datCita.Recordset.MoveFirst
                Do While datCita.Recordset.EOF = False
                    If Month(Now) = Month(datCita.Recordset.Fields(1)) Then
                        monto = monto + ((10 * montodpto) / 100)
                    End If
                    datCita.Recordset.MoveNext
                Loop
                datChequeAbogado.Recordset.Fields(2) = monto
                MsgBox "Cheque de adicionales generado", vbInformation
                datChequeAbogado.Recordset.Update
                frmMain.Show
                Unload Me
                Exit Sub
            End If
            datAbogado.Recordset.MoveNext
        Loop
        MsgBox "Ese abogado no existe", vbExclamation
    End If
End Sub

Private Sub Form_Activate()
    If datChequeAbogado.Recordset.EOF Then
        Exit Sub
    Else
        datChequeAbogado.Recordset.MoveFirst
    End If
    Do While datChequeAbogado.Recordset.EOF = False
        List1.AddItem datChequeAbogado.Recordset.Fields(0)
        List2.AddItem datChequeAbogado.Recordset.Fields(1)
        List3.AddItem datChequeAbogado.Recordset.Fields(2)
        List4.AddItem datChequeAbogado.Recordset.Fields(3)
        datChequeAbogado.Recordset.MoveNext
    Loop
End Sub
