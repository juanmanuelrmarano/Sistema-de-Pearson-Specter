VERSION 5.00
Begin VB.Form frmNuevoTitulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Nuevo titulo"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datTitulos 
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
      RecordSource    =   "titulos"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Data datAbogado 
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
      RecordSource    =   "abogado"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label13 
      Caption         =   "AAAA"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   " MM"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   " DD"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "DNI Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre titulo"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Ente emisor"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Año de titulacion"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmNuevoTitulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AboVal As Boolean
Dim Cod As Integer

Private Sub cmdAceptar_Click()
     If IsNumeric(Text3) Or IsNumeric(Text2) Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Text1) = False Or IsNumeric(Text4) = False Or IsNumeric(Text5) = False Or IsNumeric(Text6) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
 If Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "Ha dejado uno de los campos vacio", vbExclamation
    Exit Sub
 Else
    datTitulos.Recordset.MoveLast
    Cod = datTitulos.Recordset.Fields(0) + 1
    datTitulos.Recordset.MoveFirst
    datAbogado.Recordset.MoveFirst
    AboVal = False
    datTitulos.Recordset.AddNew
    datTitulos.Recordset.Fields(0) = Cod
    Do While datAbogado.Recordset.EOF = False
        If Text1 = datAbogado.Recordset.Fields(2) Then
            datTitulos.Recordset.Fields(1) = Text1
            AboVal = True
            Exit Do
        End If
        datAbogado.Recordset.MoveNext
    Loop
    If AboVal = False Then
        MsgBox "El abogado especificado no existe", vbExclamation
        Exit Sub
    End If
    If IsDate(Text4 & "/" & Text5 & "/" & Text6) Then
            datTitulos.Recordset.Fields(2) = Text4 & "/" & Text5 & "/" & Text6
        Else
            MsgBox "No es una fecha valida", vbExclamation
            Exit Sub
    End If
    datTitulos.Recordset.Fields(3) = Text3
    datTitulos.Recordset.Fields(4) = Text2
    MsgBox "Titulo añadido", vbInformation
    datTitulos.Recordset.Update
    Unload Me
    frmMain.Show
 End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub
