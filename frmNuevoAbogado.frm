VERSION 5.00
Begin VB.Form frmNuevoAbogado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Nuevo abogado"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   30
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   26
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Data datDpto 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "PearsonSpecter.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "departamento"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cboDpto 
      Height          =   315
      ItemData        =   "frmNuevoAbogado.frx":0000
      Left            =   1680
      List            =   "frmNuevoAbogado.frx":0002
      TabIndex        =   24
      Text            =   "Seleccionar..."
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   19
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Top             =   3720
      Width           =   3495
   End
   Begin VB.ComboBox cboPuesto 
      Height          =   315
      ItemData        =   "frmNuevoAbogado.frx":0004
      Left            =   1680
      List            =   "frmNuevoAbogado.frx":000E
      TabIndex        =   17
      Text            =   "Seleccionar..."
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   4560
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label19 
      Caption         =   "(M/F)"
      Height          =   255
      Left            =   2160
      TabIndex        =   36
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   " DD"
      Height          =   255
      Left            =   1680
      TabIndex        =   35
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   " MM"
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "AAAA"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   " DD"
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   " MM"
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "AAAA"
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de nacimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha de ingreso"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Arancel"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Departamento"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Jerarquia"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "DNI"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Telefono"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Celular"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Direccion"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Sexo"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "frmNuevoAbogado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If Text11 < 0 Or Text11 > 31 Or Text10 < 0 Or Text10 > 31 Then
        MsgBox "Dia invalido", vbExclamation
        Exit Sub
    End If
    If Text12 < 0 Or Text12 > 12 Or Text14 < 0 Or Text14 > 12 Then
        MsgBox "Mes invalido", vbExclamation
        Exit Sub
    End If
    If Text8 < 1000 Or Text8 > 9999 Or Text13 < 1000 Or Text13 > 9999 Then
        MsgBox "Año invalido", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Text2) Or IsNumeric(Text3) Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Text1) = False Or IsNumeric(Text5) = False Or IsNumeric(Text11) = False Or IsNumeric(Text12) = False Or IsNumeric(Text8) = False Or IsNumeric(Text10) = False Or IsNumeric(Text14) = False Or IsNumeric(Text13) = False Or IsNumeric(Text9) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Then
        MsgBox "Ha dejado uno de los campos vacio", vbExclamation
        Exit Sub
    Else
        Do While datAbogado.Recordset.EOF = False
            If Text1 = datAbogado.Recordset.Fields(2) Then
                MsgBox "Ya hay un abogado con ese DNI", vbExclamation
                Exit Sub
            End If
            datAbogado.Recordset.MoveNext
        Loop
        datAbogado.Recordset.AddNew
        datAbogado.Recordset.Fields(2) = Text1
        datAbogado.Recordset.Fields(0) = Text2
        datAbogado.Recordset.Fields(1) = Text3
        datAbogado.Recordset.Fields(3) = Text4
        datAbogado.Recordset.Fields(4) = Text5
        datAbogado.Recordset.Fields(5) = Text6
        datAbogado.Recordset.Fields(6) = Text7
        Select Case cboPuesto.Text
        Case "Junior"
            datAbogado.Recordset.Fields(7) = "Junior"
        Case "Senior"
            datAbogado.Recordset.Fields(7) = "Senior"
        Case Else
          cboPuesto.Text = "Seleccionar..."
           MsgBox "Puesto invalido", vbCritical
           Exit Sub
        End Select
        If IsDate(Text11 & "/" & Text12 & "/" & Text8) Then
            datAbogado.Recordset.Fields(8) = Text11 & "/" & Text12 & "/" & Text8
        Else
            MsgBox "No es una fecha valida", vbExclamation
            Exit Sub
        End If
        If IsDate(Text10 & "/" & Text14 & "/" & Text13) Then
            datAbogado.Recordset.Fields(9) = Text10 & "/" & Text14 & "/" & Text13
        Else
            MsgBox "No es una fecha valida", vbExclamation
            Exit Sub
        End If
        datAbogado.Recordset.Fields(10) = Text9
        Select Case cboDpto.Text
        Case "Penal"
            datAbogado.Recordset.Fields("departamento") = 0
        Case "Civil"
            datAbogado.Recordset.Fields("departamento") = 1
        Case "Constitucional"
            datAbogado.Recordset.Fields("departamento") = 2
        Case "Administrativo"
            datAbogado.Recordset.Fields("departamento") = 3
        Case "Procesal"
            datAbogado.Recordset.Fields("departamento") = 4
        Case "Financiero"
            datAbogado.Recordset.Fields("departamento") = 5
        Case "Tributario"
            datAbogado.Recordset.Fields("departamento") = 6
        Case Else
            cboDpto.Text = "Seleccionar..."
            MsgBox "Puesto invalido", vbCritical
            Exit Sub
        End Select
        datAbogado.Recordset.Update
        MsgBox "Abogado registrado", vbInformation
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
        cboDpto.AddItem datDpto.Recordset.Fields(1)
        datDpto.Recordset.MoveNext
    Loop
End Sub
