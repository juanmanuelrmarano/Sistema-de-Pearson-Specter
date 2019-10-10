VERSION 5.00
Begin VB.Form frmCalificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEARSON SPECTER - Calificacion"
   ClientHeight    =   3000
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   3000
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtCalif 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   2895
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtApellido 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nueva Calificacion"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Abogado"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCalificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If IsNumeric(txtNombre) Or IsNumeric(txtApellido) Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If IsNumeric(txtCalif) = False Then
        MsgBox "Datos invalidos", vbExclamation
        Exit Sub
    End If
    If txtNombre = "" Or txtApellido = "" Or txtCalif = "" Then
        MsgBox "Ha dejado uno de los campos vacio", vbExclamation
        Exit Sub
    Else
        If datAbogado.Recordset.EOF Then
            Exit Sub
        Else
            datAbogado.Recordset.MoveFirst
        End If
        Do While datAbogado.Recordset.EOF = False
            If txtNombre = datAbogado.Recordset.Fields(0) And txtApellido = datAbogado.Recordset.Fields(1) Then
                If txtCalif > 10 Or txtCalif < 0 Then
                    MsgBox "Esa calificacion no es valida", vbExclamation
                    Exit Sub
                Else
                    datAbogado.Recordset.Edit
                    datAbogado.Recordset.Fields("calificacion") = txtCalif
                    datAbogado.Recordset.Update
                    MsgBox "Calificacion registrada", vbInformation
                    txtNombre = ""
                    txtApellido = ""
                    txtCalif = ""
                    frmMain.Show
                    Unload Me
                    Exit Sub
                    End If
            Else
                datAbogado.Recordset.MoveNext
            End If
        Loop
        MsgBox "El nombre o el apellido no existen", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdVolver_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Activate()
    If datAbogado.Recordset.EOF Then
        Exit Sub
    Else
        datAbogado.Recordset.MoveFirst
    End If
    Do While datAbogado.Recordset.EOF = False
        List1.AddItem datAbogado.Recordset.Fields(0)
        List2.AddItem datAbogado.Recordset.Fields(1)
        List3.AddItem datAbogado.Recordset.Fields("calificacion")
        datAbogado.Recordset.MoveNext
    Loop
End Sub

