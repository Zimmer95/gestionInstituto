VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5175
   ClientLeft      =   4650
   ClientTop       =   3240
   ClientWidth     =   9015
   LinkTopic       =   "Form3"
   ScaleHeight     =   5175
   ScaleWidth      =   9015
   Begin VB.Frame Frame1 
      Caption         =   "Trabaja"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1080
      TabIndex        =   20
      Top             =   3120
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Si"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar elemento"
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5280
      TabIndex        =   14
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5280
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar lista"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3720
      TabIndex        =   9
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Indice"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Texto"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese DNI"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Edad"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form3.Hide
End Sub

Private Sub Command2_Click()
If MsgBox("¿Borrar los datos?", vbYesNo + vbCritical + vbDefaultButton2, "Alta de cliente") = vbYes Then
    List1.Clear
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
Else

End If
End Sub

Private Sub Command3_Click()
List1.AddItem (Text1.Text)
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command5_Click()
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Command6_Click()
Text2.Text = List1.List(List1.ListIndex)
End Sub

Private Sub List1_Click()
Label11.Caption = List1.ListIndex
End Sub
