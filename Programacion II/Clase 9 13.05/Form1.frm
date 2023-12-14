VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sexo"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   2880
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Masculino"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Femenino"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carrera"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2655
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "Form1.frx":0000
         Left            =   240
         List            =   "Form1.frx":000A
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":0021
      Left            =   1800
      List            =   "Form1.frx":002B
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Localidad"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "DNI"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido y Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Matricula"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ops As String
Private Sub limpia()
Command1.Caption = "Ingresar"
Text1.Text = ""
    Text1.Enabled = False
Text2.Text = ""
    Text2.Enabled = False
Text3.Text = ""
    Text3.Enabled = False
Combo1.Text = ""
    Combo1.Enabled = False
Option1 = False
Option2 = False
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Frame1.Enabled = True
    List1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Ingresar" Then
    Text1.Enabled = True
    Text1.SetFocus
Else
If Command1.Caption = "Guardar" Then
    If MsgBox("¿Desea guardar los datos?", vbYesNo + vbCritical + vbDefaultButton2, "Regristro") = vbYes Then
        If ops = "Femenino" Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If

    Dim reg As alumnos
        Open App.Path + "archivo.txt" For Append As 1 Len = Len(reg)
            Write #1, Text1.Text, Text2.Text, Text3.Text, Combo1.Text, List1.Text, ops
            Close #1    'cierra el #1
            'All         'cierra todo
              
        limpia
    End If
End If
End If
End Sub


Private Sub Command2_Click()
Dim reg As alumnos
    Open App.Path + "archivo.txt" For Input As 1 Len = Len(reg)
    
    buscar = InputBox("Ingrese Matricula")
    While Not EOF(1)
        Input #1, reg.matriculo, reg.apellidoynombre, reg.dni, reg.localidad, reg.carrera, reg.sexo
            If reg.matriculo = buscar Then
                Text1.Text = reg.matriculo
                Text2.Text = reg.apellidoynombre
                Text3.Text = reg.dni
                Combo1.Text = reg.localidad
                List1.Text = reg.carrera
                ops = reg.sexo
                
            End If
        
    Wend
    Close #1
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Frame2.Enabled = True
    Option1.SetFocus
End If
End Sub

Private Sub Option1_Click()
Command1.Caption = "Guardar"
ops = "Femenino"
    Command1.SetFocus
End Sub

Private Sub Option2_Click()
Command1.Caption = "Guardar"
ops = "Masculino"
    Command1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text2.Enabled = True
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Text3.Enabled = True
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> Empty Then
    Combo1.Enabled = True
    Combo1.SetFocus
End If
End Sub

