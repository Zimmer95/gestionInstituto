VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11130
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuingreso 
         Caption         =   "Ingreso de datos"
         Begin VB.Menu mnuconsutas 
            Caption         =   "Consultas"
         End
      End
   End
   Begin VB.Menu mnusalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuconsutas_Click()
Form3.Show
End Sub

Private Sub mnuingreso_Click()
Form2.Show
End Sub

Private Sub mnusalir_Click()
If MsgBox("¿Desea salir?", vbYesNo + vbCritical + vbDefaultButton2, "menu") = vbYes Then
    End
    End If
End Sub
