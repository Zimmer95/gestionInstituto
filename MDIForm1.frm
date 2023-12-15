VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000002&
   Caption         =   "MDIForm1"
   ClientHeight    =   9105
   ClientLeft      =   3495
   ClientTop       =   1185
   ClientWidth     =   12915
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu menu 
      Caption         =   "MENU"
      Begin VB.Menu mnu_alumnos 
         Caption         =   "Alumnos"
      End
      Begin VB.Menu mnu_carreras 
         Caption         =   "Carreras"
      End
      Begin VB.Menu mnu_profesores 
         Caption         =   "Profesores"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
frm_menu_principal.Show
frm_alumnos.Hide
End Sub

Private Sub mnu_alumnos_Click()
frm_alumnos.Show
frm_profesores.Hide
frm_carreras.Hide
frm_menu_principal.Hide
End Sub

Private Sub mnu_carreras_Click()
frm_alumnos.Hide
frm_menu_principal.Hide
frm_profesores.Hide
frm_carreras.Show
End Sub

Private Sub mnu_profesores_Click()
frm_alumnos.Hide
frm_carreras.Hide
frm_menu_principal.Hide
frm_profesores.Show
End Sub
