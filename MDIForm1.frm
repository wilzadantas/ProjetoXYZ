VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Projeto XYZ"
   ClientHeight    =   7770
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13530
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Principal 
      Caption         =   "Cadastro/Consulta"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Principal_Click()
    frmPrincipal.Show
End Sub
