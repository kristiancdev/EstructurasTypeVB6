VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TestType()
    Dim persona1 As Persona
    persona1.Id = 1
    persona1.Nombre = "Juan"
    persona1.Apellido = "Pérez"
    persona1.FechaNacimiento = #1/15/1990#
    
    ' Mostrar datos
    MsgBox "Nombre: " & persona1.Nombre & " " & persona1.Apellido
End Sub

Private Sub Form_Load()
    TestType
End Sub
