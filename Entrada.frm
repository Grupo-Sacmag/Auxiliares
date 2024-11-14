VERSION 5.00
Begin VB.Form Entrada 
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   Icon            =   "Entrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox texto 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label etiqueta 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  texto_KeyPress 13
  
End Sub

Private Sub Command2_Click()
  texto.Text = ""
  Unload Entrada
End Sub

Private Sub texto_Click()
  texto.Text = UCase(texto.Text)
End Sub

Private Sub texto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        texto.Text = UCase(texto.Text)
        ultimo.texto = texto.Text
        Unload Entrada
    End If
End Sub
