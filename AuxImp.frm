VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Imprimir 
   Caption         =   "Imprimiendo auxiliar"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   Icon            =   "AuxImp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid May4 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   7
      BackColor       =   16777215
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
 If Mes_imp = 0 Then
         Label1.Caption = "Se imprimira el Auxiliar Completo "
         Else
         Label1.Caption = "Se imprimira el Auxiliar del mes de " + mm(Mes_imp)
    End If
    Select Case Tam_imp
       Case 20
        Label2.BackColor = vbBlue
        Label2.ForeColor = vbWhite
        Label2.FontUnderline = True
        Label2.Caption = "La impresión sera de 20 Renglones por Auxiliar"
       Case 30
        Label2.BackColor = vbYellow
        Label2.FontUnderline = True
        Label2.Caption = "La impresión sera de 30 Renglones por Auxiliar"
       Case 60
        
        Label2.Caption = "La impresión sera de una hoja completa"
    End Select
    
    Imprimir.Caption = "Auxiliar :" & AuxSub.May2.TextMatrix(AuxSub.May2.Row, 0) & _
                      AuxSub.May2.TextMatrix(AuxSub.May2.Row, 1) & AuxSub.May2.TextMatrix(AuxSub.May2.Row, 2)
    
End Sub

Private Sub Form_Load()
    
    conteo = 0
   
    May4.Clear
    May4.Row = 0
    May4.Row = 0
    May4.Col = 0: May4.ColWidth(0) = 900: May4.CellAlignment = 4: May4.Text = "Fecha"
    May4.Col = 1: May4.ColWidth(1) = 800:  May4.CellAlignment = 4: May4.Text = "Poliza"
    May4.Col = 2: May4.ColWidth(2) = 3200: May4.CellAlignment = 4: May4.Text = "Descripción"
    May4.Col = 3: May4.ColWidth(3) = 1200:  May4.CellAlignment = 4: May4.Text = "Debe"
    May4.Col = 4: May4.ColWidth(4) = 1200:  May4.CellAlignment = 4: May4.Text = "Haber"
    May4.Col = 5: May4.ColWidth(5) = 1200:  May4.CellAlignment = 4: May4.Text = "Saldo"
    May4.Col = 6: May4.ColWidth(6) = 80:  May4.CellAlignment = 4: May4.Text = ""
    
    May4.Rows = 1
    May4.Row = 0
    May4.Col = 1
    Guia = AuxSub.May2.TextMatrix(AuxSub.May2.Row, 3)
    Nombre_Auxiliar = "AUXILIAR\AX" + LTrim(Str(Guia))

    Open Nombre_Auxiliar For Random As 6 Len = Len(auxiliar)
    fin_ax = LOF(6) / Len(auxiliar)
    If fin_ax <= 0 Then
         MsgBox "El auxiliar solicitado no Contiene movimientos "
         Close
         Exit Sub
         Else
         saldto = 0
         For r = 1 To fin_ax: Get 6, r, auxiliar
                 If auxiliar.impo > 0 Then
                     debe = Format(auxiliar.impo, "###,###,##0.00")
                     haber = ""
                     Else
                     haber = Format(auxiliar.impo, "###,###,##0.00")
                     debe = ""
                 End If
                 fechita = Mid(auxiliar.fech, 4, 2) + "/" + Left(auxiliar.fech, 2)
                 saldto = saldto + auxiliar.impo
                 May4.AddItem auxiliar.fech & Chr(9) & auxiliar.po & Chr(9) & _
                 (" " + auxiliar.re) & Chr(9) & debe & Chr(9) & haber & Chr(9) & _
                 Format(saldto, "###,###,##0.00") & Chr(9) & fechita
         Next r
         Close 6
    End If
    colanti = May4.Col
    renati = May4.Row
    May4.Row = 1
    May4.Col = 6
    May4.RowSel = May4.Rows - 1
    May4.Sort = 5
    May4.Col = colanti
    May4.Row = renati
    calsaldo1
    
    Rem Form_Unload vbCancel
End Sub
Sub calsaldo1()
     saldto = 0
  For r = 1 To May4.Rows - 1
       If May4.TextMatrix(r, 3) <> "" Then
                saldto = saldto + May4.TextMatrix(r, 3)
                Else
                saldto = saldto + May4.TextMatrix(r, 4)
       End If
       May4.TextMatrix(r, 5) = Format(saldto, "###,###,##0.00")
  Next r

End Sub
Private Sub Form_Unload(Cancel As Integer)
   'Dim Msg, Response   ' Declara variables.
   'Msg = "¿Desea guardar los datos antes de cerrar?"
   'Response = MsgBox(Msg, vbQuestion + vbYesNoCancel, "Diálogo Cerrar")
   'Select Case Response
      'Case vbCancel   ' No se permite cerrar.
         'Cancel = -1
         'Msg = "Se ha cancelado el comando."
      'Case vbYes
      ' Introduzca código para guardar los datos aquí.
         'Msg = "Datos guardados." '
      'Case vbNo
         'Msg = "Datos no guardados."
   'End Select
   'MsgBox Msg, vbOKOnly, "Confirmación"   ' Mostrar mensaje.
   Unload Imprimir
End Sub
Sub mexico()
 Rem a ver que onda

End Sub
