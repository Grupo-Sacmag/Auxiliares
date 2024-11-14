VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AuxForm 
   Caption         =   "Auxiliar : "
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9990
   Icon            =   "AuxForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid May3 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu AxImp 
      Caption         =   "&Impresión"
      Begin VB.Menu ImpImp 
         Caption         =   "&Imprimir"
      End
   End
   Begin VB.Menu AxEdic 
      Caption         =   "&Edicion"
      Begin VB.Menu AxEdCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdSelT 
         Caption         =   "&Seleccionar todo"
         Shortcut        =   ^S
      End
      Begin VB.Menu EdSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdDep 
         Caption         =   "&Depurar"
      End
      Begin VB.Menu EdSep3 
         Caption         =   "-"
      End
      Begin VB.Menu orPol 
         Caption         =   "&Ordenar por poliza"
      End
   End
End
Attribute VB_Name = "AuxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Guia As Long, Nombre_Auxiliar, Tabu_l As Long, Max_Tab As Long, Te_xto As String
Dim saldto As Currency, conteo As Integer, conteonum As Integer, Saldo_Par As Currency
Dim Fin_ax As Long, MientraS, Mien_tras As Integer

Sub rotulos()
    Printer.FontBold = True
    Tabu_l = 1200
    For l = 0 To 5
        Printer.CurrentX = Tabu_l
        centrar pone, May3.ColWidth(l), May3.TextMatrix(0, l)
        Printer.CurrentX = Printer.CurrentX + pone
        Printer.Print May3.TextMatrix(0, l);
        Tabu_l = Tabu_l + May3.ColWidth(l)
    Next l
    Printer.Print
    Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 30), , BF
    Printer.FontBold = False
End Sub
Sub Titulos()
On Error GoTo error
    Max_Tab = 1200
    For l = 0 To 5: Max_Tab = Max_Tab + May3.ColWidth(l): Next l
    Printer.FontSize = 10
    Te_xto = RTrim(Datos.D1)
    centrar pone, Max_Tab, Te_xto
    Printer.CurrentX = 1200 + pone
    Printer.Print RTrim(Datos.D1)
    Printer.CurrentX = 1200
    If ultimo.num = 0 Then
        Printer.Print RTrim(Mayor.May1.TextMatrix(Mayor.May1.Row, 0)); " "; RTrim(Mayor.May1.TextMatrix(Mayor.May1.Row, 1));
        Else
        Printer.Print RTrim(CATMAY.B1); " "; RTrim(CATMAY.B2);
        
    End If
    Printer.CurrentX = Max_Tab - 900
    Printer.Print "Año : "; Datos.a_o
    Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 30), , BF
    Printer.FontSize = 8
    
Exit Sub
error:
    MsgBox (Err.Number & " " & Err.Description)
End Sub



Private Sub AxEdCopiar_Click()
   Dim Temporal1
   Clipboard.Clear
   Temporal1 = Temporal1 + AuxForm.Caption & Chr(13)
   difer = May3.RowSel - May3.Row
   For i = May3.Row To May3.RowSel
      For F = May3.Col To May3.ColSel
            Temporal1 = Temporal1 + May3.TextMatrix(i, F)
            If F < May3.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next F
      Temporal1 = Temporal1 & Chr(13)
      
   Next i
    Clipboard.SetText Temporal1
   'May3.Row = 0: May3.Col = 0
  
End Sub

Private Sub EdDep_Click()
  Depuracion.Show 1
End Sub

Private Sub EdSelT_Click()
   May3.Col = 0: May3.Row = 0
   May3.RowSel = May3.Rows - 1
   May3.ColSel = May3.Cols - 2
     
End Sub

Private Sub ImpImp_Click()
    Tam_imp = 60
    conteo = 0
    Titulos
    rotulos
    For r = 1 To May3.Rows - 1
      conteo = conteo + 1
      For l = 0 To 2
        
       If l > 0 Then
              
            Tabu_l = Tabu_l + May3.ColWidth(l - 1)
            Else
            Tabu_l = 1200
            Printer.CurrentX = Tabu_l
       End If
       If May3.TextMatrix(r, l) <> "" Then
          If l = 0 Then Printer.Print May3.TextMatrix(r, l);
          If l = 1 Then
                bala = May3.TextMatrix(r, l)
                valor$ = Format(bala, "#####0"): uso$ = "#####0"
                pone = 0: colocar pone, valor$, uso$:
                Printer.CurrentX = Printer.CurrentX + pone
                Printer.Print valor$;
          End If
                
          If l = 2 Then Printer.Print "   "; Format(May3.TextMatrix(r, l), "&&&");
       End If
      Next l
      For l = 3 To 5
         Tabu_l = Tabu_l + May3.ColWidth(l - 1)
         Printer.CurrentX = Tabu_l
         If May3.TextMatrix(r, l) <> "" Then
                bala = May3.TextMatrix(r, l)
                valor$ = Format(bala, "##,###,##0.00"): uso$ = "##,###,##0.00"
                pone = 0: colocar pone, valor$, uso$:
                Printer.CurrentX = Printer.CurrentX + pone
                Printer.Print valor$;
         End If
      Next l
      Printer.Print
      If conteo = (Tam_imp - 2) Then
                
                pie
                mueve
                conteo = 0
                
                Titulos
                rotulos
       End If
    Next r
    If conteo < (Tam_imp - 2) Then recorre
    Printer.EndDoc
    conteonum = 0
    If impresionactivada = 1 Then impresionactivada = 0
End Sub
 Sub mueve()
     conteonum = conteonum + 1
    Select Case Tam_imp
                   Case 60
                        Printer.CurrentY = Printer.ScaleHeight
                        conteonum = 0
                        Rem Printer.NewPage
                   Case 20
                      Select Case conteonum
                        Case 1
                        Printer.CurrentY = Printer.ScaleHeight / 3
                        
                        Case 2
                        Printer.CurrentY = (Printer.ScaleHeight / 3 * 2)
                        Case 3
                        Printer.CurrentY = Printer.ScaleHeight
                        
                        Rem Printer.NewPage
                        conteonum = 0
                      End Select
                   Case 30
                        Select Case conteonum
                            Case 1
                            Printer.CurrentY = Printer.ScaleHeight / 2
                            Case 2
                            Printer.CurrentY = Printer.ScaleHeight
                            
                            Rem Printer.NewPage
                            conteonum = 0
                        End Select
                End Select
    
 End Sub
 Sub recorre()
     
    For i = conteo To (Tam_imp - 2)
        Printer.Print Chr(160)
                
    Next i
    pie
    mueve
    
 End Sub
Sub pie()
  Printer.FontSize = 10
  Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 20), , BF
  If ultimo.num = 0 Then
      Printer.Print Tab(20); AuxSub.May2.TextMatrix(AuxSub.May2.Row, 0); " "; AuxSub.May2.TextMatrix(AuxSub.May2.Row, 1)
  Else
      Printer.Print
      Printer.Print Tab(20); RTrim(Anexo.Anexo1.TextMatrix(Anexo.Anexo1.Row, 0)); " "; RTrim(Anexo.Anexo1.TextMatrix(Anexo.Anexo1.Row, 1))
  End If
  Printer.FontSize = 8

End Sub

Private Sub May3_KeyPress(KeyAscii As Integer)
On Error GoTo Zacatecas
   Select Case KeyAscii
   Case 27
      Clipboard.Clear
      Unload AuxForm
   Case 13
      Open "DATOS" For Random As 11 Len = Len(Datos)
      Get 11, 1, Datos
      llave = RTrim(Datos.No_arch)
      Close 11
      fechatran = Mid(May3.TextMatrix(May3.Row, 0), 4, 2)
      flecha = Mid(May3.TextMatrix(May3.Row, 0), 1, 6)
      If flecha = "01/01/" Then fechatran = "13"
      
      Arch_Act = llave + fechatran
      numpol = May3.TextMatrix(May3.Row, 1)
      Open Arch_Act For Random As 3 Len = Len(oper)
      jm = LOF(3) / Len(oper)
      
      For r = 1 To jm: Get 3, r, oper
      If oper.identi = "A" Then
          If Val(oper.cta) = numpol Then
                Rango_Inic = numpol
                Rango_Final = numpol
                ultimo.num = r
                Mien_tras = Mes_Act
                Mes_Act = fechatran

                Exit For
          End If
      End If
      Next r
      Close 3
      Poliza1.ArPol.Visible = False
      Poliza1.verpoliza
      Poliza1.Show 1
      
      Unload Poliza1
      Mes_Act = Mien_tras
   End Select
Zacatecas:
End Sub

Private Sub May3_LeaveCell()
If May3.Rows > 1 Then
    If May3.Col > 0 And May3.Row > 0 Then
         May3.CellBackColor = vbWhite
    End If
 End If
End Sub
    
Private Sub May3_ENTERCell()
 If (May3.Row >= 1) And (May3.Col >= 1) Then
  May3.CellBackColor = vbYellow
 End If
End Sub

Private Sub Form_Load()
    conteo = 0
    If ultimo.num = 0 Then
        If Mes_imp = 0 Then
            Label1.Caption = "Se imprimira el Auxiliar Completo "
            Else
            Label1.Caption = "Se imprimira el Auxiliar del mes de " + mm(Mes_imp)
        End If
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
    If ultimo.num = 0 Then
        AuxForm.Caption = "Auxiliar :" & AuxSub.May2.TextMatrix(AuxSub.May2.Row, 0) & _
                      AuxSub.May2.TextMatrix(AuxSub.May2.Row, 1) & AuxSub.May2.TextMatrix(AuxSub.May2.Row, 2)
                      
                      If IsNumeric(AuxSub.May2.TextMatrix(AuxSub.May2.Row, 3)) Then
                                    Guia = AuxSub.May2.TextMatrix(AuxSub.May2.Row, 3)
                      End If
    Else
                      
        'AuxForm.Caption = "Auxiliar :" & Anexo.Anexo1.TextMatrix(Anexo.Anexo1.Row, 0) & Anexo.Anexo1.TextMatrix(Anexo.Anexo1.Row, 1) & Anexo.Anexo1.TextMatrix(Anexo.Anexo1.Row, 2)
        
        Guia = Anexo.Anexo1.TextMatrix(Anexo.Anexo1.Row, 3)
    End If
    
    Nombre_Auxiliar = "AUXILIAR\AX" + LTrim(Str(Guia))

    May3.Clear
    May3.Row = 0
    May3.Row = 0
    May3.Col = 0: May3.ColWidth(0) = 900: May3.CellAlignment = 4: May3.Text = "Fecha"
    May3.Col = 1: May3.ColWidth(1) = 800:  May3.CellAlignment = 4: May3.Text = "Poliza"
    May3.Col = 2: May3.ColWidth(2) = 3200: May3.CellAlignment = 4: May3.Text = "Descripción"
    May3.Col = 3: May3.ColWidth(3) = 1200:  May3.CellAlignment = 4: May3.Text = "Debe"
    May3.Col = 4: May3.ColWidth(4) = 1200:  May3.CellAlignment = 4: May3.Text = "Haber"
    May3.Col = 5: May3.ColWidth(5) = 1200:  May3.CellAlignment = 4: May3.Text = "Saldo"
    May3.Col = 6: May3.ColWidth(6) = 80: May3.CellAlignment = 4: May3.Text = ""
    May3.Rows = 1
    May3.Row = 0
    May3.Col = 1
    Open Nombre_Auxiliar For Random As 6 Len = Len(auxiliar)
    Fin_ax = LOF(6) / Len(auxiliar)
    If Mes_imp > 0 Then saldoini
    If Fin_ax <= 0 Then
         MsgBox "El auxiliar solicitado no Contiene movimientos "
         Close 6
         Exit Sub
         Else
         saldto = 0: If Mes_imp > 0 Then saldto = Saldo_Par
         For r = 1 To Fin_ax: Get 6, r, auxiliar
             If (Mes_imp = 0) Or (Mes_imp = Val(Mid(auxiliar.fech, 4, 2))) Then
                 If auxiliar.impo > 0 Then
                     Debe = Format(auxiliar.impo, "###,###,##0.00")
                     Haber = ""
                     Else
                     Haber = Format(auxiliar.impo, "###,###,##0.00")
                     Debe = ""
                 End If
                 If Len(RTrim(Mid(auxiliar.fech, 1, 2))) < 2 Then
                     MientraS = "0" + RTrim(Mid(auxiliar.fech, 1, 2))
                     Mid(auxiliar.fech, 1, 2) = MientraS
                 End If
                 If Mid(auxiliar.fech, 4, 2) = "13" Then Mid(auxiliar.fech, 4, 2) = "00"
                 fechita = Right(auxiliar.fech, 2) + "/" + Mid(auxiliar.fech, 4, 2) + "/" + Left(auxiliar.fech, 2) + "/" + String(6 - Len(LTrim(Str(auxiliar.po))), "0") + LTrim(Str(auxiliar.po))
                 Rem If Len(LTrim(Mid(auxiliar.fech, 1, 2))) < 2 Then
                 saldto = saldto + auxiliar.impo
                 May3.AddItem auxiliar.fech & Chr(9) & auxiliar.po & Chr(9) & _
                 (" " + auxiliar.re) & Chr(9) & Debe & Chr(9) & Haber & Chr(9) & _
                 Format(saldto, "###,###,##0.00") & Chr(9) & fechita
                 Else
                 Rem nada
             End If
         Next r
         If May3.Rows < 2 Then
           
           MsgBox "El auxiliar solicitado no Contiene movimientos "
           Close
           Exit Sub
         End If
         Close 6
    End If
    colanti = May3.Col
    renati = May3.Row
    May3.Row = 1
    May3.Col = 6
    May3.RowSel = May3.Rows - 1
    May3.Sort = 5
    May3.Col = colanti
    May3.Row = renati
    calsaldo
    
    If impresionactivada = 1 Then ImpImp_Click
End Sub
Sub saldoini()
   
   Saldo_Par = 0
   For r = 1 To Fin_ax: Get 6, r, auxiliar
        If (Val(Mid(auxiliar.fech, 4, 2)) < Mes_imp) Or (Val(Mid(auxiliar.fech, 4, 2)) = 13) Then
                    Saldo_Par = Saldo_Par + auxiliar.impo
        End If
   Next r
   fechita = "00/00/000000"
   saldto = Saldo_Par
   May3.AddItem "00/00/00" & Chr(9) & "0" & Chr(9) & _
                 (" --> SALDO INICIAL") & Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                 Format(saldto, "###,###,##0.00") & Chr(9) & fechita

   
End Sub
Sub calsaldo()
  saldto = 0: If Mes_imp > 0 Then saldto = Saldo_Par
  
  For r = 1 To May3.Rows - 1
       If IsNumeric(May3.TextMatrix(r, 3)) Then
                saldto = saldto + May3.TextMatrix(r, 3)
                ElseIf May3.TextMatrix(r, 4) <> "" Then
                saldto = saldto + May3.TextMatrix(r, 4)
       End If
       May3.TextMatrix(r, 5) = Format(saldto, "###,###,##0.00")
  Next r
End Sub

Private Sub orPol_Click()
    colanti = May3.Col
    renati = May3.Row
    May3.Row = 1
    May3.Col = 1
    May3.RowSel = May3.Rows - 1
    May3.Sort = 5
    May3.Col = colanti
    May3.Row = renati
End Sub
