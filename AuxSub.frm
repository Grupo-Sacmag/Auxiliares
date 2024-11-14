VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AuxSub 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Subcuentas"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8895
   Icon            =   "AuxSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid May5 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   7
      BackColorBkg    =   14737632
      BorderStyle     =   0
   End
   Begin MSFlexGridLib.MSFlexGrid May2 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   4
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
   Begin VB.Menu EdT 
      Caption         =   "&Edicion"
      Begin VB.Menu EdCpr 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdSlTd 
         Caption         =   "&Seleccionar Todo"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu impre 
      Caption         =   "&Impresion"
      Begin VB.Menu ImpCta 
         Caption         =   "&Cuenta"
      End
      Begin VB.Menu ImpAux 
         Caption         =   "&Auxiliar"
      End
   End
   Begin VB.Menu SubOrd 
      Caption         =   "&Ordenar"
      Begin VB.Menu OrdAlf 
         Caption         =   "Alfabeticamente"
      End
      Begin VB.Menu OrdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu OrdNum 
         Caption         =   "&Numericamente"
      End
      Begin VB.Menu OrdSep2 
         Caption         =   "-"
      End
      Begin VB.Menu OrdDep 
         Caption         =   "&Depurar"
      End
   End
   Begin VB.Menu AxFor 
      Caption         =   "&Formato"
      Begin VB.Menu ForMes 
         Caption         =   "&Mes"
         Begin VB.Menu Mes1 
            Caption         =   "&Enero"
         End
         Begin VB.Menu Mes2 
            Caption         =   "&Febrero"
         End
         Begin VB.Menu Mes3 
            Caption         =   "&Marzo"
         End
         Begin VB.Menu Mes4 
            Caption         =   "&Abril"
         End
         Begin VB.Menu Mes5 
            Caption         =   "Ma&yo"
         End
         Begin VB.Menu Mes6 
            Caption         =   "&Junio"
         End
         Begin VB.Menu Mes7 
            Caption         =   "J&ulio"
         End
         Begin VB.Menu Mes8 
            Caption         =   "A&gosto"
         End
         Begin VB.Menu Mes9 
            Caption         =   "&Septiembre"
         End
         Begin VB.Menu Mes10 
            Caption         =   "&Octubre"
         End
         Begin VB.Menu Mes11 
            Caption         =   "&Noviembre"
         End
         Begin VB.Menu Mes12 
            Caption         =   "&Diciembre"
         End
         Begin VB.Menu mes0 
            Caption         =   "&Completo"
         End
      End
      Begin VB.Menu ForSep1 
         Caption         =   "-"
      End
      Begin VB.Menu For 
         Caption         =   "&Tamaño"
         Begin VB.Menu Tam1 
            Caption         =   "&20 Renglones"
         End
         Begin VB.Menu Tam2 
            Caption         =   "&30 Renglones"
         End
         Begin VB.Menu Tam3 
            Caption         =   "&60 Renglones"
         End
      End
   End
End
Attribute VB_Name = "AuxSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inicio As Long, final As Long, Iniciar As Integer, Te_xto As String
Dim Guia As Long, Nombre_Auxiliar, Tabu_l As Long, Max_Tab As Long
Dim Saldo_Par As Currency, Fin_ax As Long
Dim saldto As Currency, conteo As Integer, conteonum As Integer
Sub etiqueta()
   If Mes_imp = 0 Then
         
          Label1.Caption = "Se Imprimira el Auxiliar Completo "
          Label1.BackColor = vbWhite
          Label1.FontUnderline = False

          Else
          Label1.BackColor = vbCyan
          Label1.FontUnderline = True
          Label1.Caption = "Se Imprimira el Auxiliar del Mes de " + RTrim(mm(Mes_imp))
     End If
    Select Case Tam_imp
       Case 20
        Label2.BackColor = vbBlue
        Label2.ForeColor = vbWhite
        Label2.FontUnderline = True
        Label2.Caption = "La Impresión Sera De 20 Renglones Por Auxiliar"
       Case 30
        Label2.BackColor = vbYellow
        Label2.FontUnderline = True
        Label2.Caption = "La Impresión Sera De 30 Renglones Por Auxiliar"
       Case 60
         Label2.BackColor = vbWhite
        Label2.FontUnderline = False
         Label2.Caption = "La Impresión Sera De Una Hoja Completa"
    End Select

End Sub
Sub rotulos1()
On Error GoTo error
    Printer.FontBold = True
    Tabu_l = 1200
    For l = 0 To 5
        Printer.CurrentX = Tabu_l
        centrar pone, May5.ColWidth(l), May5.TextMatrix(0, l)
        Printer.CurrentX = Printer.CurrentX + pone
        Printer.Print May5.TextMatrix(0, l);
        Tabu_l = Tabu_l + May5.ColWidth(l)
    Next l
    Printer.Print
    Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 30), , BF
    Printer.Print
    Printer.FontBold = False
Exit Sub

error:
    
MsgBox (Err.Number & " " & Err.Description)
End Sub
Sub Titulos1()
    
On Error GoTo errorHandler

    Max_Tab = 1200
    For l = 0 To 5: Max_Tab = Max_Tab + May5.ColWidth(l): Next l
    Printer.FontSize = 10
    Te_xto = RTrim(Datos.D1)
    centrar pone, Max_Tab, Te_xto
    Printer.CurrentX = 1200 + pone
    Printer.Print RTrim(Datos.D1)
    Printer.CurrentX = 1200
    Printer.Print RTrim(Mayor.May1.TextMatrix(Mayor.May1.Row, 0)); " "; RTrim(Mayor.May1.TextMatrix(Mayor.May1.Row, 1));
    Printer.CurrentX = Max_Tab - 900
    Printer.Print "Año : "; Datos.a_o
    Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 30), , BF
    Printer.FontSize = 8
    

Exit Sub
errorHandler:

    MsgBox (Err.Number & " " & Err.Description)
    

End Sub

Sub saldoini1()
   
   Saldo_Par = 0
   For r = 1 To Fin_ax: Get 6, r, auxiliar
        If (Val(Mid(auxiliar.fech, 4, 2)) < Mes_imp) Or (Val(Mid(auxiliar.fech, 4, 2)) = 13) Then
                    Saldo_Par = Saldo_Par + auxiliar.impo
        End If
   Next r
   fechita = "00/00/000000"
   saldto = Saldo_Par
   May5.AddItem "00/00/00" & Chr(9) & "0" & Chr(9) & _
                 (" --> SALDO INICIAL") & Chr(9) & "" & Chr(9) & "" & Chr(9) & _
                 Format(saldto, "###,###,##0.00") & Chr(9) & fechita

   
End Sub
Sub calsaldo1()
  saldto = 0: If Mes_imp > 0 Then saldto = Saldo_Par
  
  For r = 1 To May5.Rows - 1
       If May5.TextMatrix(r, 3) <> "" Then
                saldto = saldto + May5.TextMatrix(r, 3)
                ElseIf May5.TextMatrix(r, 4) <> "" Then
                saldto = saldto + May5.TextMatrix(r, 4)
       End If
       May5.TextMatrix(r, 5) = Format(saldto, "###,###,##0.00")
  Next r
End Sub



Private Sub EdCpr_Click()
   Dim Temporal1
    Clipboard.Clear
   Temporal1 = Temporal1 + AuxForm.Caption & Chr(13) & Chr(10)
   Rem MAY2.RowSel = MAY2.Rows - 1
   Rem MAY2.ColSel = MAY2.Cols - 2
   For i = 0 To May2.RowSel
      For F = 0 To May2.ColSel
            Temporal1 = Temporal1 + May2.TextMatrix(i, F) & Chr(9)
      Next F
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
   Next i
   Clipboard.SetText Temporal1
   difer = May2.RowSel - May2.Row
   
End Sub

Private Sub EdSlTd_Click()
   May2.Col = 0: May2.Row = 0
   May2.RowSel = May2.Rows - 1
   May2.ColSel = May2.Cols - 2

End Sub

Private Sub Form_Resize()
        
        Select Case Iniciar
           Case 1
            AuxSub.Refresh
            AuxSub.Height = 9495
            AuxSub.Width = 9225
            May2.Height = 1695
            May5.Visible = True
            May5.Width = ScaleWidth * 0.98
            May5.Height = ScaleHeight * 0.9
            Rem Label1.Visible = True
            Rem Label2.Visible = True
           Case 0
            May5.Visible = False
            AuxSub.Height = 4030
            AuxSub.Width = 9225
            May2.Height = ScaleHeight * 0.9
            Rem Label1.Visible = False
            Rem Label2.Visible = False
        End Select
End Sub

Private Sub ImpAux_Click()
         impresionactivada = 1
         AuxForm.Show 1
         Clipboard.Clear
         
End Sub

Private Sub ImpCta_Click()
  For r = 1 To May2.Rows - 1
        May2.Row = r
        Iniciar = 1
        Form_Resize
        etiqueta
        imprepantalla
        Rem impresora
        Iniciar = 0
        Form_Resize
        etiqueta
  Next r
        Printer.EndDoc
        conteonum = 0

End Sub

Private Sub May2_DblClick()
   ultimo.num = 0
   AuxForm.Show 1
End Sub

Private Sub May2_GotFocus()
     Iniciar = 0
     Form_Resize
     
End Sub

Private Sub May2_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then May2_DblClick
End Sub

Private Sub May2_LeaveCell()
  If May2.Rows > 1 Then
      If (May2.Col > 0) And (May2.Row > 0) Then
               May2.CellBackColor = vbWhite
      End If
  End If
End Sub
    
Private Sub May2_ENTERCell()
     Iniciar = 0
     Form_Resize

 If (May2.Row >= 1) And (May2.Col >= 1) Then
  May2.CellBackColor = vbGreen
 End If
End Sub

Private Sub Form_Load()
    Iniciar = 0
    Tam_imp = 60
    Form_Resize
    etiqueta
    May2.Clear
    Tam_imp = 60
    Mes_imp = 0
    May2.Row = 0
    May2.Col = 0: May2.ColWidth(0) = 600: May2.CellAlignment = 4: May2.Text = "SubCta"
    May2.Col = 1: May2.ColWidth(1) = 3350:  May2.CellAlignment = 4: May2.Text = "Nombre"
    May2.Col = 2: May2.ColWidth(2) = 1200:  May2.CellAlignment = 4: May2.Text = "Importe"
    May2.Col = 3: May2.ColWidth(3) = 80:  May2.CellAlignment = 4: May2.Text = ""
    May2.Rows = 1
    
    Open "Cataux" For Random As #3 Len = Len(CATAUX)
    dm = LOF(3) / Len(CATAUX)
    inicio = Mayor.May1.TextMatrix(Mayor.May1.Row, 3)
    final = Mayor.May1.TextMatrix(Mayor.May1.Row, 4)
    For r = inicio To final: Get 3, r, CATAUX
        If Val(CATAUX.C1) > 0 Then
                May2.AddItem Format(CATAUX.C1, "#####") & Chr(9) & _
                             (" " + CATAUX.C2) & Chr(9) & _
                             Format(CATAUX.C3, "##,###,##0.00") & Chr(9) & _
                             r
        End If
    Next r
    May2.Col = 1: May2.Row = 0
    May2_LeaveCell
    May2_ENTERCell
    Close 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close 3
    
End Sub

Private Sub mes0_Click()
   Mes_imp = 0
   etiqueta
   Label1.BackColor = vbWhite
   Label1.FontUnderline = False
End Sub

Private Sub Mes1_Click()
   Mes_imp = 1
   etiqueta
End Sub

Private Sub Mes10_Click()
    Mes_imp = 10
    etiqueta
End Sub

Private Sub Mes11_Click()
   Mes_imp = 11
   etiqueta
End Sub

Private Sub Mes12_Click()
   Mes_imp = 12
   etiqueta
End Sub

Private Sub Mes2_Click()
    Mes_imp = 2
    etiqueta
End Sub

Private Sub Mes3_Click()
    Mes_imp = 3
    etiqueta
End Sub

Private Sub Mes4_Click()
   Mes_imp = 4
   etiqueta
End Sub

Private Sub Mes5_Click()
   Mes_imp = 5
   etiqueta
End Sub

Private Sub Mes6_Click()
   Mes_imp = 6
   etiqueta
End Sub

Private Sub Mes7_Click()
   Mes_imp = 7
   etiqueta
End Sub

Private Sub Mes8_Click()
  Mes_imp = 8
  etiqueta
End Sub

Private Sub Mes9_Click()
  Mes_imp = 9
  etiqueta
End Sub

Private Sub OrdAlf_Click()
    May2_LeaveCell
    colanti = May2.Col
    renati = May2.Row
    May2.Row = 1
    May2.Col = 1
    May2.RowSel = May2.Rows - 1
    May2.Sort = 1
    May2.Col = colanti
    May2.Row = renati
    May2_LeaveCell
    Rem May2_ENTERCell
    May2.SetFocus
  
End Sub

Private Sub OrdDep_Click()
Dim R1 As Integer, SdIto As Currency
Dim Y As Integer, Y1 As Integer

Open "ANTSDOS.PRN" For Output As #10
   For R1 = 1 To May2.Rows - 1
           May2.Row = R1: May2.Col = 2
          If IsNumeric(May2.Text) Then SdIto = May2.Text
          
          If SdIto <> 0 Then
                
                Rem AuxSub.Hide
                Load AuxForm
                Rem AuxForm.Hide
                Rem AuxForm.Show 1
                
                Load Depuracion
                
                For Y = 1 To Depuracion.Dep1.Rows - 1
                   Depuracion.Dep1.Row = Y
                   For Y1 = 0 To 4
                      Depuracion.Dep1.Col = Y1
                      Select Case Y1
                        Case 0
                        FcH = Depuracion.Dep1.Text
                        Case 1
                        NMb = Depuracion.Caption
                        Case 2
                        CptO = Depuracion.Dep1.Text
                        Case 3
                        If IsNumeric(Depuracion.Dep1.Text) Then
                                impte = Depuracion.Dep1.Text
                        End If
                        Case 4
                        If IsNumeric(Depuracion.Dep1.Text) Then
                                impte = Depuracion.Dep1.Text
                        End If
                        
                      End Select
                   Next Y1
                   Write #10, FcH, NMb, impte, CptO
                Next Y
                
               Unload AuxForm
               Unload Depuracion
        End If
        May2.TopRow = May2.TopRow + 1
   Next R1
   Close #10
End Sub

Private Sub OrdNum_Click()
    colanti = May2.Col
    renati = May2.Row
    May2.Row = 1
    May2.Col = 0
    May2.RowSel = May2.Rows - 1
    May2.Sort = 3
    May2.Col = colanti
    May2.Row = renati
    May2.SetFocus

End Sub
Sub impresora()
On Error GoTo error
    conteo = 0
    Titulos1
    rotulos1
    For r = 1 To May5.Rows - 1
      conteo = conteo + 1
      For l = 0 To 2
        
       If l > 0 Then
              
            Tabu_l = Tabu_l + May5.ColWidth(l - 1)
            Else
            Tabu_l = 1200
            Printer.CurrentX = Tabu_l
       End If
       If May5.TextMatrix(r, l) <> "" Then
          If l = 0 Then Printer.Print May5.TextMatrix(r, l);
          If l = 1 Then
                bala = May5.TextMatrix(r, l)
                valor$ = Format(bala, "#####0"): uso$ = "#####0"
                pone = 0: colocar pone, valor$, uso$:
                Printer.CurrentX = Printer.CurrentX + pone
                Printer.Print valor$;
          End If
                
          If l = 2 Then Printer.Print "   "; Format(May5.TextMatrix(r, l), "&&&");
       End If
      Next l
      For l = 3 To 5
         Tabu_l = Tabu_l + May5.ColWidth(l - 1)
         Printer.CurrentX = Tabu_l
         If May5.TextMatrix(r, l) <> "" Then
                bala = May5.TextMatrix(r, l)
                valor$ = Format(bala, "##,###,##0.00"): uso$ = "##,###,##0.00"
                pone = 0: colocar pone, valor$, uso$:
                Printer.CurrentX = Printer.CurrentX + pone
                Printer.Print valor$;
         End If
      Next l
      Printer.Print
      If conteo = (Tam_imp - 2) Then
                
                pie1
                mueve1
                conteo = 0
                
                Titulos1
                rotulos1
       End If
    Next r
    If conteo < (Tam_imp - 2) Then recorre1
Exit Sub

error:
    MsgBox (Err.Number & " " & Err.Description)

End Sub
 Sub recorre1()
     
    For i = conteo To (Tam_imp - 2)
        Printer.Print Chr(160)
                
    Next i
    pie1
    mueve1
    
 End Sub

Sub pie1()
  Printer.FontSize = 10
  Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 20), , BF
  Printer.Print
  Printer.Print Tab(20); AuxSub.May2.TextMatrix(AuxSub.May2.Row, 0); " "; AuxSub.May2.TextMatrix(AuxSub.May2.Row, 1)
  Printer.FontSize = 8

End Sub
 Sub mueve1()
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

Sub imprepantalla()
   If AuxSub.May2.TextMatrix(AuxSub.May2.Row, 3) <> "" Then
    Guia = AuxSub.May2.TextMatrix(AuxSub.May2.Row, 3)
    Nombre_Auxiliar = "AUXILIAR\AX" + LTrim(Str(Guia))

    May5.Clear
    May5.Row = 0
    May5.Row = 0
    May5.Col = 0: May5.ColWidth(0) = 900: May5.CellAlignment = 4: May5.Text = "Fecha"
    May5.Col = 1: May5.ColWidth(1) = 800:  May5.CellAlignment = 4: May5.Text = "Poliza"
    May5.Col = 2: May5.ColWidth(2) = 3200: May5.CellAlignment = 4: May5.Text = "Descripción"
    May5.Col = 3: May5.ColWidth(3) = 1200:  May5.CellAlignment = 4: May5.Text = "Debe"
    May5.Col = 4: May5.ColWidth(4) = 1200:  May5.CellAlignment = 4: May5.Text = "Haber"
    May5.Col = 5: May5.ColWidth(5) = 1200:  May5.CellAlignment = 4: May5.Text = "Saldo"
    May5.Col = 6: May5.ColWidth(6) = 80:  May5.CellAlignment = 4: May5.Text = ""
    May5.Rows = 1
    May5.Row = 0
    May5.Col = 1
    Open Nombre_Auxiliar For Random As 6 Len = Len(auxiliar)
    Fin_ax = LOF(6) / Len(auxiliar)
    If Mes_imp > 0 Then saldoini1
    If Fin_ax < 1 Then
         Rem MsgBox "El auxiliar solicitado no Contiene movimientos "
         Close
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
                 fechita = Mid(auxiliar.fech, 4, 2) + "/" + Left(auxiliar.fech, 2) + "/" + String(6 - Len(LTrim(Str(auxiliar.po))), "0") + LTrim(Str(auxiliar.po))
                 saldto = saldto + auxiliar.impo
                 May5.AddItem auxiliar.fech & Chr(9) & auxiliar.po & Chr(9) & _
                 (" " + auxiliar.re) & Chr(9) & Debe & Chr(9) & Haber & Chr(9) & _
                 Format(saldto, "###,###,##0.00") & Chr(9) & fechita
                 Else
                 Rem nada
             End If
         Next r
         If May5.Rows < 2 Then
           Rem MsgBox "El auxiliar solicitado no Contiene movimientos "
         Close
         Exit Sub
        
         End If
         Close 6
         Rem If May5.Rows = 2 And Saldo_Par = 0 Then Exit Sub

    End If
    colanti = May5.Col
    renati = May5.Row
    May5.Row = 1
    May5.Col = 6
    May5.RowSel = May5.Rows - 1
    May5.Sort = 5
    May5.Col = colanti
    May5.Row = renati
    calsaldo1
    impresora
   End If
End Sub
Private Sub Tam1_Click()
    Tam_imp = 20
    etiqueta
End Sub

Private Sub Tam2_Click()
    Tam_imp = 30
    etiqueta
End Sub

Private Sub Tam3_Click()
    Tam_imp = 60
    etiqueta
End Sub
