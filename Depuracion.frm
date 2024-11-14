VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Depuracion 
   Caption         =   "Depuracion "
   ClientHeight    =   3165
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Dep1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.Menu DepImp 
      Caption         =   "&Imprimir"
      Begin VB.Menu ImpImp 
         Caption         =   "&Impresion"
         Shortcut        =   ^I
      End
      Begin VB.Menu ImpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu ImpGar 
         Caption         =   "&Generar Archivo"
      End
      Begin VB.Menu ImpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu ImpSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu DepEd 
      Caption         =   "&Edicion"
      Begin VB.Menu DepSel 
         Caption         =   "&Seleccionar Todo"
         Shortcut        =   ^S
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu DepCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "Depuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumFact As Integer, CapLin, Max_Tab As Long
Sub DepImpte()
   Dim VaLorIniCial As Currency, ValorFinal As Currency, FinalRecorrido As Integer
    FinalRecorrido = Dep1.Rows - 1
    r = 0
    Do Until r = FinalRecorrido
      r = r + 1
      If IsNumeric(Dep1.TextMatrix(r, 3)) Then
            VaLorIniCial = Dep1.TextMatrix(r, 3)
        For w = 1 To FinalRecorrido
            If IsNumeric(Dep1.TextMatrix(w, 4)) Then
               ValorFinal = Dep1.TextMatrix(w, 4)
                If (ValorFinal + VaLorIniCial) = 0 Then
                    Dep1.TextMatrix(r, 3) = ""
                    Dep1.TextMatrix(w, 4) = ""
                    Exit For
                End If
            End If
        Next w
      End If
    Loop
    CanCeLacion
    recorrer
    AlReves
    PorNombre 0
    SuMaTT
End Sub
Sub CanCeLacion()
    final = Dep1.Rows - 1
    r = 0
    Rem If r = final Then
     Rem NADA
    Rem Else
    Do Until r = final
      r = r + 1
      Rem If r = final Then Exit Do
      If (Dep1.TextMatrix(r, 3) = "") And (Dep1.TextMatrix(r, 4) = "") Then
                If final <= 1 Then Exit Do
                Dep1.RemoveItem r
                r = r - 1: final = final - 1
                
      End If
    Loop
    
    Rem End If
End Sub
Sub AlReves()
    Dim VaLorIniCial As Currency, ValorFinal As Currency, FinalRecorrido As Integer
    Dim Rec(500)
    FinalRecorrido = Dep1.Rows - 1
    r = 0
    Do Until r = FinalRecorrido
      r = r + 1
      If IsNumeric(Dep1.TextMatrix(r, 4)) Then
            VaLorIniCial = Dep1.TextMatrix(r, 4)
            ValorFinal = 0: F = 0
        For w = 1 To FinalRecorrido
            If IsNumeric(Dep1.TextMatrix(w, 3)) Then
                F = F + 1
                Rec(F) = w
                ValorFinal = ValorFinal + Dep1.TextMatrix(w, 3)
                If (VaLorIniCial + ValorFinal) = 0 Then
                   Dep1.Row = r: Dep1.Col = 4: Dep1.CellBackColor = vbYellow
                   Dep1.TextMatrix(r, 4) = ""
                  For w1 = 1 To F
                    Dep1.Row = Rec(w1): Dep1.Col = 3: Dep1.CellBackColor = vbCyan
                    Dep1.TextMatrix(Rec(w1), 3) = ""
                  Next w1
                 Exit For
               End If
            End If
        Next w
      End If
    Loop
    CanCeLacion
End Sub

Sub recorrer()
    Dim VaLorIniCial As Currency, ValorFinal As Currency, FinalRecorrido As Integer
    Dim Rec(60)
    FinalRecorrido = Dep1.Rows - 1
    
    r = 0
    Do Until r = FinalRecorrido
      r = r + 1
      If IsNumeric(Dep1.TextMatrix(r, 3)) Then
            VaLorIniCial = Dep1.TextMatrix(r, 3)
            ValorFinal = 0: F = 0
        For w = 1 To FinalRecorrido
            If IsNumeric(Dep1.TextMatrix(w, 4)) Then
                F = F + 1
                Rec(F) = w
                ValorFinal = ValorFinal + Dep1.TextMatrix(w, 4)
                If (VaLorIniCial + ValorFinal) = 0 Then
                   Dep1.TextMatrix(r, 3) = ""
                  For w1 = 1 To F
                    Dep1.TextMatrix(Rec(w1), 4) = ""
                  Next w1
                 Exit For
               End If
            End If
        Next w
      End If
    Loop
    CanCeLacion
End Sub


Private Sub Dep1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Depuracion
    End If
End Sub
Sub TTitulos()
On Error GoTo errorHandler:

   Rem ******************** AuxForm.Titulos ****************************
    Max_Tab = 1200
    For l = 0 To 5: Max_Tab = Max_Tab + Dep1.ColWidth(l): Next l
    Printer.FontSize = 10
    
    centrar pone, Max_Tab, Trim(Datos.D1)
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
     
    Rem ******************** TITULOS *************************************
 

Exit Sub
errorHandler:

    MsgBox (Err.Number & Err.Description)
    
End Sub
Sub RRotulos()
    Printer.FontBold = True
    Tabu_l = 1200
    For l = 0 To 5
        Printer.CurrentX = Tabu_l
        centrar pone, Dep1.ColWidth(l), Dep1.TextMatrix(0, l)
        Printer.CurrentX = Printer.CurrentX + pone
        Printer.Print Dep1.TextMatrix(0, l);
        Tabu_l = Tabu_l + Dep1.ColWidth(l)
    Next l
    Printer.Print
    Printer.Line (1200, Printer.CurrentY)-(Max_Tab, Printer.CurrentY + 30), , BF
    Printer.FontBold = False
    
End Sub

Private Sub DepCop_Click()
    Dim Temporal1
    Clipboard.Clear
    Temporal1 = Temporal1 + AuxForm.Caption & Chr(13) & (10)
    Rem Dep1.RowSel = Dep1.Rows - 1
    Rem Dep1.ColSel = Dep1.Cols - 2
    For i = 0 To Dep1.RowSel
      For F = 0 To Dep1.ColSel
            Temporal1 = Temporal1 + Dep1.TextMatrix(i, F) & Chr(9)
      Next F
      Clipboard.SetText Temporal1 & Chr(13) & Chr(10)
   Next i
   difer = Dep1.RowSel - Dep1.Row

End Sub

Private Sub DepImp_Click()
    Tam_imp = 60
    conteo = 0
    
    TTitulos
    RRotulos
    For r = 1 To Dep1.Rows - 1
      conteo = conteo + 1
      For l = 0 To 2
        
       If l > 0 Then
              
            Tabu_l = Tabu_l + Dep1.ColWidth(l - 1)
            Else
            Tabu_l = 1200
            Printer.CurrentX = Tabu_l
       End If
       If Dep1.TextMatrix(r, l) <> "" Then
          If l = 0 Then Printer.Print Dep1.TextMatrix(r, l);
          If l = 1 Then
                bala = Dep1.TextMatrix(r, l)
                valor$ = Format(bala, "#####0"): uso$ = "#####0"
                pone = 0: colocar pone, valor$, uso$:
                Printer.CurrentX = Printer.CurrentX + pone
                Printer.Print valor$;
          End If
                
          If l = 2 Then Printer.Print "   "; Format(Dep1.TextMatrix(r, l), "&&&");
       End If
      Next l
      For l = 3 To 5
         Tabu_l = Tabu_l + Dep1.ColWidth(l - 1)
         Printer.CurrentX = Tabu_l
         If Dep1.TextMatrix(r, l) <> "" Then
                bala = Dep1.TextMatrix(r, l)
                valor$ = Format(bala, "##,###,##0.00"): uso$ = "##,###,##0.00"
                pone = 0: colocar pone, valor$, uso$:
                Printer.CurrentX = Printer.CurrentX + pone
                Printer.Print valor$;
         End If
      Next l
      Printer.Print
      
      If Printer.CurrentY >= (Printer.Height - 3000) Then
                Printer.CurrentY = Printer.Height - 2800
                PPie
                Printer.NewPage
                TTitulos
                RRotulos
       End If
    Next r
    If Printer.CurrentY <= (Printer.Height - 3000) Then
           Printer.CurrentY = (Printer.Height - 2800)
           PPie
           Else
           PPie
    End If
    Printer.EndDoc
    Rem conteonum = 0
    Rem f impresionactivada = 1 Then impresionactivada = 0
End Sub

Sub PPie()
   
   Printer.FontSize = 10
   Printer.Line (1200, Printer.CurrentY)-((Max_Tab + 1200), Printer.CurrentY + 20), , BF
   Printer.Print
   Printer.Print Tab(20); Depuracion.Caption
   Printer.FontSize = 8
End Sub

Private Sub DepSel_Click()
   Dep1.Col = 0: Dep1.Row = 0
   Dep1.RowSel = Dep1.Rows - 1
   Dep1.ColSel = Dep1.Cols - 2
      

End Sub

Private Sub Form_Load()
Dim FeCHa As String, mm1 As Integer, Ao1 As Integer
   Dep1.Clear
   Dep1.Cols = 7: Dep1.FixedCols = 0
   z1 = "#,##0.00;(#,##0.00)"
   Iniciar
   Dep1.Rows = 1
   Depuracion.Caption = Mid(AuxForm.Caption, 11, 32)
   For r = 1 To AuxForm.May3.Rows - 1
         FecA = AuxForm.May3.TextMatrix(r, 0)
         mm1 = Val(Mid(FecA, 4, 2)):  Ao1 = Val(Mid(FecA, 7, 2) + 2000)
        
         FeCHa = Left(mm(mm1), 3) + "-" + Trim(Str(Ao1))
         
         CapLin = FeCHa & Chr(9) & AuxForm.May3.TextMatrix(r, 1) & Chr(9) & _
                  AuxForm.May3.TextMatrix(r, 2) & Chr(9) & AuxForm.May3.TextMatrix(r, 3) & Chr(9) & _
                  AuxForm.May3.TextMatrix(r, 4) & Chr(9) & AuxForm.May3.TextMatrix(r, 5) & Chr(9) & _
                  AuxForm.May3.TextMatrix(r, 6)
                  Dep1.AddItem CapLin
         
   Next r
   If Dep1.Rows > 2 Then DepImpte
End Sub
Sub SuMaTT()
Dim SSTT As Currency
  SSTT = 0
  For r = 1 To Dep1.Rows - 1
     If IsNumeric(Dep1.TextMatrix(r, 3)) Then SSTT = SSTT + Dep1.TextMatrix(r, 3)
     If IsNumeric(Dep1.TextMatrix(r, 4)) Then SSTT = SSTT + Dep1.TextMatrix(r, 4)
     Dep1.TextMatrix(r, 5) = Format(SSTT, z1)
  Next r
End Sub
Sub Iniciar()
    Dep1.Row = 0
    Dep1.Col = 0: Dep1.ColWidth(0) = 900: Dep1.CellAlignment = 4: Dep1.Text = "Fecha"
    Dep1.Col = 1: Dep1.ColWidth(1) = 800:  Dep1.CellAlignment = 4: Dep1.Text = "Poliza"
    Dep1.Col = 2: Dep1.ColWidth(2) = 3200: Dep1.CellAlignment = 4: Dep1.Text = "Descripción"
    Dep1.Col = 3: Dep1.ColWidth(3) = 1200:  Dep1.CellAlignment = 4: Dep1.Text = "Debe"
    Dep1.Col = 4: Dep1.ColWidth(4) = 1200:  Dep1.CellAlignment = 4: Dep1.Text = "Haber"
    Dep1.Col = 5: Dep1.ColWidth(5) = 1200:  Dep1.CellAlignment = 4: Dep1.Text = "Saldo"
    Dep1.Col = 6: Dep1.ColWidth(6) = 80: Dep1.CellAlignment = 4: Dep1.Text = ""
    Depuracion.Width = 9000
End Sub

Private Sub Form_Resize()
   If Depuracion.WindowState <> 1 Then
       Dep1.Width = Depuracion.Width - 400
       Dep1.Height = Depuracion.Height - 1200
   End If
   
End Sub
Sub PorNombre(ReFerencia As Long)
   Dim Refer As String, DaTo, SuMaTo As Currency, Refer1 As String
   Dim ReFerenCia1 As Long, VaLorIniCial As Currency, ValorFinal As Currency
   final = Dep1.Rows - 1
  For r = 1 To final
        Refer = "": VaLorIniCial = 0
        For w3 = 1 To Len(Dep1.TextMatrix(r, 2))
            DaTo = Mid(Dep1.TextMatrix(r, 2), w3, 1)
            If IsNumeric(DaTo) Then
                Refer = Refer + DaTo
                If w3 < Len(Dep1.TextMatrix(r, 2)) Then
                   If IsNumeric(Mid(Dep1.TextMatrix(r, 2), w3 + 1, 1)) Then
                        Rem nada
                        Else
                        Exit For
                   End If
                End If
            End If
        Next w3
        If IsNumeric(Refer) Then
              ReFerencia = Val(Refer)
              If IsNumeric(Dep1.TextMatrix(r, 3)) Then VaLorIniCial = Dep1.TextMatrix(r, 3)
              If IsNumeric(Dep1.TextMatrix(r, 4)) Then VaLorIniCial = Dep1.TextMatrix(r, 4)
              SuMaTo = 0: ValorFinal = 0
              For w1 = 1 To final
              If w1 = r Then GoTo SiGue1
                For w2 = 1 To Len(Dep1.TextMatrix(w1, 2))
                        dato1 = Mid(Dep1.TextMatrix(w1, 2), w2, 1)
                        '***** Si el 1er. Número coincide entonces compara la cadena
                        If dato1 = Left(Refer, 1) Then
                             Refer1 = Mid(Dep1.TextMatrix(w1, 2), w2, Len(Refer))
                             w2 = w2 - 1 + Len(Refer)
                             If Refer1 = Refer Then
                                    Exit For
                                    Else
                                    Refer1 = ""
                             End If
                              
                        End If
                Next w2
                If Refer1 = "" Then Refer1 = "0"
                ReFerenCia1 = Refer1
                If ReFerenCia1 = ReFerencia Then
                   If IsNumeric(Dep1.TextMatrix(w1, 3)) Then ValorFinal = Dep1.TextMatrix(w1, 3)
                   If IsNumeric(Dep1.TextMatrix(w1, 4)) Then ValorFinal = Dep1.TextMatrix(w1, 4)
                   If VaLorIniCial < 0 And ValorFinal < 0 Then
                         VaLorIniCial = 0: ValorFinal = 0
                         Exit For
                   End If
                   If VaLorIniCial > 0 And ValorFinal > 0 Then
                         VaLorIniCial = 0: ValorFinal = 0
                         Exit For
                   End If
                   SuMaTo = VaLorIniCial + ValorFinal
                   Select Case SuMaTo
                      Case 0
                         Dep1.TextMatrix(r, 3) = "": Dep1.TextMatrix(r, 4) = ""
                         Dep1.TextMatrix(w1, 3) = "": Dep1.TextMatrix(w1, 4) = ""
                         Exit For
                      Case Is > 0
                         If VaLorIniCial > 0 Then
                            
                            VaLorIniCial = 0: ValorFinal = 0
                            If (Len(Dep1.TextMatrix(w1, 2)) - w3) > 0 Then
                                transfer = Right(Dep1.TextMatrix(w1, 2), Len(Dep1.TextMatrix(w1, 2)) - w3)
                                Else
                                transfer = ""
                            End If
                            transfer = Left(Dep1.TextMatrix(w1, 2), w2 - 1) + transfer
                            Rem Mid(transfer, (w2 + 1 - Len(Refer1)), Len(Refer1)) = Chr(Caracter) + Right(Refer1, Len(Refer1) - 1)
                            Rem Mid(transfer, (w2 + 1 - Len(Refer1)), Len(Refer1)) = String(Len(Refer1), "x")
                            Dep1.TextMatrix(w1, 2) = transfer: transfer = ""
                            Dep1.TextMatrix(r, 3) = Format(SuMaTo, z1)
                            Dep1.TextMatrix(w1, 4) = ""
                            Refer1 = ""
                            Exit For
                            Else
                            
                            If VaLorIniCial < 0 Then
                                VaLorIniCial = 0: ValorFinal = 0
                                If (Len(Dep1.TextMatrix(r, 2)) - w3) > 0 Then
                                    transfer = Right(Dep1.TextMatrix(r, 2), Len(Dep1.TextMatrix(r, 2)) - w3)
                                    Else
                                    transfer = ""
                                End If
                                transfer = Left(Dep1.TextMatrix(r, 2), w3 - Len(Refer)) + transfer
                                Rem Mid(transfer, (w3 + 1 - Len(Refer)), Len(Refer)) = Chr(Caracter) + Right(Refer, Len(Refer) - 1)
                                Dep1.TextMatrix(r, 2) = transfer: transfer = ""
                                Dep1.TextMatrix(w1, 3) = Format(SuMaTo, z1)
                                Dep1.TextMatrix(r, 4) = ""
                            Refer = ""
                            Exit For
                            End If
                         End If
                      Case Is < 0
                         If VaLorIniCial > 0 Then
                            
                            VaLorIniCial = 0: ValorFinal = 0
                            If (Len(Dep1.TextMatrix(w1, 2)) - w2) > 0 Then
                                transfer = Right(Dep1.TextMatrix(w1, 2), Len(Dep1.TextMatrix(w1, 2)) - w2)
                                Else
                                transfer = ""
                            End If
                            transfer = Left(Dep1.TextMatrix(w1, 2), w2 - 1) + transfer
                            Dep1.TextMatrix(w1, 2) = transfer: transfer = ""
                            Dep1.TextMatrix(w1, 4) = Format(SuMaTo, z1)
                            Dep1.TextMatrix(r, 3) = ""
                            Refer1 = ""
                            Exit For
                            Else
                            If ValorFinal > 0 Then
                                VaLorIniCial = 0: ValorFinal = 0
                                If (Len(Dep1.TextMatrix(r, 2)) - w3) > 0 Then
                                    transfer = Right(Dep1.TextMatrix(r, 2), Len(Dep1.TextMatrix(r, 2)) - w3)
                                    Else
                                    transfer = ""
                                End If
                                transfer = Left(Dep1.TextMatrix(r, 2), w3 - Len(Refer)) + transfer
                                Dep1.TextMatrix(r, 2) = transfer: transfer = ""
                                Dep1.TextMatrix(r, 4) = Format(SuMaTo, z1)
                                Dep1.TextMatrix(w1, 3) = ""
                                Refer = ""
                            Exit For
                            End If
                         End If
                    End Select
                End If
                Refer1 = ""
             Rem End If
SiGue1:
              Next w1
        End If
        Refer = 0
 Next r
 CanCeLacion
End Sub

Private Sub ImpSal_Click()
   Close
   Unload Depuracion
End Sub
