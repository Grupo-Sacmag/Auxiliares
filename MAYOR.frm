VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mayor 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Cuentas de Mayor"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6630
   Icon            =   "MAYOR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FijarDir1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid May1 
      Height          =   4575
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      BackColorBkg    =   14737632
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Menu MayAr 
      Caption         =   "&Archivo"
      Begin VB.Menu MayAct 
         Caption         =   "&Actualización"
      End
      Begin VB.Menu Arsep2 
         Caption         =   "-"
      End
      Begin VB.Menu ArCamb 
         Caption         =   "&Cambio de Subdirectorio"
      End
      Begin VB.Menu ArSep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArVer 
         Caption         =   "&Verificar"
      End
      Begin VB.Menu ArSep3 
         Caption         =   "-"
      End
      Begin VB.Menu ArSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu AuEd 
      Caption         =   "&Edicion"
      Begin VB.Menu AuEdCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "Mayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Apertura()
    MiArchivo = Dir("Auxiliar", vbDirectory)
    Rem ******   Verifica que Exista el archivo de auxiliares *****
    
    If MiArchivo = "" Then
       respuesta = MsgBox("No existe el Subdirectorio de Auxiliares desea crearlo ", vbYesNo, "Auxiliares ")
       If respuesta = vbYes Then
              
              MkDir "AUXILIAR"
              
       End If
    End If
    
End Sub

Private Sub ArVer_Click()
    Verificar.Show 1
End Sub

Private Sub AuEdCop_Click()
On Error GoTo error:
    Dim Temporal1 As String
    Dim i As Integer, F As Integer

    ' Limpiar el portapapeles
    Clipboard.Clear

    ' Construir el texto para copiar al portapapeles
    For i = 0 To May1.Rows - 1
        For F = 0 To May1.Cols - 1
            Temporal1 = Temporal1 & May1.TextMatrix(i, F) & vbTab
        Next F
        Temporal1 = Temporal1 & vbCrLf
    Next i

    ' Establecer el texto en el portapapeles
    On Error Resume Next
    Clipboard.SetText Temporal1
    On Error GoTo 0

    ' Verificar si se pudo establecer el texto en el portapapeles
    If Err.Number <> 0 Then
        MsgBox "No se pudo copiar al portapapeles. Error: " & Err.Description, vbExclamation
    End If

Exit Sub
error:

    MsgBox (Err.Number & " _ " & Err.Description)
    Clipboard.Clear

End Sub

Function IsArrayInitialized(arr() As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = IsArray(arr) And Not IsError(LBound(arr, 1))
    On Error GoTo 0
End Function



Private Sub Form_Unload(Cancel As Integer)
   Close
   End
End Sub

Private Sub May1_DblClick()
   Clipboard.Clear
End Sub

Private Sub May1_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        AuxSub.Show 1
      End If
End Sub

Private Sub MAY1_LeaveCell()
If May1.Rows > 1 Then
 If May1.Col > 0 And May1.Row > 0 Then
         May1.CellBackColor = vbWhite
 End If
 End If
End Sub
    
Private Sub MAY1_ENTERCell()
 If May1.Row >= 1 Then
  May1.CellBackColor = vbYellow
 End If
End Sub

Sub inicio()
   Rem MsgBox "ESTOY EMPEZANDO "
On Error GoTo saltalo
    Open (Ruta_Acceso_Contr + "\Gcont.Arr") For Random As 3 Len = Len(SCont)
        Get 3, 1, SCont
    If SCont.guarda <= " " Then
        ChDir "C:\"
        
        Else
        
        If Left(SCont.guarda, 1) <> "C" Then
                    ChDrive Left(SCont.guarda, 1)
        End If

        ChDir SCont.guarda
    End If

    

    Close #1
    Archivo = "DATOS"
    Open Archivo For Random As #1 Len = Len(Datos)
    cm = LOF(1) / Len(Datos)
    Get 1, 1, Datos
saltalo:
     Rem MsgBox "YA PASE"
     
    Close #3
End Sub
Sub sigpaso()
   If cm < 1 Then
        MsgBox "No Existe Datos de Contabilidad Cambie subdirectorio "
        Close
        Rem Kill "DATOS"
        Else
        Get 1, 1, Datos
        Label1.BackColor = vbYellow
        
        Label1.Caption = RTrim(Datos.D1) + " Auxiliares " + Datos.a_o
        Open "CATMAY" For Random As 2 Len = Len(CATMAY)
        cm = LOF(2) / Len(CATMAY)
        For r = 1 To cm: Get 2, r, CATMAY
           If Val(CATMAY.B4) > 0 Then
               May1.AddItem Format(CATMAY.B1, "#####") & Chr(9) & _
                            CATMAY.B2 & Chr(9) & _
                            Format(CATMAY.B3, "###,##0.00") & Chr(9) & _
                            Val(CATMAY.B4) & Chr(9) & _
                            Val(CATMAY.B5)
           End If
        Next r
       If May1.Rows > 3 Then Apertura
       Close 2
    End If

End Sub

Private Sub ArCamb_Click()
  Close #1
MientraS = ""
On Err GoTo Errhandler
    MIDIR = CurDir
    MIDIR = RTrim(MIDIR)
    If Right(MIDIR, 1) = "\" Then
        MIDIR = Left(MIDIR, Len(MIDIR) - 1)
    End If
    FijarDir1.InitDir = MIDIR
    FijarDir1.Flags = cdlOFNHideReadOnly
    FijarDir1.Filter = "Archivos de Auxiliares(Dat*.*)|Dat*.*"
    FijarDir1.ShowOpen
    If FijarDir1.FileName <> "" Then
        For i = 1 To Len(FijarDir1.FileName)
                If Mid(FijarDir1.FileName, i, 1) = "\" Then tope = i
        Next i
        MientraS = Mid(FijarDir1.FileName, 1, tope)
        
        ChDir MientraS
        Close 3
        Open (Ruta_Acceso_Contr + "\Gcont.Arr") For Random As 3 Len = Len(SCont)
        SCont.guarda = MientraS
        Put 3, 1, SCont
        Close 3
        cm = 0
        May1.Clear
        May1.Rows = 1
        inicio
        sigpaso
    End If
Errhandler:

   
End Sub

Private Sub ArSal_Click()
   Close: End
End Sub

Private Sub Form_Load()
    MiArchivo = Dir("C:\GconTA", vbDirectory)
    Rem ******   Verifica que Exista de controles *****
    If MiArchivo = "" Then
       Rem respuesta = MsgBox("No existe el Subdirectorio Controles desea crearlo ", vbYesNo, "Contabilidad General ")
       Rem If respuesta = vbYes Then
              MkDir "C:\GconTA"
       Rem End If
    End If
    Ruta_Acceso_Contr = "C:\GconTA"
    inicio
    Ruta_Acceso = App.Path
    z1 = "##,###,##0.00"
    mm(1) = "ENERO": mm(2) = "FEBRERO": mm(3) = "MARZO": mm(4) = "ABRIL"
    mm(5) = "MAYO": mm(6) = "JUNIO": mm(7) = "JULIO": mm(8) = "AGOSTO"
    mm(9) = "SEPTIEMBRE": mm(10) = "OCTUBRE": mm(11) = "NOVIEMBRE": mm(12) = "DICIEMBRE"
    dd(1) = 31: dd(2) = 28: dd(3) = 31: dd(4) = 30
    dd(5) = 31: dd(6) = 30: dd(7) = 31: dd(8) = 31
    dd(9) = 30: dd(10) = 31: dd(11) = 30: dd(12) = 31
    May1.Row = 0
    May1.Row = 0
    May1.Col = 0: May1.ColWidth(0) = 600: May1.CellAlignment = 4: May1.Text = "Cuenta"
    May1.Col = 1: May1.ColWidth(1) = 3350:  May1.CellAlignment = 4: May1.Text = "Nombre"
    May1.Col = 2: May1.ColWidth(2) = 1200:  May1.CellAlignment = 4: May1.Text = "Importe"
    May1.Col = 3: May1.ColWidth(3) = 80:  May1.CellAlignment = 4: May1.Text = ""
    May1.Col = 4: May1.ColWidth(4) = 80:  May1.CellAlignment = 4: May1.Text = ""
    May1.Rows = 1
    May1.Row = 0
    May1.Col = 1
    sigpaso
    If May1.Rows > 1 Then
        May1.Col = 1: May1.Row = 1
        MAY1_LeaveCell
        MAY1_ENTERCell
  
     End If
End Sub

Private Sub MayAct_Click()
        
    CommonDialog1.CancelError = True
    On Error GoTo Errguarda
    MIDIR = CurDir
    MIDIR = RTrim(MIDIR)
    If Right(MIDIR, 1) = "\" Then
        MIDIR = Left(MIDIR, Len(MIDIR) - 1)
    End If
    CommonDialog1.InitDir = MIDIR
    Get 1, 1, Datos
    If Datos.No_arch = "" Then
        Archivo = InputBox("Teclee el nombre del archivo de datos ")
        If Len(Archivo) > 6 Then
            MsgBox "Nombre no valido", vbCritical
            Exit Sub
            Else
            Datos.No_arch = Archivo
            Put 1, 1, Datos
        End If
    End If

    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNNoChangeDir
    CommonDialog1.FileName = MIDIR + "\" + RTrim(Datos.No_arch) + "*.*"
    anuncio = ""
    anuncio = "Archivos de Operaciones (" + RTrim(Datos.No_arch) + "*.*)|" + RTrim(Datos.No_arch) + "*.*|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = anuncio
    
    Rem CommonDialog1.InitDir = SCont.guarda
    
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DefaultExt = " *"
    If CommonDialog1.FileName = "" Then
            CommonDialog1.FileName = RTrim(Datos.No_arch) + "*.*"
    End If
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = RTrim(Datos.No_arch) + "*.*" Then
            CommonDialog1.FileName = RTrim(Datos.No_arch) + "01"
    End If
    donde = Len(RTrim(Datos.No_arch)) + 1
    
    Arch_Act = RTrim(CommonDialog1.FileTitle)
    
    For i = 1 To Len(Arch_Act)
       If (Mid(Arch_Act, i, 1) >= Chr(48)) And (Mid(Arch_Act, i, 1)) <= Chr(57) Then
                       
            Mes_Act = Val(Right(RTrim(Arch_Act), 2))
            If (Mid(Arch_Act, i + 2, 1)) = "." Then
                 MsgBox "Archivo no Valido ", vbCritical
                 Exit Sub
            End If
            
            Exit For
       End If
    Next i
    Open "AUXILIAR\CONTR.AUX" For Random As 5 Len = Len(Veri_ficar)
    f_m = LOF(5) / Len(Veri_ficar)
    Open Arch_Act For Random As 4 Len = Len(oper)
    em = LOF(4) / Len(oper)
    Get 5, Mes_Act, Veri_ficar
    rgtro = Val(Veri_ficar.record) + 1
    If rgtro >= em Then
       MsgBox RTrim(Arch_Act) + " ya esta Actualizado ", vbCritical
       Close 4, 5
       Else
       Close 4, 5
       ultimo.num = 1
       Actual.Show 1, Me
    End If
Errguarda:
  Exit Sub

End Sub

