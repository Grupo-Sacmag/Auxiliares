VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Actual 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "Actualización"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   Icon            =   "Actual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar Actando 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Actual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num_ro As Currency, finaux As Long, s_do As Currency
Dim fec_ha, p_ol As Long, rgtro As Long, r As Long, mensaje

Private Sub Command1_Click()
  Form_Initialize
  Form_Load
End Sub

Private Sub Command1_GotFocus()
  Form_Initialize
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Form_Initialize()
 
 Label1.FontBold = False
 mensaje = RTrim(Datos.D1) + " " + RTrim(mm(Mes_Act)) + " " + Datos.a_o
 Label1.Caption = mensaje
 
End Sub
Sub Act_Auxiliar()
    If Command1.Value = True Then
       Open "AUXILIAR\CONTR.AUX" For Random As 5 Len = Len(Veri_ficar)
        f_m = LOF(5) / Len(Veri_ficar)
        Open Arch_Act For Random As 4 Len = Len(oper)
        em = LOF(4) / Len(oper)
        Get 5, Mes_Act, Veri_ficar
        rgtro = Val(Veri_ficar.record) + 1
        Actando.Scrolling = ccScrollingSmooth
        If rgtro > em Then
           Close 5, 4
           Exit Sub
        End If
        Actando.Max = em
        Actando.Min = 1
        Actando.Value = rgtro
               
       For r = rgtro To em: Get 4, r, oper
            Actual.Refresh
            Actando.Value = r
            
            Select Case oper.identi
              Case "A"
                  m_me = LTrim(Str(Mes_Act))
                  If Len(m_me) = 1 Then m_me = "0" + m_me
                  If m_me = 13 Then m_me = "01"
                  fec_ha = oper.fe + "/" + m_me + "/" + Right(RTrim(Datos.a_o), 2)
                  p_ol = Val(oper.cta)
                  
              Case "C"
                  num_ro = Val(oper.cta)
                  Label1.Caption = mensaje & Chr(13) & "Actualizado : Registro " + Format(r, "####0")
                  Open "AUXILIAR\AX" + LTrim(Str(num_ro)) For Random As 6 Len = Len(auxiliar)
                  finax = LOF(6) / Len(auxiliar)
                  If finax > 0 Then
                     Get 6, finax, auxiliar
                     s_do = auxiliar.sal
                     finax = finax + 1
                     Else
                     s_do = 0
                     finax = finax + 1
                  End If
                  auxiliar.fech = fec_ha
                  auxiliar.re = oper.descr
                  auxiliar.impo = Val(oper.impte)
                  auxiliar.sal = s_do + Val(oper.impte)
                  auxiliar.po = p_ol
                  Put 6, finax, auxiliar
                  
                  Close 6
              End Select
                  
       Next r
       Label1.FontBold = True
       Label1.Caption = mensaje & Chr(13) & "Actualizado : Registro " + Str(em) + " Poliza : " + Str(p_ol)
       Get 5, Mes_Act, Veri_ficar
       Veri_ficar.record = em
       Veri_ficar.poliza = p_ol
       Put 5, Mes_Act, Veri_ficar
       Close 4, 5
      
   End If
End Sub
Sub Act_Ana()
If Command1.Value = True Then
      Open "CatAux" For Random As 6 Len = Len(CATAUX)
      Open "CatMay" For Random As 7 Len = Len(CATMAY)
      Actando.Scrolling = ccScrollingSmooth
      Actando.Max = dm
      Actando.Min = 0
      If ultimo.num = 2 Then
            Actando.Value = ultimo.ubi
            Else
            Actando.Value = dm
      End If
      For r = ultimo.ubi To dm: Get 12, r, oper
                Actual.Refresh
           If ultimo.num = 2 Then
                Actando.Value = r
                Else
                Actando.Value = Actando.Max - r
           End If
            Select Case oper.identi
              Case "A"
              Label1.Caption = mensaje & Chr(13) & "Actualizado : Registro " + Format(r, "####0")
              Datos.UltimaPol = LTrim(oper.cta)
              Case "B"
                  Rem Actualizacion cuenta
                  Label1.Caption = mensaje & Chr(13) & "Actualizado : Registro " + Format(r, "####0")
                  num_ro = Val(oper.real)
                  Get 7, num_ro, CATMAY
                  s_do = Val(CATMAY.B3)
                  If ultimo.num = 2 Then
                        s_do = s_do + Val(oper.impte)
                        CATMAY.B3 = Str(s_do)
                        Put 7, num_ro, CATMAY
                        Else
                        s_do = s_do - Val(oper.impte)
                        CATMAY.B3 = Str(s_do)
                        Put 7, num_ro, CATMAY
                  End If
                Case "C"
                  Rem Actualizacion Subcuenta
                  Label1.Caption = mensaje & Chr(13) & "Actualizado : Registro " + Format(r, "####0")
                  num_ro = Val(oper.cta)
                  Get 6, num_ro, CATAUX
                  s_do = Val(CATAUX.C3)
                  If ultimo.num = 2 Then
                        s_do = s_do + Val(oper.impte)
                        CATAUX.C3 = Str(s_do)
                        Put 6, num_ro, CATAUX
                        Else
                        s_do = s_do - Val(oper.impte)
                        CATAUX.C3 = Str(s_do)
                        Put 6, num_ro, CATAUX
                   End If
              End Select
                  
       Next r
       If ultimo.num = 2 Then
              Label1.FontBold = True
              
              Label1.Caption = mensaje & Chr(13) & "Actualizado : Registro " + Str(dm) + " Poliza : " + LTrim(Datos.UltimaPol)
              Datos.UltimoReg = Str(dm)
              Put 1, Mes_Act, Datos
              Else
              Label1.FontBold = True
              Label1.Caption = mensaje & Chr(13) & "Archivo DesActualizado "
              Datos.UltimoReg = Str(0)
              Datos.UltimaPol = "0"
              Put 1, Mes_Act, Datos
        End If
       Close
End If
End Sub
Private Sub Form_Load()
    Select Case ultimo.num
        Case 1
        Act_Auxiliar
        Case 2
        Act_Ana
        Case 3
        Act_Ana
    End Select
End Sub


Private Sub Timer1_Timer()
      Timer1.Interval = 900
End Sub
