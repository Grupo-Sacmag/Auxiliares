VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Verificar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2775
   Icon            =   "Verificar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Verificar.frx":0442
   ScaleHeight     =   3960
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid VerArc1 
      CausesValidation=   0   'False
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   14
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483644
      BackColorBkg    =   12632256
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "Verificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pic


Private Sub Form_Load()
     Verificar.Caption = RTrim(Datos.D1)
     Rem Pic = "C:\Archivos de Programa\Captura\open.bmp"
     Rem Pic = Ruta_Acceso + "\open.bmp"
     VerArc1.Row = 0
     VerArc1.Col = 0: VerArc1.ColWidth(0) = 280:  VerArc1.CellAlignment = 9: VerArc1.Text = ""
     VerArc1.Col = 1: VerArc1.ColWidth(2) = 500:  VerArc1.CellAlignment = 1: VerArc1.Text = "Archivo"
     VerArc1.Col = 2: VerArc1.ColWidth(3) = 500:  VerArc1.CellAlignment = 4: VerArc1.Text = "Reg."
     VerArc1.Col = 3: VerArc1.ColWidth(3) = 500:  VerArc1.CellAlignment = 4: VerArc1.Text = "Pol."
     VerArc1.Width = 2280
     VerArc1.Height = (300 * 13)
     Open "AUXILIAR\CONTR.AUX" For Random As 5 Len = Len(Veri_ficar)
     f_m = LOF(5) / Len(Veri_ficar)
     ren = 1
     For r = 1 To f_m: Get 5, r, Veri_ficar
          
          If Val(Veri_ficar.record) > 0 Then
             VerArc1.Row = ren: VerArc1.Col = 0
             Set VerArc1.CellPicture = LoadPicture(Pic)
             rr = Str(r)
             If Len(LTrim(rr)) < 2 Then
                rr = "0" + LTrim(rr)
                Else
                rr = LTrim(rr)
             End If
             VerArc1.Col = 1: VerArc1.Text = RTrim(Datos.No_arch) + rr
             VerArc1.Col = 2: VerArc1.Text = Format(Val(Veri_ficar.record), "####0")
             VerArc1.Col = 3: VerArc1.Text = Format(Val(Veri_ficar.poliza), "####0")
             ren = ren + 1
          End If
     Next r
     Close 5
End Sub
