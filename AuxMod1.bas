Attribute VB_Name = "AuxMod1"
 Type CAT_MA
    B1 As String * 6
    B2 As String * 32
    B3 As String * 16
    B4 As String * 5
    B5 As String * 5
End Type
Type CAT_AX
    C1 As String * 6
    C2 As String * 32
    C3 As String * 16
    C4 As String * 5
    C5 As String * 5
End Type
Type DAT_OS
    D1 As String * 64
    D2 As String * 60
    D3 As String * 45
    No_arch As String * 15
    a_o As String * 5
    others1  As String * 25
    UltimaPol As String * 5
    UltimoReg As String * 5
    others As String * 12
End Type

Type oper_aciones
    cta As String * 6
    descr As String * 30
    fe As String * 2
    impte As String * 16
    identi As String * 1
    real As String * 9
End Type

Type cheques
         num As Integer
         beneficiario As String * 40
         importe As Currency
         Clave As String * 1
         numreal As Integer
         Refer As Integer
         conta As Integer
End Type
Type tra_cta
       num As Integer
       nombre As String * 32
       donde As Integer
       inicia As Integer
       termina As Integer
       Clave As String * 1
 End Type
Type tra_Scta
       num As Integer
       nombre As String * 32
       donde As Integer
       Refer As Integer
       Clave As String * 1
 End Type
 Type ult
     num As Long
     ubi As Integer
     renglon As Long
     texto As String
     poliza As Integer
     impresion As Integer
     TipoCap As Integer
 End Type
 Type sc
    guarda As String * 64
End Type
Type aux
    fech As String * 8
    po As Integer
    re As String * 30
    impo As Double
    sal As Double
 End Type
 Type Veri
    record As String * 8
    poliza As String * 8
 End Type
 Public Veri_ficar As Veri
 Public Mes_Act As Integer
 Public auxiliar As aux
 Public SCont As sc
 Public ultimo As ult, ultimo1 As ult
 Public trcta As tra_cta
 Public trscta As tra_Scta
 Public cheque As cheques
 Public Datos As DAT_OS
 Public CATMAY As CAT_MA
 Public CATAUX As CAT_AX, Ruta_Acceso, Ruta_Acceso_Contr
 Public oper As oper_aciones
 Public z1 As String, z2 As String, valcelant, fin_oper As Long
 Public ltotal As Long
 Public Arch_Act As String * 60
 Public cm As Integer, dm As Integer, em As Integer, qm As Integer
 Public mm(15) As String * 20, dd(15) As Integer, m_m As Integer, dia As Integer
 Public Tam_imp As Integer, Mes_imp As Integer, impresionactivada As Integer
 
 Sub colocar(ancho2, valor$, us_o As String)
     ancho2 = 0
     ancho = Printer.TextWidth(valor$)
     ancho1 = Printer.TextWidth(us_o)
     ancho2 = ancho1 - ancho
     Rem Printer.CurrentX = Printer.currex + ancho2
 End Sub
Sub derecha(ancho2, ltotal, cadena As String)
    ancho2 = 0
    ancho2 = (ltotal - Printer.TextWidth(cadena))
End Sub
Sub centrar(ancho2, ltotal, cadena As String)
    ancho2 = 0
    ancho2 = (ltotal - Printer.TextWidth(cadena)) / 2
    
    
End Sub
