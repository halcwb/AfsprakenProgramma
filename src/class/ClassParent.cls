VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Energy As Double
Private m_Eiwit As Double
Private m_KH As Double
Private m_Vet As Double
Private m_Na As Double
Private m_K As Double
Private m_Ca As Double
Private m_P As Double
Private m_Mg As Double
Private m_Fe As Double
Private m_VitD As Double
Private m_Cl As Double
Private m_Product As String

Public Property Get Name() As String

    Name = m_Name

End Property

Public Property Let Name(ByVal strName As String)

    m_Name = strName

End Property

Public Property Get Product() As String

    Product = m_Product

End Property

Public Property Let Product(ByVal strProduct As String)

    m_Product = strProduct

End Property

Public Property Get Energy() As Double

    Energy = m_Energy

End Property

Public Property Let Energy(ByVal dblEnergy As Double)

    m_Energy = dblEnergy

End Property

Public Property Get Eiwit() As Double

    Eiwit = m_Eiwit

End Property

Public Property Let Eiwit(ByVal dblEiwit As Double)

    m_Eiwit = dblEiwit

End Property

Public Property Get KH() As Double

    KH = m_KH

End Property

Public Property Let KH(ByVal dblKH As Double)

    m_KH = dblKH

End Property

Public Property Get Vet() As Double

    Vet = m_Vet

End Property

Public Property Let Vet(ByVal dblVet As Double)

    m_Vet = dblVet

End Property

Public Property Get Na() As Double

    Na = m_Na

End Property

Public Property Let Na(ByVal dblNa As Double)

    m_Na = dblNa

End Property

Public Property Get K() As Double

    K = m_K

End Property

Public Property Let K(ByVal dblK As Double)

    m_K = dblK

End Property

Public Property Get Ca() As Double

    Ca = m_Ca

End Property

Public Property Let Ca(ByVal dblCa As Double)

    m_Ca = dblCa

End Property

Public Property Get P() As Double

    P = m_P

End Property

Public Property Let P(ByVal dblP As Double)

    m_P = dblP

End Property

Public Property Get Mg() As Double

    Mg = m_Mg

End Property

Public Property Let Mg(ByVal dblMg As Double)

    m_Mg = dblMg

End Property

Public Property Get Fe() As Double

    Fe = m_Fe

End Property

Public Property Let Fe(ByVal dblFe As Double)

    m_Fe = dblFe

End Property

Public Property Get VitD() As Double

    VitD = m_VitD

End Property

Public Property Let VitD(ByVal dblVitD As Double)

    m_VitD = dblVitD

End Property

Public Property Get Cl() As Double

    Cl = m_Cl

End Property

Public Property Let Cl(ByVal dblCl As Double)

    m_Cl = dblCl

End Property

