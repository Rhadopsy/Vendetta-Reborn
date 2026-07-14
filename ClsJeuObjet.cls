VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJeuObjet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Vendetta - Standalone Role Playing game, simulate a MMORPG with AI
' Copyright (C) 2003-2007  S. HARLAUT (sharlaut@netcourrier.com)
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
Option Explicit

Public Numero As Integer
Public Nombre As Integer
Public Nom As String
Public Description As String
Public Indesctructible As Boolean
Public Insepararable As Boolean
Public Categorie As Integer
Public Competence As Integer
Private PrixRessources() As Integer
Public PrixVente As Integer
Public BonusVie As Integer
Public BonusEnergie As Integer
Public BonusMagie As Integer
Public BonusMoral As Integer
Public BonusAttaque As Integer
Public BonusPorteeAttaque As Integer
Public EffetAttaque As Integer
Public SurfaceAttaque As Integer
Public BonusDefense As Integer
Public BonusDeplacement As Integer
Private BonusCompetencesSpeciales() As Integer
Private BonusCompetencesRessources() As Integer
Private BonusCompetencesServices() As Integer
Private BonusCompetencesObjets() As Integer
Public BonusArmure As Integer
Public BonusMaxRessources As Integer
Public BonusVitesseRegenerationVie As Single
Public BonusVitesseRegenerationEnergie As Single
Public BonusVitesseRegenerationMagie As Single
Public BonusVitesseRegenerationMoral As Single
Public PerteVie As Double
Public PerteEnergie As Double
Public PerteMagie As Double
Public PerteMoral As Double
Public PerteArgent As Long
Public PerteRessourceIndice As Integer
Public PerteRessourceQuantite As Single
Public ApparencePerso As Integer
Public ChangerApparencePerso As Boolean
Public Utilisable As Boolean
Public UtilisationUnique As Boolean
Public UtilisationMort As Boolean
Public Experience As Double
Public Invocation As Integer
Public Quantite As Single
Public GainVie As Double
Public GainEnergie As Double
Public GainMagie As Double
Public GainMoral As Double
Public GainArgent As Single
Private GainRessource() As Single
Public Resurrection As Long
Public SelfResurrection As Boolean

Public PerteSiAttaque As Boolean

Public Sub Init(ByVal Parametres As ClsJeuParam)
    Numero = -1
    PerteRessourceIndice = -1
    ReDim PrixRessources(Parametres.NombreRessources - 1)
    ReDim BonusCompetencesSpeciales(Parametres.NombreCompetencesSpeciales - 1)
    ReDim BonusCompetencesRessources(Parametres.NombreRessources - 1)
    ReDim BonusCompetencesServices(Parametres.NombreServices - 2)
    ReDim BonusCompetencesObjets(Parametres.NombreCompetencesObjets - 1)
    ReDim GainRessource(Parametres.NombreRessources - 1)
End Sub

Public Sub Charger(ByVal Indice As Integer, _
                   ByRef Parametres As ClsJeuParam)
    Dim i As Integer
    Numero = Indice
    
    With Parametres
    
    Nom = .Objet_Nom(Indice)
    Nombre = 1
    Description = .Objet_Description(Indice)
    Indesctructible = .Objet_Indestructible(Indice)
    Insepararable = .Objet_Inseparable(Indice)
    Categorie = .Objet_Categorie(Indice)
    Competence = .Objet_Competence(Indice)
    For i = 0 To UBound(PrixRessources())
        PrixRessources(i) = .Objet_PrixRessources(Indice, i)
    Next i
    PrixVente = .Objet_PrixVente(Indice)
    BonusVie = .Objet_BonusVie(Indice)
    BonusEnergie = .Objet_BonusEnergie(Indice)
    BonusMagie = .Objet_BonusMagie(Indice)
    BonusMoral = .Objet_BonusMoral(Indice)
    BonusPorteeAttaque = .Objet_BonusPorteeAttaque(Indice)
    EffetAttaque = .Objet_EffetAttaque(Indice)
    SurfaceAttaque = .Objet_SurfaceAttaque(Indice)
    BonusAttaque = .Objet_BonusAttaque(Indice)
    BonusDefense = .Objet_BonusDefense(Indice)
    BonusDeplacement = .Objet_BonusDeplacement(Indice)
    For i = 0 To UBound(BonusCompetencesSpeciales())
        BonusCompetencesSpeciales(i) = .Objet_BonusCompetencesSpeciales(Indice, i)
    Next i
    For i = 0 To UBound(BonusCompetencesRessources())
        BonusCompetencesRessources(i) = .Objet_BonusCompetencesRessources(Indice, i)
    Next i
    For i = 0 To UBound(BonusCompetencesServices())
        BonusCompetencesServices(i) = .Objet_BonusCompetencesServices(Indice, i)
    Next i
    For i = 0 To UBound(BonusCompetencesObjets())
        BonusCompetencesObjets(i) = .Objet_BonusCompetencesObjets(Indice, i)
    Next i
    BonusArmure = .Objet_BonusArmure(Indice)
    BonusMaxRessources = .Objet_BonusMaxRessources(Indice)
    BonusVitesseRegenerationVie = .Objet_BonusVitesseRegenerationVie(Indice)
    BonusVitesseRegenerationEnergie = .Objet_BonusVitesseRegenerationEnergie(Indice)
    BonusVitesseRegenerationMagie = .Objet_BonusVitesseRegenerationMagie(Indice)
    BonusVitesseRegenerationMoral = .Objet_BonusVitesseRegenerationMoral(Indice)
    PerteVie = .Objet_PerteVie(Indice)
    PerteEnergie = .Objet_PerteEnergie(Indice)
    PerteMagie = .Objet_PerteMagie(Indice)
    PerteMoral = .Objet_PerteMoral(Indice)
    PerteArgent = .Objet_PerteArgent(Indice)
    PerteRessourceIndice = .Objet_PerteRessourceIndice(Indice)
    PerteRessourceQuantite = .Objet_PerteRessourceQuantite(Indice)
    ApparencePerso = .Objet_ApparencePerso(Indice)
    ChangerApparencePerso = .Objet_ChangerApparencePerso(Indice)
    Utilisable = .Objet_Utilisable(Indice)
    UtilisationUnique = .Objet_UtilisationUnique(Indice)
    UtilisationMort = .Objet_UtilisationMort(Indice)
    Experience = .Objet_Experience(Indice)
    Invocation = .Objet_Invocation(Indice)
    Quantite = .Objet_Quantite(Indice)
    GainVie = .Objet_GainVie(Indice)
    GainEnergie = .Objet_GainEnergie(Indice)
    GainMagie = .Objet_GainMagie(Indice)
    GainMoral = .Objet_GainMoral(Indice)
    GainArgent = .Objet_GainArgent(Indice)
    For i = 0 To UBound(GainRessource())
        GainRessource(i) = .Objet_GainRessource(Indice, i)
    Next i
    Resurrection = .Objet_Resurrection(Indice)
    SelfResurrection = .Objet_SelfResurrection(Indice)
    
    PerteSiAttaque = PerteVie > 0 Or _
                     PerteEnergie > 0 Or _
                     PerteMagie > 0 Or _
                     PerteMoral > 0 Or _
                     UtilisationUnique Or _
                     PerteArgent > 0 Or _
                     PerteRessourceIndice >= 0 And _
                     PerteRessourceQuantite > 0
    End With
    
End Sub

Public Sub Decharger()
    Dim i As Long
    If Nombre > 1 Then
        Nombre = Nombre - 1
    Else
        Numero = -1
        Nombre = 0
        Nom = ""
        Description = ""
        Indesctructible = False
        Insepararable = False
        Categorie = 0
        Competence = 0
        For i = 0 To UBound(PrixRessources())
            PrixRessources(i) = 0
        Next i
        PrixVente = 0
        BonusVie = 0
        BonusEnergie = 0
        BonusMagie = 0
        BonusMoral = 0
        BonusAttaque = 0
        BonusPorteeAttaque = 0
        EffetAttaque = 0
        SurfaceAttaque = 0
        BonusDefense = 0
        BonusDeplacement = 0
        For i = 0 To UBound(BonusCompetencesSpeciales())
            BonusCompetencesSpeciales(i) = 0
        Next i
        For i = 0 To UBound(BonusCompetencesRessources())
            BonusCompetencesRessources(i) = 0
        Next i
        For i = 0 To UBound(BonusCompetencesServices())
            BonusCompetencesServices(i) = 0
        Next i
        For i = 0 To UBound(BonusCompetencesObjets())
            BonusCompetencesObjets(i) = 0
        Next i
        BonusArmure = 0
        BonusMaxRessources = 0
        BonusVitesseRegenerationVie = 0
        BonusVitesseRegenerationEnergie = 0
        BonusVitesseRegenerationMagie = 0
        BonusVitesseRegenerationMoral = 0
        PerteVie = 0
        PerteEnergie = 0
        PerteMagie = 0
        PerteMoral = 0
        PerteArgent = 0
        PerteRessourceIndice = -1
        PerteRessourceQuantite = 0
        ApparencePerso = 0
        ChangerApparencePerso = False
        Utilisable = False
        UtilisationUnique = False
        Experience = 0
        Invocation = 0
        Quantite = 0
        GainVie = 0
        GainEnergie = 0
        GainMagie = 0
        GainMoral = 0
        GainArgent = 0
        For i = 0 To UBound(GainRessource())
            GainRessource(i) = 0
        Next i
        Resurrection = 0
        
        PerteSiAttaque = False
    End If
End Sub

Public Property Get Actif() As Boolean
    'Actif = Not Nom = ""
    Actif = Numero > -1
End Property
Public Property Let Actif(ByVal Valeur As Boolean)
    If Not Valeur Then
        Nom = ""
    Else
        Nom = "Temp"
    End If
End Property

Public Property Get Param_PrixRessources(ByVal Indice As Integer) As Integer
    Param_PrixRessources = PrixRessources(Indice)
End Property

'Bonus des compétences.
Public Property Get Param_BonusCompetencesSpeciales(ByVal Indice As Integer) As Integer
    Param_BonusCompetencesSpeciales = BonusCompetencesSpeciales(Indice)
End Property
Public Property Get Param_BonusCompetencesRessources(ByVal Indice As Integer) As Integer
    Param_BonusCompetencesRessources = BonusCompetencesRessources(Indice)
End Property
Public Property Get Param_BonusCompetencesServices(ByVal Indice As Integer) As Integer
    Param_BonusCompetencesServices = BonusCompetencesServices(Indice)
End Property
Public Property Get Param_BonusCompetencesObjets(ByVal Indice As Integer) As Integer
    Param_BonusCompetencesObjets = BonusCompetencesObjets(Indice)
End Property
Public Property Get Param_GainRessource(ByVal Indice As Integer) As Single
    Param_GainRessource = GainRessource(Indice)
End Property
