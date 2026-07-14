VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAffObjets"
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

Const Chemin1 = "Objets\"
Const Chemin2 = "Ressources\"

Const LargeurImage As Long = 32
Const HauteurImage As Long = 32

Const Transparent = True

Private Sprites() As ClsAffSprite

Public Sub Charger(ByRef dd As DirectDraw7, _
                   ByVal Parametres As ClsJeuParam, _
                   Optional ByVal Ressources As Boolean)
    Dim i As Integer
    If Ressources Then
        ReDim Sprites(Parametres.NombreRessources - 1)
        For i = 0 To UBound(Sprites())
            Set Sprites(i) = New ClsAffSprite
            If Parametres.Ressources_CheminImage(i) = "" Then
                Sprites(i).Charger dd, Chemin2 & Parametres.Ressources_Image(i)
            Else
                Sprites(i).Charger dd, Parametres.Ressources_CheminImage(i) & Parametres.Ressources_Image(i), True
            End If
            Sprites(i).Selection 0, HauteurImage, 0, LargeurImage
        Next i
    Else
        ReDim Sprites(Parametres.NombreObjets - 1)
        For i = 0 To UBound(Sprites())
            Set Sprites(i) = New ClsAffSprite
            If Parametres.Objet_CheminImage(i) = "" Then
                Sprites(i).Charger dd, Chemin1 & Parametres.Objet_Image(i)
            Else
                Sprites(i).Charger dd, Parametres.Objet_CheminImage(i) & Parametres.Objet_Image(i), True
            End If
            Sprites(i).Selection 0, HauteurImage, 0, LargeurImage
        Next i
    End If
End Sub

Public Sub Afficher(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal TypeObjet As Long, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal ValeurTransparence As Single, _
                    Optional ByVal IndiceApparence As Integer)
    'On charge la bonne image en fonction des paramètres.
    'Sprite.Selection 0, _
                     HauteurImage, _
                     TypeObjet * LargeurImage, _
                     (TypeObjet + 1) * LargeurImage
    Sprites(TypeObjet).Selection HauteurImage * IndiceApparence, _
                                 HauteurImage * (IndiceApparence + 1), _
                                 0, _
                                 LargeurImage
    'On affiche l'objet sur l'écran.
    If ValeurTransparence > 0 Then
        Sprites(TypeObjet).Afficher X, Y, SurfBack, Transparent, Translucide, , , ValeurTransparence
    Else
        Sprites(TypeObjet).Afficher X, Y, SurfBack, Transparent, Translucide
    End If
End Sub

Public Sub Afficher2(ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal TypeObjet As Long, _
                    ByRef SurfBack As DirectDrawSurface7, _
                    Optional ByVal Translucide As Boolean, _
                    Optional ByVal IndiceApparence As Integer, _
                    Optional ByVal Z As Single = 1)
    'On charge la bonne image en fonction des paramètres.
    Sprites(TypeObjet).Selection HauteurImage * IndiceApparence, _
                                 HauteurImage * (IndiceApparence + 1), _
                                 0, _
                                 LargeurImage
    'On affiche l'objet sur l'écran.
    Sprites(TypeObjet).Afficher2 X, Y, SurfBack, Transparent, Z
End Sub

Public Property Get Largeur() As Long
    Largeur = LargeurImage
End Property

Public Property Get Hauteur() As Long
    Hauteur = HauteurImage
End Property

