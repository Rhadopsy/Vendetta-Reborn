VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe utilisant la réference Microsoft XML V3.0<BR>
' Cette classe vous permet de lire et d&#39;écrire dans un fichier XML.<BR>
' Elle utilise egalement la syntaxe de commentaire <a href="http://sourceforge.net/projects/vbdox" target="_blank">VBDox</a>(freeware),
' pour la génération d&#39;une documentation technique.<br>
' Cette classe est libre de tout droit, mais si vous apportiez des modifications, je vous serais reconnaissant de m'en faire part afin de mettre ce code source à jour si le besoin s'en faisait sentir.
'
'@require   Visual Basic 6<BR>La réfencerence Microsoft XML,V3.0 (msxml3.dll) dans votre projet
'@todo      Travailler sur des fichiers XML.<br>Vous pouvez lire,ecrire,créer votre fichier XML. La recherche de noeud dans le fichier est possible également via son nom et ses attributs.
'@bug       Aucun pour le moment (sauf l&#39;orthographe...)
'@history   Ne trouvant rien sur le Net pour travailler avec VB6 et XML de facon simple,<br> j&#39;ai créé ma classe.
'@author    Thomas Darmengeat
'@version   <BR>Version 1.0
'@date      2004/06/16

'Nouvelles fonctions ajoutées par Syl pour Vendetta.

''
'Variable contenant le fichier de facon globale
Private fXML As New DOMDocument30
''
'Variable contenant le noeud principale sert de base pour les recherches dans le fichier
Private root As IXMLDOMElement
''
'Variable de création d&#39;un nouvel attribut
Private newAtt As IXMLDOMAttribute
''
'Variable de mappage de sous noeuds
Private namedNodeMap As IXMLDOMNamedNodeMap
''
'Variable de création de noeuds
Private newNode As IXMLDOMNode
''
'Variable listant tout les sous noeuds
Private listNode As IXMLDOMNodeList
''
'Variable de gestion des erreurs XML (pas encore utilisé)
Private erreurXML As IXMLDOMParseError
''
'Variable contenant le nom du fichier charger dans l&#39;objet.<BR>
'Elle est initialisé à l&#39;ouverture du fichier.
Private fichierXML As String
''
'Variable qui indique si le fichier a été modifié et doit être enregistré.
'Par défaut : non
Private modifierXML As Boolean
''
'Variable qui indique si les attributs recherchés n&#39;existe pas <br> on considère que le noeud est valide ou non.<br>
'Par défaut : non
Private validerNoeud As Boolean
''
'Variable contenant une description de l&#39;erreur.
Private msgErreur As String
''
' Methode appeler à la création de l&#39;objet.
'
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Private Sub Class_Initialize()
 validerNoeud = False
 modifierXML = False
 fichierXML = ""
 msgErreur = ""
End Sub
''
' Methode appeler à la destruction de l&#39;objet.
'
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Private Sub Class_Terminate()
fXML = Nothing
root = Nothing
newAtt = Nothing
Set namedNodeMap = Nothing
Set newNode = Nothing
Set listNode = Nothing
Set erreurXML = Nothing
fichierXML = ""
modifierXML = False
validerNoeud = False
msgErreur = ""
End Sub
''
' Propriété de la classe contenant le nom du fichier XML charger.
'
'@return    Le nom du fichier XML charger dans l&#39;objet
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Public Property Get fichier() As String
    fichier = fichierXML
End Property
''
' Propriété de la classe indiquant si le fichier XML a été modifier.
'
'@return    Le nom du fichier XML charger dans l&#39;objet
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Public Property Get modifier() As Boolean
    modifier = modifierXML
End Property
''
' Renvoit si un noeud sans attributs est prise en compte lors d&#39;action avec demande d'attributs.
'
'
'@return    Boolean indiquant si le noeud sera prise en charge
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/18
Public Property Get noeudSansAttributs() As Boolean
    noeudSansAttributs = validerNoeud
End Property
''
' Définit si un noeud sans attributs est prise en compte lors d&#39;action avec demande d'attributs.
'
'@param validation Boolean indiquant si le noeud sera prise en charge
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/18
Public Property Let noeudSansAttributs(validation As Boolean)
    validerNoeud = validation
End Property
''
' Renvoi le message d&#39;erreur.
'
'
'@return    Le message d&#39;erreur
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/18
Public Property Get Erreur() As String
    Erreur = msgErreur
End Property
''
' Création d&#39;un document XML
'
'@param nomXML Nom du fichier XML à créer avec son extension (ex: toto.xml)
'@param chemin Arborescence du fichier à créer(ex: c:\testXML\, par défaut: répertoire du programme)
'@param cle_racine Nom de la clé principale du fichier (par défaut: root)
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Public Sub creerXML(nomXML As String, Optional chemin, Optional cle_racine)
    If IsMissing(chemin) Then chemin = App.Path
    If IsMissing(cle_racine) Then cle_racine = "root"
    fich = FreeFile
    Open chemin & "\" & nomXML For Output As #fich
    Print #fich, "<?xml version='1.0'?>"
    Print #fich, "<" & cle_racine & ">" & vbCrLf & "</" & cle_racine & ">"
    Close fich
End Sub
''
' Fonction de chargement d&#39;un fichier XML.<br>En cas d&#39;erreur lire le message dans la propriété erreur
'
'@param nomXML Nom du fichier XML avec son extension (ex: toto.xml)
'@param chemin Arborescence du fichier à ouvrir(ex: c:\testXML\)
'
'@return    Boolean indiquant si l&#39;ouverture du fichier c&#39;est bien déroulée
'@todo      Si il y a un problème à l&#39;ouverture, le message d&#39;erreur est dans la propriété d'erreur
'@history   Voilà maintenant l&#39;erreur d&#39;ouverture est gérée !
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Public Function ouvrirXML(nomXML As String, Optional chemin) As Boolean
    If IsMissing(chemin) Then chemin = App.Path
    fXML.async = False
    ouvrirXML = fXML.Load(chemin & "\" & nomXML)
    If ouvrirXML Then
        fichierXML = chemin & "\" & nomXML
    Else
        Set erreurXML = fXML.parseError
        msgErreur = erreurXML.reason
    End If
'MsgBox fXML.xml '** debogage **'
End Function

Public Function ouvrirXML2(nomXML As String) As Boolean
    fXML.async = False
    ouvrirXML2 = fXML.loadXML(nomXML)
    If Not ouvrirXML2 Then
        Set erreurXML = fXML.parseError
        msgErreur = erreurXML.reason
    End If
'MsgBox fXML.xml '** debogage **'
End Function

''
' Fonction qui simule la fermeture du fichier XML.<br>
' Si le fichier est valide dans son formatage,
' il est sauvegardé avant d&#39;en ouvrir un nouveau.
'
'
'@return    Boolean indiquant si le fichier est valide et a été enregistré.
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Public Function fermerXML() As Boolean
    fermerXML = sauverXML
    fichierXML = ""
End Function
''
' Sauvegarde le fichier XML chargé dans l&#39;objet.<br>
'Verifie si le fichier est ecrit correctement et le sauvegarde sinon la sauvegarde echoue.<br>
'Cependant on peut forcer la sauvegarde.
'
'@param force Boolean indiquant si la sauvegarde doit etre effectuer alors que la validation n&#39;est pas juste.
'
'@return    Boolean indiquant si la sauvegarde a été effectuer
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/16
Public Function sauverXML(Optional force As Boolean = False) As Boolean
    If force Then
        fXML.Save fichierXML
        sauverXML = True
        modifierXML = False
    Else
        If fXML.Validate Then
            fXML.Save fichierXML
            sauverXML = True
            modifierXML = False
        Else
            sauverXML = False
            msgErreur = "Votre fichier XML n'est pas valide, les modifcations ne peuvent donc être enregistrées."
        End If
    End If
End Function
''
' Ecrit un nouveau noeud dans le fichier XML ou on le souhaite dans le fichier XML
'
'@param nvNoeud Le noeud que l&#39;on souhaite créer
'@param chemin L&#39;arborescence dans le fichier XML ou écrire ce nouveau noeud
'@param atts Paramètre regroupant des attributs du noeud choisit afin de sélectionner certains noeud en particulier ou même 1 seul.
'@history Au début cette méthode écrivait dans le premier élément trouvé.
'@author    Thomas Darmengeat
'@version   1.1
'@date  2004/06/21
Public Sub ecrireNoeudXML(nvNoeud As String, chemin As String, Optional atts As String = "")
    Dim noeud2 As IXMLDOMNode
    Set root = fXML.documentElement
    fXML.SetProperty "SelectionLanguage", "XPath"
    Set newNode = fXML.createNode(NODE_ELEMENT, nvNoeud, "")
    If chemin <> "" Then
        Set listNode = root.getElementsByTagName(chemin)
        For i = 0 To listNode.length - 1
            Set noeud2 = listNode.Item(i)
            If verifierAttribut(noeud2, atts) Or atts = "" Then
                noeud2.appendChild newNode
            End If
        Next
    Else
        root.appendChild newNode
    End If
    modifierXML = True
'fXML.save fichierXML '** debogage **'
'MsgBox fXML.xml '** debogage **'
End Sub
''
' Ecrit un attribut à un noeud du fichier XML ou on le souhaite.
'
'@param noeud Chemin du noeud et noeud sous la forme (../../..)
'@param attribut Nom de l&#39;attribut à acrire
'@param valeur Valeur de l&#39;attribut
'@param atts Paramètre regroupant des attributs du noeud choisit afin de sélectionner certains noeud en particulier ou même 1 seul.
'@bug       Rajoute l&#39;attribut à tout les noeuds trouvé en fonctions du nom et des attributs en parametres
'@todo      Ecrit l&#39;attribut dans tout les noeuds correspondant au nom et aux attributs passé en parametres
'@history   Au début cette fonction n&#39;écrivait que dans le premier noeud trouvé.
'@author    Thomas Darmengeat
'@version   1.1
'@date      2004/06/16
Public Sub ecrireAttXML(noeud As String, attribut As String, valeur As String, Optional atts As String = "")
    fXML.SetProperty "SelectionLanguage", "XPath"
    Set root = fXML.documentElement
    Set newAtt = fXML.createAttribute(attribut)
    newAtt.nodeValue = valeur
    Set listNode = root.getElementsByTagName(noeud)
    For Each newNode In listNode
        If atts = "" Or verifierAttribut(newNode, atts) Then
            Set namedNodeMap = newNode.Attributes
            namedNodeMap.setNamedItem newAtt
            modifierXML = True
        End If
    Next
'fXML.save fichierXML '** debogage **'
'MsgBox fXML.xml '** debogage **'
End Sub
''
' Ecrit des données pour le noeud choisi, n&#39;importe ou dans le fichier XML.<BR>
' Ecrit sous la forme <...>XXX<\...><br>
' Il faudra préciser le plus possible pour un noeud qui peut être présent plusieurs fois,<br> car pour chaque occurence trouvé les données seront écrites.<br>
' Les attributs ont donc un important si vous travailler sur un noeud multi-occurence.
'@param noeud Noeud et arborescence du noeud dans le fichier
'@param data Donnée à ecrire dans le noeud
'@param atts Paramètre optionnel il sert à n&#39;écrire que dans le noeud choisit si il existe plusieurs occurences de ce noeud.
'@bug       Ecrit dans tout les élements trouvés correspondant aux attributs.
'@todo      Ecrit des données dans tout les noeuds correspondant au nom et aux attributs passé en parametres
'@history   Au début cette fonction n&#39;écrivait que dans le premier noeud trouvé.
'@author    Thomas Darmengeat
'@version   1.1
'@date      2004/06/16
Public Sub ecrireDataXML(noeud As String, data As String, Optional atts As String = "")
    Dim datasec As IXMLDOMCDATASection
    fXML.SetProperty "SelectionLanguage", "XPath"
    Set root = fXML.documentElement
    Set datasec = fXML.createCDATASection(data)
    Set listNode = root.getElementsByTagName(noeud)
    For Each newNode In listNode
        If atts = "" Or verifierAttribut(newNode, atts) Then
            newNode.appendChild datasec
            modifierXML = True
        End If
    Next
'fXML.save fichierXML '** debogage **'
MsgBox fXML.xml '** debogage **'
End Sub
''
' Lit l&#39;ensemble ou une parite du fichier XML.<BR>
' On ne peut obtenir que la première partie du fichier si on ne veut qu&#39;une partie seulement du fichier;
'
'@param nbLigne Nombre de ligne souhaitée en retour, si pas de valeur alors on retourne tout le fichier.
'
'@return    Une chaine contenant l&#39;ensemble du texte retourné.
'@author    Thomas Darmengeat
'@version   1.0
'@date      2004/06/17
Public Function lirexml(Optional nbLigne As Integer = 0) As String
    If nbLigne = 0 Then
        lirexml = fXML.xml
    Else
        Dim tmp As String, tbtmp() As String
        tmp = fXML.xml
        tbtmp = Split(tmp, vbCrLf)
        For i = 0 To nbLigne - 1
            lirexml = lirexml & tbtmp(i) & vbCrLf
        Next
    End If
End Function
''
' Récupere les attributs d&#39;un noeud.<br>
' Retourne tout les noeuds qui portent le nom passé en argument.
' Pour un noeud situé en profondeur il faut donner son arborescence.<br>
' Le tableau sera organisé comme cela:<br>
' <BLOCKQUOTE>noeud:nomatt=valeuratt,...;noeud:nomatt=valeuratt,...;... (un ; correspond a une nouvelle ligne du tableau)
'
'@param nomNoeud Nom du noeud à recuperer.
'@param atts Paramètre regroupant des attributs du noeud choisit afin de sélectionner certains noeud en particulier ou même 1 seul.
'
'@return    Tableau contenant les noeuds,les attributs et la valeurs des attributs
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/17
Public Function lireNoeudXML(nomNoeud As String, Optional atts As String = "") As String()
    Dim noeud As IXMLDOMNode
    Dim msg As String, retour() As String
    ReDim retour(0)
    fXML.SetProperty "SelectionLanguage", "XPath"
    Set root = fXML.documentElement
    Set listNode = root.getElementsByTagName(nomNoeud)
    For Each noeud In listNode
        'on stock comme cela noeud:nomatt=valeuratt,...;noeud:nomatt=valeuratt,...;
        msg = ""
        If retour(UBound(retour)) <> "" Then ReDim Preserve retour(UBound(retour) + 1)
        If atts = "" Or verifierAttribut(noeud, atts) Then
            msg = noeud.baseName & ":"
            For i = 0 To noeud.Attributes.length - 1
                Set newAtt = noeud.Attributes.Item(i)
                nomatt = newAtt.Name
                valatt = newAtt.Value
                msg = msg & nomatt & "=" & valatt
                If i <> noeud.Attributes.length - 1 Then msg = msg & ","
            Next
            If msg <> "" Then retour(UBound(retour)) = msg
        End If
    Next
    lireNoeudXML = retour
End Function

''
' Fonction qui lit les données d&#39;un noeud.<BR>
' Si plusieurs occurences du nom de noeud desiré existe le tableau contiendra toute les valeurs ayant été lu.
'
'@param nomNoeud Nom du noeud dont on souhaite avoir les données.
'@param atts Paramètre regroupant des attributs du noeud choisit afin de sélectionner certains noeud en particulier ou même 1 seul.
'
'@return    Tableau contenant les données lu.
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/17
Public Function lireDataXML(nomNoeud As String, Optional atts As String = "") As String()
    Dim noeud As IXMLDOMNode
    Dim retour() As String
    Dim datasec As IXMLDOMCDATASection
    ReDim retour(0)
    fXML.SetProperty "SelectionLanguage", "XPath"
    Set root = fXML.documentElement
    Set listNode = root.getElementsByTagName(nomNoeud)
    For Each noeud In listNode
        If atts = "" Or verifierAttribut(noeud, atts) Then
            If retour(UBound(retour)) <> "" Then ReDim Preserve retour(UBound(retour) + 1)
            If noeud.hasChildNodes Then
                For i = 0 To noeud.childNodes.length - 1
                    If noeud.childNodes.Item(i).nodeType = NODE_CDATA_SECTION Then
                       Set datasec = noeud.childNodes.Item(i)
                       retour(UBound(retour)) = datasec.data
                    End If
                Next
             End If
        End If
    Next
    lireDataXML = retour
End Function

Public Function lireDataXML2(nomNoeud As String, Optional atts As String = "") As String()
    Dim noeud As IXMLDOMNode
    Dim retour() As String
    Dim datasec As IXMLDOMCDATASection
    ReDim retour(0)
    fXML.SetProperty "SelectionLanguage", "XPath"
    Set root = fXML.documentElement
    Set listNode = root.getElementsByTagName(nomNoeud)
    For Each noeud In listNode
        If atts = "" Or verifierAttribut(noeud, "") Then
            If retour(UBound(retour)) <> "" Then ReDim Preserve retour(UBound(retour) + 1)
            If noeud.hasChildNodes Then
                For i = 0 To noeud.childNodes.length - 1
                    If noeud.childNodes.Item(i).baseName = atts Then
                       'Set datasec = noeud.childNodes.Item(i)
                       retour(UBound(retour)) = noeud.childNodes.Item(i).Text
                    End If
                Next
             End If
        End If
    Next
    lireDataXML2 = retour
End Function

''
' Fonction de recherche d&#39;un noeud.<br>
' La recherche de base sur le nom du noeud et les attributs demander.<br>
' Si pour les attributs demander, il existe plusieurs resultats ils sont tous retourner.<br>
' Le noeud sera retourner avec son arborescence complete.<br>(ex: root\noeud\noeud1\noeud_cherhcer)
'
'@param nomNoeud Le nom du noeud rechercher
'@param atts Les attributs et leur valeurs <br>(ex: lng=fr,pays=fr || lng=fr,pays=be)
'@return    Tableau contenant le ou les noeuds trouvés
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/17
Public Function rechercherNoeudXML(nomNoeud As String, Optional atts As String = "") As String()
Dim noeud As IXMLDOMNode 'variable pour parcourir la liste de noeuds trouver
Dim noeud2 As IXMLDOMNode 'variable pour parcourir l'arborescence de noeud
Dim parent As String 'variable qui stock l'arborescence du fichier
Dim tmp() As String 'variable stockant le tableau retourner
ReDim tmp(0)
Set root = fXML.documentElement
Set listNode = root.getElementsByTagName(nomNoeud)
For Each noeud In listNode
    If atts = "" Or verifierAttribut(noeud, atts) Then
        If tmp(UBound(tmp)) <> "" Then ReDim Preserve tmp(UBound(tmp) + 1)
        parent = ""
        Set noeud2 = Nothing
        Set noeud2 = noeud.parentNode
        While (Not noeud2.parentNode Is Nothing)
            parent = noeud2.baseName & "/" & parent
            Set noeud2 = noeud2.parentNode
        Wend
        tmp(UBound(tmp)) = parent & noeud.baseName
    End If
Next
rechercherNoeudXML = tmp
End Function
''
' Fonction qui verifie la présence des attributs passés en arguments.
'
'@param noeud Noeud dont on verifie les attributs
'@param atts Les attributs et leur valeurs (ex: lng=fr,pays=fr || lng=fr,pays=be)
'
'@return    Boolean indiquant si les attributs sont biens présent.
'@author    Thomas Darmengeat
'@version   1.0
'@date  2004/06/17
Private Function verifierAttribut(noeud As IXMLDOMNode, atts As String) As Boolean
Dim att() As String 'tableau contenant chaque attribut et sa valeur sous la forme att=valeur
Dim att_tmp() As String 'tableau contenant le nom de l'attribut et sa valeur
    att = Split(atts, ",")
    For i = 0 To UBound(att)
        att_tmp = Split(att(i), "=")
        For j = 0 To noeud.Attributes.length - 1
            If noeud.Attributes.Item(j).nodeName = att_tmp(0) Then
                If noeud.Attributes.Item(j).nodeValue <> att_tmp(1) Then
                    verifierAttribut = False
                    Exit Function
                End If
            End If
        Next
    Next
    verifierAttribut = True
End Function

