# Documentation d'oPPTimiz

## Technologies utilisées

L'outil oPPTimiz repose sur plusieurs technologies :
- L'addin PowerPoint est un complément de l'application PowerPoint au format VSTO (C#). Il est basé sur la version 4.8 du .NET Framework
- Le script en ligne de commande oPPTimiz est un script PowerShell qui a été compilé en fichier exécutable. Ce script intéragit avec PowerPoint au travers d'un composant COM.

La documentation de génération se base sur l'utilisation de l'IDE Visual Studio pour l'addin VSTO.
Le script en ligne de commande peut être compilé par n'importe quel utilitaire prévu à cet effet.

La solution oPPTimiz est compatible avec les versions PowerPoint 2013 et supérieures, y compris les versions M365.

## Principe de fonctionnement

L'optimisation des présentations PowerPoint au travers de l'outil oPPTimiz (directement via PowerPoint, ou en mode ligne de commande) consiste en plusieurs étapes :
- Retrait de l'option de présentation marquée comme finale si activée sur le document
- Retrait des masques et des dispositions inutilisés
- Compression des images présentes dans le document
- Enregistrement du fichier pour récupérer le gain réalisé
- Ajout des informations de gain d'optimisation dans les propriétés du fichier
- Reconfiguration de l'option de présentation marquée comme finale si elle était précédemment active
- Enregistrement du document avec les nouveaux éléments

Le script en mode ligne de commande reprend la même logique, à ceci près qu'il doit au préalable lancer PowerPoint et gérer le cas ou l'application est déjà lancée.

### Gestion de l'option présentation marquée comme finale

Les présentations PowerPoint peuvent être marquée comme finale pour prévenir toute modification ultérieure. Afin de permettre à l'outil oPPTimiz de pouvoir modifier ce type de présentations, il est nécessaire de désactiver ce marquage en tant que présentation finale. Cette modification du statut de la présentation se caractérise par :
```PowerShell
$application = New-Object -ComObject powerpoint.application
$presentation = $application.Presentations.open($pptFile)
$presentation.Final = $false
```
```csharp
PowerPoint.Presentation presentation = Application.ActivePresentation;
presentation.Final = true;
```
Dès lors qu'une présentation était marquée comme finale avant l'optimisation, l'option est réactivée en fin d'opération. Les présentations qui n'étaient pas marquées comme finales ne sont pas passées en mode finale à la fin de l'optimisation.

### Gestion de la compression des images

Afin de garantir la compression de toutes les images, peu importe où elles sont situées dans la présentation, l'outil oPPTimiz ajoute temporairement une image de 220ppi sur la première slide et la sélectionne. 
```csharp 
PowerPoint.Shape shape = presentation.Slides[1].Shapes.AddPicture2($@"{path-to-image}", Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 0, 0, -1, -1);
shape.Select();
```
L'étape suivante est d'exécuter la fonctionalité de compression d'image de PowerPoint (bouton "*Compreser les images*" dans le menu "*Format de l'image*"). Cette fonctionalité correspond au MSO *PictureCompress* :
```csharp
Application.CommandBars.ExecuteMso(Constants.MsoPictureCompress);
```
Cette fonctionalité de compression des images ouvre une fenêtre déportée qui propose les options disponibles. L'outil va envoyer les entrées clavier pour reproduire la configuration choisie (compression maximale ou intermédiaire) puis valider l'action.

Après exécution, le MSO *PictureCompress* est rééxécuté pour décocher les cases modifiées par la première exécution pour restaurer l'interface à son état initial. Sans cette étape, il ne serait pas possible de réexécuter plusieurs compressions successives.

L'image ajoutée en prérequis est ensuite supprimée de la présentation.

>**NB**: A date, il n'existe pas de méthode native, sans passer par l'interface utilisateur, permettant de compresser les images de la présentation. Il n'y a donc pas d'alternative au fait de devoir déclencher les entrées clavier associées aux actions dans les menus de la fenêtre de compression des images.

>**Limitation**: Le recours aux entrées clavier associées aux menus de l'interface de compression des images rend la solution sensible à la langue d'affichage de PowerPoint pour les versions M365. Seules les langues **FR** et **EN** sont actuellement supportées. Pour plus d'informations pour intégrer de nouvelles langues supportées, référez-vous à la section **Personnalisation** de ce document.

### Récupération et stockage des informations de gain

Une fois la présentation optimisée, elle est sauvegardée afin de déterminer le gain réalisé par rapport à la version initiale.
Les informations de l'optimisation du fichier sont stockées dans des propriétés custom du fichier :
- **oPPTimizDate** : stockage de la date de la dernière optimisation
- **oPPTimizGain** : gain en octets de l'optimisation de la présentation
- **oPPTimizRatio** : pourcentage du gain par rapport à la taille initiale de la présentation

```csharp
//Ajout d'une propriété
object[] customPropertyArgs = { oArg.Name, false, oArg.Type, oArg.Value };
typeDocCustomProps.InvokeMember(Constants.pptPropertyMethodAdd, BindingFlags.Default | BindingFlags.InvokeMethod, null, customProperty, customPropertyArgs);

//Mise à jour d'une propriété
object[] customPropertyArgs = { oArg.Name, oArg.Value };
typeDocCustomProps.InvokeMember(Constants.pptPropertyMethodUpdate, BindingFlags.Default | BindingFlags.SetProperty, null, customProperty, customPropertyArgs);
```

## Personnalisation du contenu de l'addin

### Gestion des langues supportées

Dans l'addin, les éléments dépendants de la langue de PowerPoint sont centralisés dans le fichier `LanguageResources.cs`. Ce fichier défini la classe du même nom qui permet de définir un ensemble de propriétés qui dépendent de la langue d'affichage de PowerPoint, avec par exemple :
- Les textes affichés dans le bandeau de PowerPoint
- Les textes des éléments de l'interface de configuration de l'optimisation de la présentation
- Les messages affichés à la fin de l'opération d'optimisation de la présentation
- Les raccourcis clavier dépendants de la langue (voir la section **Gestion de la compression des images** de ce document)

Pour ajouter une nouvelle langue supportée, il faut ajouter un nouveau `case` avec le code ISO 639-1 de deux lettres correspondant à la langue à ajouter dans le `switch` sur la langue d'affichage de l'utilisateur. Cette dernière est récupérée via la méthode `GetCulture` mise à disposition par le SDK Office.

Par défaut, si la langue d'affichage n'est pas le français, c'est la langue anglaise qui est utilisée pour tous les éléments listés ci-dessus.
___
Dans le script en ligne de commande, la langue est déterminée via la commande `(Get-WinSystemLocale).LCID`. Les éléments dépendants de la langue sont centralisés en début de fichier, dans la région **Localization** :
- Les messages d'erreur à afficher à l'utilisateur
- Les raccourcis clavier dépendants de la langue (voir la section **Gestion de la compression des images** de ce document)

Pour ajouter une langue, il est nécessaire d'ajouter les variables correspondantes dans la région **Localization** et de modifier :
- La méthode `Compress-Picture` dont les raccourcis clavier dépendent de la langue pour les version M365 (voir la section **Gestion de la compression des images** de ce document)
- La gestion d'erreur dans la région **Handle previous running PowerPoint instances**

Par défaut, si la langue n'est pas le français, c'est la langue anglaise qui est utilisée pour tous les éléments listés ci-dessus.

### Modification des constantes

Dans l'addin, l'ensemble des constantes est rassemblé dans le fichier `Constants.cs`. Il est notamment possible de modifier dans ce fichier :
- Le suffixe des documents PowerPoint optimisés (une distinction est possible selon le mode d'optimisation)
```csharp
public const string IntermediateOptimizedFilenameSuffix = "_oPPTimiz";
public const string MaximumOptimizedFilenameSuffix = "_oPPTimiz";
```
- La clé de registre de configuration de l'outil pour le seuil de compression (si le pourcentage de gain de taille par rapport au fichier initial est inférieur à ce seuil, le fichier est considéré comme déjà optimisé et n'est pas modifié). Par défaut, la valeur du seuil est fixée à 5%.
```csharp
public const string RegKeyOpptimiz = @"Software\oPPTimiz";
public const string RegValueThreshold = "OptimizationThreshold";
```
- Le nom des propriétés ajoutées à la présentation
```csharp
public const string pptPropertyDate = "oPPTimizDate";
public const string pptPropertyGain = "oPPTimizGain";
public const string pptPropertyRatio = "oPPTimizRatio";
```
___

Dans le script en ligne de commande, on retrouve la même logique, regroupée dans la région **Constants**, avec notamment :
- L'emplacement des logs d'exécution
```PowerShell
$sLogDirectory = [string]::Format("{0}\Souche\Logs\", $env:PROGRAMDATA)
```
- La taille maximale du fichier de log avant rotation automatique (par défaut 1 Mo)
- La configuration du seuil d'optimisation
```PowerShell
$RegKeyOpptimiz = "HKCU:\SOFTWARE\oPPTimiz"
$RegValueThreshold = "OptimizationThreshold"
$DefaultOptimizationThreshold = 5
```
- Le nom des propriétés ajoutées à la présentation
```PowerShell
$pptPropertyDate = "oPPTimizDate"
$pptPropertyGain= "oPPTimizGain"
$pptPropertyRatio = "oPPTimizRatio"
```

### Personalisation de la fenêtre d'options de l'addin

Dans l'addin, il est possible d'ajouter un logo dans la fenêtre de configuration de l'optimisation du fichier. Pour ce faire, il faut remplacer le fichier `Src\oPPTimiz\Resources\Logo.png`.
Il est par exemple possible d'ajouter un logo "*Numérique Responsable*" si vous avez obtenu ce label.

## Mode de génération des binaires

### Génération de l'addin VSTO

Afin de garantir l'installation de l'addin, le code de ce dernier doit être signé à la compilation, afin de garantir son intégrité dans l'application PowerPoint. Pour ce faire, il est nécessaire de disposer d'un certificat de signature de code issu d'une autorité de certification de confiance déployée sur les postes ciblés.
Ce certificat doit être ajouté dans l'IDE Visual Studio, via la procédure suivante :
- Aller dans les propriétés du projet "**oPPTimiz**"
- Aller dans la rubrique "**Signature**"
- Ajouter le certificat de signature depuis le magasin de certificat du poste ou depuis un fichier. L'usage d'un certificat auto-signé est à éviter pour un usage en production.
- Veiller à ce que la case "**Signer l'assembly**" soit bien cochée

Vous pouvez ensuite **Regénérer** le projet et récupérer les binaires suivants :
- Resources *(répertoire)*
  - 220ppi.png *(image utilisée pour la compression des images de la présentation)*
- Microsoft.Office.Tools.Common.v4.0.Utilities.dll
- oPPTimiz.dll
- oPPTimiz.dll.manifest
- oPPTimiz.vsto

Tous ces éléments doivent être déployés sur les postes pour que l'addin soit fonctionnel.

### Génération du script en ligne de commande

Le script PowerShell doit être compilé en fichier exécutable afin d'être exécuté correctement via le menu contextuel de Windows. 
Plusieurs outils existent pour répondre à ce besoin et n'ont pas d'incidence particulère sur le fonctionnement d'oPPTimiz.

## Packaging de la solution oPPTimiz

La solution oPPTimiz peut être intégrée dans un package MSI afin de faciliter son déploiement (dépose des fichiers et configurations dans le registre Windows).

Ce paragraphe décrit la réalisation de ce package MSI au travers de l'outil *InstallShield*.

- Ouvrir le fichier `MSI\oPPTimiz\oPPTimiz.ism` avec le logiciel InstallShield
- Activer le mode d'affichage "**View List**"
- Se rendre dans la rubrique "**MEDIA\Path variables**"
- Modifier la variable "**PATH_TO_RESOURCES_FILES**" en remplacaçant `[PATH_TO_PROJECT]` par l'emplacement du projet en chemin absolu
- *(Facultatif)* Dans la rubrique "**INSTALLATION INFORMATION\General Information**", modifier l'emplacement d'installation de l'addin si nécessaire via le champ `INSTALLDIR`
- Déposer les ressources suivantes dans le répertoire `[PATH_TO_PROJECT]\MSI\oPPTimiz\Resources` :
  - Microsoft.Office.Tools.Common.v4.0.Utilities.dll *(généré dans le dossier de sortie du projet oPPTimiz.sln)*
  - oPPTimiz.dll *(généré dans le dossier de sortie du projet oPPTimiz.sln)*
  - oPPTimiz.dll.manifest *(généré dans le dossier de sortie du projet oPPTimiz.sln)*
  - oPPTimiz.vsto *(généré dans le dossier de sortie du projet oPPTimiz.sln)*
  - Resources\220ppi.png *(généré dans le dossier de sortie du projet oPPTimiz.sln)*
  - oPPTimiz.ico *(icône du projet localisée sous `Src\oPPTimiz\Resources\oPPTimiz.ico`)*
  - oPPTimiz.exe *(script PowerShell compilé)*
- Cliquer sur "**Build**"

Le package MSI est généré dans le répertoire `[PATH_TO_PROJECT]\MSI\oPPTimiz\oPPTimiz\oPPTimiz.msi`.