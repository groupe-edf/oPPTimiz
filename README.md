# Bienvenue sur oPPTimiz

## Qu'est-ce que c'est ?

oPPTimiz est un addin intégré dans l'interface PowerPoint qui permet de réduire la taille des présentations en compressant les images et en effaçant les masques de diapositives inutilisés.

Il fonctionne dans l'environnement Windows pour les fichiers stockés sur le poste en local ou sur des ressources partagées (OneDrive, Teams, ...).

La solution oPPTimiz est également disponible en mode ligne de commande, et via le menu contextuel (clic droit) de l'explorateur Windows.

## Comment l'utiliser ?

### Utilisation depuis PowerPoint

Depuis le bandeau de PowerPoint, un groupe "**Numérique responsable**" est ajouté. Il comporte deux boutons :
- **Optimiser** qui est la fonctionalité d'optimisation de la taille de la présentation sur le disque apportée par l'addin
- **Vérifier l'accessibilité** qui est un raccourci vers la fonctionalité Microsoft de vérification de l'accessibilité de la présentation

Une fois que la présentation est terminée, cliquer sur le bouton "**Optimiser**". Une fenêtre d'options s'ouvre et permet de sélectionner le niveau d'optimisation souhaité :
- "**Optimisé**" (*sélectionné par défaut*) : compatible avec la majorité des usages de présentation PowerPoint
- "**Standard**" : pour l'utilisation sur grand écran ou les impressions en Haute Définition (HD)

Deux modes d'exécution sont possibles :
- "**Optimiser et remplacer**" : permet de remplacer le fichier courant par sa version optimisée
- "**Optimiser et conserver**" : permet de conserver le fichier initial en plus de sa version optimisée. 
Le fichier optimisé est créé au même emplacement que le fichier initial avec le suffixe "*_oPPTimiz.pptx*".
A privilégier si la présentation doit être enrichie de diapositives basées sur une nouvelle disposition.

Une fenêtre résumant l'optimisation de la présentation s'affiche en fin de traitement pour indiquer le gain réalisé. Si le pourcentage de réduction de taille par rapport à la taille initiale du document est inférieur à un certain seuil (5% par défaut), le document est considéré comme étant déjà optimisé, et le message indique que le document l'est déjà. Si l'optimisation a été lancée en mode "**Optimiser et conserver**", le fichier optimisé est également supprimé.
Ce seuil est paramétrable au travers de la clé de registre :
```
[HKEY_CURRENT_USER\SOFTWARE\oPPTimiz]
OptimizationThreshold (DWORD)
```

### Utilisation en mode ligne de commande

```
oPPTimiz.exe -pptFile source [-compressionLevel [Maximal | Intermediate]] [-keepFile [0 | 1]]

source                  Fichier à optimiser
-compressionLevel       Niveau de compression des images à appliquer (par défaut le niveau est Maximal)
-keepFile               Conserve ou écrase le fichier initial (par défaut le fichier est écrasé (valeur 0))
```
**Exemples :**
```
oPPTimiz.exe -pptFile source -compressionLevel Maximal
```
```
oPPTimiz.exe -pptFile source -compressionLevel Intermediate -keepFile 1
```

#### Utilisation via le menu contextuel (clic droit) de l'explorateur Windows

Il est possible de lancer l'optimisation d'un ou plusieurs fichiers en les sélectionnant dans l'explorateur Windows via le menu contextuel (clic droit). Les deux modes de fonctionnement, décrits dans la section **Utilisation depuis PowerPoint**, sont disponibles via les options :
- Ecraser le fichier d'origine
- Conserver le fichier d'origine

Dans ces deux modes le niveau de compression est configuré à *Maximal*.
## Comment l'installer ?

### Installation rapide

L'addin peut être installé via l'installation d'un package MSI (voir Notice.md pour plus d'information sur la génération d'un tel package)
``` console
msiexec /i oPPTimiz.msi
```

### Installation manuelle
Pour enregistrer l'addin dans PowerPoint, il est nécessaire de créer les éléments de registre suivants (pensez à remplacer `[INSTALLDIR]` par le répertoire où sont situées les sources) :
```
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins\oPPTimiz]
"Description"="oPPTimiz"
"FriendlyName"="oPPTimiz"
"LoadBehavior"=dword:00000003
"Manifest"="file:///[INSTALLDIR]oPPTimiz.vsto|vstolocal"

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins\oPPTimiz]
"Description"="oPPTimiz"
"FriendlyName"="oPPTimiz"
"Manifest"="file:///[INSTALLDIR]oPPTimiz.vsto|vstolocal"
"LoadBehavior"=dword:00000003
```

Pour enregistrer le mode de lancement depuis le menu contextuel (clic droit), il est nécessaire de créer les éléments de registre suivants (pensez à remplacer `[INSTALLDIR]` par le répertoire où sont situées les sources) :

```
[HKEY_CLASSES_ROOT\SystemFileAssociations\.ppt\shell\oPPTimiz]
"MUIVerb"="oPPTimiz"
"Icon"="[INSTALLDIR]Resources\\oPPTimiz.ico"
"subcommands"="oPPtimiz.overrideFile;oPPTimiz.keepFile"

[HKEY_CLASSES_ROOT\SystemFileAssociations\.pptx\shell\oPPTimiz]
"MUIVerb"="oPPTimiz"
"Icon"="[INSTALLDIR]Resources\\oPPTimiz.ico"
"subcommands"="oPPtimiz.overrideFile;oPPTimiz.keepFile"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\oPPTimiz.overrideFile]
@="Ecraser le fichier d'origine"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\oPPTimiz.overrideFile\Command]
@="[INSTALLDIR]oPPTimiz.exe -pptFile \"%1\" -keepFile 0"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\oPPTimiz.keepFile]
@="Conserver le fichier d'origine"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\oPPTimiz.keepFile\Command]
@="[INSTALLDIR]oPPTimiz.exe -pptFile \"%1\" -keepFile 1"
```