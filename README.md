
# Apache POI (Poor Obfuscation Implementation)

- projet open source du groupe Apache, pour manipuler les fichiers de **Microsoft Office** en Java.
- le format **OLE2** a pour nom complé Microsoft Office Binary Document Container Format, comprend divers format de fichiers
- il existe d'autres bibliothèques comme **Fastexcel** qui est plus simple mais moins complète
- il existe aussi des bibilothèques pour manipuler les formats *Open Document* comme **jOpenDocument** mais peut utilisée

Il existe plusieurs composants :
- **POIFS** (Poor Obfuscation Implementation File System) : lien entre les objets OLE2 et Java
- **HSSF et XSSF** : fichiers Excel
- **HWPF** (Horrible Word Processor Format) : fichiers Word
- **HSLF** (Horrible Slide Layout Format) : fichiers Powerpoint
- d'autres pour les fichiers Dessin, Publisher, Visio, Outlook...

Nous allons ici nous concenter sur les fichiers **Excel**. Il existe 2 implémentations pour lire/écrire les fichiers Excel :
- **HSSF** (Horrible SpreadSheet Format) : Excel 2003 ou avant (xls)
- **XSSF** (XML SpreadSheet Format) : Excel 2007 ou après (xlsx)

avantages du XLSX vs XLS :
- nouveau format respecte la norme XML
- XLSX taille réduite
- davantage de lignes et de colonnes (1 million de lignes vs 65 000 et 16 000 colonnes vs 256)
- XLS peut enregistrer des macros, XLSX ne peut pas (XLSM oui)

## Les interfaces

- **Workbook :** classeur Excel
- **Sheet :** feuille Excel (grille de cellules), étend l'interface **java.lang.Iterable**
- **Row :** ligne de la feuille, étend l'interface **java.lang.Iterable**
- **Cell :** représente une cellule dans une ligne d'une feuille
- **CellStyle** : pour manipuler le style des cellules

## Les classes

- les classes **XLS** : HSSFWorkbook, HSSFSheet, HSSFRow, HSSFCell
- les casses **XLSX** : XSSFWorkbook, XSSFSheet, XSSFRow, XSSFCell

## La bibliothèque

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
</dependency>
```

## Lire un fichier Excel (xls et xlsx)

https://www.javatpoint.com/how-to-read-excel-file-in-java

## Ecrire un fichier Excel

- https://www.baeldung.com/java-microsoft-excel : lecture et écriture d'un fichier simple + comparaison autres bibliothèques
- https://www.geeksforgeeks.org/how-to-write-data-into-excel-sheet-using-java/ : exemple d'écriture avec une boucle d'une liste
- https://www.codejava.net/coding/how-to-write-excel-files-in-java-using-apache-poi : exemple un peu plus complet
