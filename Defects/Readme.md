# QC - Quality-Center => Extraction of project Defects


Dashboard :

 - Add new Excel report :
    - Query : With Query Builder, add 2 queries named :
     - "Defects" with the query "Defects-Extraction.sql"
     - "linked" with the query "Defects-links-with-others-elements.sql"
    - Post-processing :
     - add the Post-processing code in Excel VBA "Make Report with Excel.vb"
     - check box run "post-processing"
    - Generation Settings :
     - check box "Launch Report in Excel"


With Query Builder add a Query parameter named "LookingForDefects"  that contains a part of the defect summary (exemple abcd%)

Then click on "Generate"


===========================================================================

Formules pour mettre en place un graphique de type chandelier :
Ouverture = stock d'anomalie ouverte
Stock max = NB ano ouvertes cumulées - NB ano fermées cumulées de la semaine précedente
Stock min = NB ano ouvertes cumulées - NB ano fermées cumulées de la semaine suivante
Fermeture = stock d'anomalie ouverte de la semaine suivante
