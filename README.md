# QC - Quality-Center

   => Extraction of project Defects


Dashboard :

 - Add new Excel report :
    - Query : With Query Builder, add 2 queries named :
     - "Defects List" with the query "Defects-Extraction.sql"
     - "linked" with the query "Defects-links-with-others-elements.sql"
    - Post-processing :
     - add the Post-processing code in Excel VBA
     - check box run "post-processing"
    - Generation Settings :
     - check box "Launch Report in Excel"


With Query Builder add a Query parameter named "LokingFor" (exemple abcd%)

Then click on Generate
