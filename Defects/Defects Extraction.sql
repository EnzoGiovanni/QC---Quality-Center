SELECT
------------------------------------------------------------
BG_BUG_ID As "id",
------------------------------------------------------------
to_char(BG_DETECTION_DATE, 'iyyy')||'-'||to_char(BG_DETECTION_DATE, 'iw') As "Opened Y-W",
'to_char(BG_DETECTION_DATE, 'iyyy') As "Opened Year",
'to_char(BG_DETECTION_DATE, 'mm') As "Opened Month",
'to_char(BG_DETECTION_DATE, 'iw') As "Opened Week",
to_char(BG_DETECTION_DATE, 'yyyy/mm/dd') As "Opened Date",
------------------------------------------------------------
to_char(BG_CLOSING_DATE, 'iyyy')||'-'||to_char(BG_CLOSING_DATE, 'iw') As "CLosed Y-W",
'to_char(BG_CLOSING_DATE, 'iyyy') As "Closed Year",
'to_char(BG_CLOSING_DATE, 'mm') As "Closed Month",
'to_char(BG_CLOSING_DATE, 'iw') As "Closed Week",
to_char(BG_CLOSING_DATE, 'yyyy/mm/dd') As "Closed Date",
------------------------------------------------------------
BG_STATUS As "Etat",
BG_SEVERITY As "Gravitée",
BG_SUMMARY As "Résumé",
------------------------------------------------------------
BG_DETECTED_BY As "detected by",
BG_RESPONSIBLE As "assigned to"
------------------------------------------------------------
FROM BUG
Where BG_SUMMARY Like '@LookingForDefects@'
------------------------------------------------------------
Order By
BG_DETECTION_DATE Asc,
BG_BUG_ID Asc
