SELECT
------------------------------------------------------------
LINK.LN_BUG_ID As "id",
LINK.LN_ENTITY_TYPE As "Type",
LINK.LN_ENTITY_ID As "id-Type",
LINK.LN_LINK_COMMENT As "Comment"
------------------------------------------------------------
FROM Link
Where  LINK.LN_BUG_ID in(Select BUG.BG_BUG_ID From BUG Where BG_SUMMARY Like '@LookingFor@' )
------------------------------------------------------------
Order By
LINK.LN_BUG_ID Asc,
LINK.LN_ENTITY_TYPE asc
