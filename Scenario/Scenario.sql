Select
TC_CYCLE_ID As "Test Set id",
(Select CY_CYCLE From CYCLE Where CY_CYCLE_ID = TC_CYCLE_ID) As "Scénario",
TC_TEST_ORDER As "Ordre",
TC_TEST_INSTANCE As "Instance",
TC_TEST_ID As "Test id",
TC_STATUS As "Etat",
(Select TS_NAME From TEST Where TS_TEST_ID = TC_TEST_ID ) As "Test Name"
From TESTCYCL
Where TC_CYCLE_ID = '3630'
Order By TC_CYCLE_ID asc, TC_TEST_ORDER asc
