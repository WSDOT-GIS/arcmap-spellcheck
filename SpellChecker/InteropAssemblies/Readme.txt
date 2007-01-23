The contents of this folder are the COM Interop dlls necessary in order to use the MSWORD9.OLB.

This needed to be done at the command line in order to give the Word.dll metadata assembly a strong name 
(so that the ArcMapSpellCheck assembly that references it could have a strong name).