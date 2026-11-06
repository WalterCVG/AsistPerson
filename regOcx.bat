if not exist c:\windows\system32\MSWINSCK.OCX 
   rem copy \\hepdc1\netlogon\MSWINSCK.OCX  c:\windows\system32 
   echo  copy \\hepdc1\netlogon\MSWINSCK.OCX  c:\windows\system32    
   rem REGSVR32 /s c:\windows\system\MSWINSCK.OCX 
   echo  REGSVR32 /s c:\windows\system\MSWINSCK.OCX 
goto Chau
else
   rem REGSVR32 /s c:\windows\system32\MSWINSCK.OCX 
   echo REGSVR32 /s c:\windows\system32\MSWINSCK.OCX 
:chau



