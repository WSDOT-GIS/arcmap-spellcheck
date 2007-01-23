@echo off
rem Set the environment variables necessary for the sn and tlbimp tools.
call "c:\Program Files\Microsoft Visual Studio .NET 2003\Common7\Tools\vsvars32.bat"

rem files needing interop
rem "C:\Program Files\Microsoft Office\Office\MSWORD9.OLB"

rem set directory variables
rem set OfficeDir=C:\Program Files\Microsoft Office\Office\
set OfficeDir=%ProgramFiles%\Microsoft Office\Office

rem set COM path variables
set Word=%OfficeDir%\MSWORD9.OLB

@echo Creating strong name key files.
@echo on
if not exist Word.snk sn -k Word.snk 
else echo Word.snk already exists.

@echo Creating interop files.
@rem tlbimp "%Word%" /keyfile:Word.snk /out:Word.dll /namespace:Word
tlbimp "%Word%" /keyfile:Word.snk

