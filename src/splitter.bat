@echo off

set CURRENT_DIR=%~dp0

set CLASSPATH=%CURRENT_DIR%;%CURRENT_DIR%\..\lib\javax.mail.jar;%CURRENT_DIR%\..\lib\log4j-1.2.17.jar;%CURRENT_DIR%\..\lib\poi-3.12-20150511.jar;%CURRENT_DIR%\lib\poi-examples-3.12-20150511.jar;%CURRENT_DIR%\..\lib\poi-excelant-3.12-20150511.jar;%CURRENT_DIR%\..\lib\poi-ooxml-3.12-20150511.jar;%CURRENT_DIR%\..\lib\poi-ooxml-schemas-3.12-20150511.jar;%CURRENT_DIR%\..\lib\poi-scratchpad-3.12-20150511.jar;%CURRENT_DIR%\..\lib\xmlbeans-2.6.0.jar

java -classpath %CLASSPATH% Splitter

pause