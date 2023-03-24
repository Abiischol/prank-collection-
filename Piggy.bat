@echo off

rem Set the name and location of the output file
set OUTPUT_FILE=output.rtf

rem Create the output file
echo {\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Calibri;}}\uc0\pard\tx720\cf1\f0\fs28 This little piggy went wee wee wee all the way home.} > %OUTPUT_FILE%

rem Open the output file
start %OUTPUT_FILE%

rem Delete the batch file
del "%~f0"
