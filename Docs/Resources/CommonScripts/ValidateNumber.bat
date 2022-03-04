set NAME=%1
set VALUE=%2
set MIN=%3
set MAX=%4
if %VALUE% LSS %MIN% (
	set ERRORMESSAGE=%NAME%=%VALUE% is invalid: it must be between %MIN% and %MAX%
	set ERROR=1
)
if %VALUE% GTR %MAX% (
	set ERRORMESSAGE=%NAME%=%VALUE% is invalid: it must be between %MIN% and %MAX%
	set ERROR=1
)
exit /B 
