Dim message, sapi
if WScript.Arguments.Count > 0 then
	message=Wscript.Arguments(0)
	Set sapi=CreateObject( "sapi.spvoice" )
elseif WScript.Arguments.Count = 0 then
	message="Parameters are missing"
	Set sapi=CreateObject( "sapi.spvoice" )
end if
sapi.Speak message