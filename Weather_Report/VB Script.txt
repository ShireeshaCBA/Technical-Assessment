set objShell = WScript.CreateObject ("WScript.Shell")
Dim exo,wbo,wso,x,i,y,z,a,b,c,d,p
set exo = createObject("Excel.application")
exo.visible = true
set wbo = exo.workbooks.open(objShell.CurrentDirectory & "\" & "Weather.xlsx",1,true)
set wso = wbo.worksheets("Sheet1")

p = Trim(UCase(inputbox("Enter the location name")))
if (p = "SYDNEY") Then 
		x = wso.cells(2,1)
		y = wso.cells(2,2)
		z = wso.cells(2,3)
		a = wso.cells(2,4)
		b = wso.cells(2,5)
		c = wso.cells(2,6)
		d = wso.cells(2,7)
		msgbox("Output in the format requested :" & vbCrLf & x & "|" & "-" & y  & "|" & z & "|" & a & "|" & b & "|" & c & "|" & d)
Elseif (p = "MELBOURNE") Then
	    x = wso.cells(3,1)
		y = wso.cells(3,2)
		z = wso.cells(3,3)
		a = wso.cells(3,4)
		b = wso.cells(3,5)
		c = wso.cells(3,6)
		d = wso.cells(3,7)
		msgbox("Output in the format requested :" & vbCrLf & x & "|" & y  & "|" & z & "|" & a & "|" & b & "|" & c & "|" & d)
Elseif (p = "ADELAIDE") Then
	    x = wso.cells(4,1)
		y = wso.cells(4,2)
		z = wso.cells(4,3)
		a = wso.cells(4,4)
		b = wso.cells(4,5)
		c = wso.cells(4,6)
		d = wso.cells(4,7)
		msgbox("Output in the format requested :" & vbCrLf & x & "|" & y  & "|" & z & "|" & a & "|" & b & "|" & c & "|" & d)
Elseif (p = "GOA") Then
	    x = wso.cells(5,1)
		y = wso.cells(5,2)
		z = wso.cells(5,3)
		a = wso.cells(5,4)
		b = wso.cells(5,5)
		c = wso.cells(5,6)
		d = wso.cells(5,7)
		msgbox("Output in the format requested :" & vbCrLf & x & "|" & y  & "|" & z & "|" & a & "|" & b & "|" & c & "|" & d)
Elseif (p = "SRILANKA") Then
	    x = wso.cells(6,1)
		y = wso.cells(6,2)
		z = wso.cells(6,3)
		a = wso.cells(6,4)
		b = wso.cells(6,5)
		c = wso.cells(6,6)
		d = wso.cells(6,7)
		msgbox("Output in the format requested :" & vbCrLf & x & "|" & y  & "|" & z & "|" & a & "|" & b & "|" & c & "|" & d)
Else
		msgbox("Please enter a valid location")
End if
exo.quit
set wso = nothing
set wbo = nothing
set exo = nothing