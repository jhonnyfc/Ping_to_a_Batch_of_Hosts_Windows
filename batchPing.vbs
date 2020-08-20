'################################################################### 
'##           Script to check the status of machines              ## 
'##           Author: Unknown                                     ## 
'##           Date: 03-30-2012                                    ## 
'##           modified by: Vikas Sukhija                          ##
'##           Date: 19-08-2020                                    ## 
'##           modified by: Jhonny F. Chicaiza -> github: jhonnyfc ## 
'################################################################### 
 
'# call excel applicationin visible mode 
Dim out
 
Set objExcel = CreateObject("Excel.Application") 
  
objExcel.Visible = True 
  
objExcel.Workbooks.Add 
  
intRow = 2 
  
'# Define Labels  
  
objExcel.Cells(1, 1).Value = "Machine Name" 
  
objExcel.Cells(1, 2).Value = "Results" 

objExcel.Cells(1, 3).Value = "ip" 
  
  
'# Create file system object for reading the hosts from text file  
Set Fso = CreateObject("Scripting.FileSystemObject") 
  
Set InputFile = fso.OpenTextFile("impLis.txt")

'# Create shell object for Pinging the host machines 
Set WshShell = WScript.CreateObject("WScript.Shell")
	
  
'# Loop thru the text file till the end
Do While Not (InputFile.atEndOfStream)
  
	HostName = InputFile.ReadLine
	
	Set exe = WshShell.Exec("%COMSPEC% /c ping -n 3 """ & HostName)
	out = exe.StdOut.ReadAll

	objExcel.Cells(intRow, 1).Value = HostName
	If InStr(out, "Respuesta") = 0 Then
		objExcel.Cells(intRow, 2).Value = "down"
		objExcel.Cells(intRow,3).Value = "no encontrado"
	Else
		objExcel.Cells(intRow, 2).Value = "ok"
		objExcel.Cells(intRow,3).Value = getIP(out)
	End If
	
	intRow = intRow + 1
Loop


Function getIP(text)
	Dim line,aux
	
	for each x in Split(text,vbCrLf)
		if 1 = InStr(x, "Respuesta desde") then
			aux = Split(x) 
			getIP = Replace(aux(2),":","")
			Exit for
		End If
	next
End Function