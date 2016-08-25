Dim txt
Dim InputFile
Dim Path
Dim objFile
Dim sScriptFileName
Dim arrCSVfile
Dim arrSwitchperStore()
Dim arrSwitchperCSV()
Dim strStore
Dim arrStores
Dim TimeZoneReboot_LogFile
Dim MDF_IDF_Verify_LogFile
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 


arrCSVfile = Array("MDF","IDF2","IDF3","IDF4","IDF5","IDF6","IDF7","IDF8","IDF9","IDF10","IDF11")
sScriptFileName = "CreateCVS.vbs"
MDF_IDF_Verify_LogFile = "MDF_IDF_VerifyLog.txt"

Set Connection = CreateObject("ADODB.Connection")
Set RecordSet = CreateObject("ADODB.recordset")
Set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.GetFile(sScriptFileName)
Set args=wscript.arguments

' Checking Argument list
If Args.Count  < 3 Then 
	InputFile = args(0)

Else
	Wscript.Echo "You didn't enter a Valid argument"
	WScript.quit(1)

End IF

'Connection to MS Excel file(XLS) that contains MDF and IDF IP switch information per Store uding ODB
Path = Left(objFile.Path,len(objFile.Path) - len(sScriptFileName))
Set objInputFile = objFSO.OpenTextFile(Path & InputFile, 1)

With Connection
   .Provider = "Microsoft.Jet.OLEDB.4.0"
   .ConnectionString = "Data Source=" & Path & "Switch Stack Ip Addresses.xls;Extended Properties=Excel 8.0;"
   .Open
End With

'Reading file information, assign to variable 
sStore = objInputFile.ReadAll
arrStores = Split(sStore, vbCrLf)
objInputFile.Close





ReDim arrSwitchperStore(Ubound(arrStores) + 1)
ReDim arrSwitchperCSV(Ubound(arrStores) + 1)





ResetStoreCounters(arrSwitchperStore)
ClearSwitchperCSV(arrSwitchperCSV)




If Args.Count = 2 Then 	
	TimeZoneName = args(1)
	'Wscript.Echo TimeZoneName
	Select Case LCase(TimeZoneName)
	 Case "e"
		TimeZoneName = "Eastern"
		TimeZoneReboot_LogFile = TimeZoneName & "_RebootLog.txt"	
	 Case "c"
		TimeZoneName = "Central"
		TimeZoneReboot_LogFile = TimeZoneName & "_RebootLog.txt"
	 Case "m"
		TimeZoneName = "Mountain"
		TimeZoneReboot_LogFile = TimeZoneName & "_RebootLog.txt"
	 Case "p"	
		TimeZoneName = "Pacific"  
		TimeZoneReboot_LogFile = TimeZoneName & "_RebootLog.txt"
	End Select

	TimeZone TimeZoneName,arrStores,arrSwitchperCSV,arrCSVfile, TimeZoneReboot_LogFile

Else
	Generate_MDF_IDF_files arrStores,arrSwitchperCSV,arrCSVfile,MDF_IDF_Verify_LogFile

End IF







Sub Generate_MDF_IDF_files(arrStores,arrSwitchperCSV,arrCSVfile,Logfile)


Set objMDFTextFile = objFSO.OpenTextFile _
    (Path & arrCSVfile(0) & ".csv", ForAppending, True)

Set objVerifyTextFile = objFSO.OpenTextFile _
    (Path & "Verify.csv", ForAppending, True)

objMDFTextFile.WriteLine("primaryIPAddress,hostName,scriptField1,scriptField2,scriptField3")
objVerifyTextFile.WriteLine("primaryIPAddress,hostName,scriptField1,scriptField2,scriptField3")


Wscript.Echo "Working on " & arrCSVfile(0) & " Switch below"

For Index = Lbound(arrStores) to Ubound(arrStores)
	strStore = UCase(arrStores(Index))
	Wscript.Echo "Finding " & arrCSVfile(0) & " Record for Store: " & strStore

	RecordSet.Open "SELECT  ['MDF-IDF Ip Addresses$'].MDF_IP_addresses , ['MDF-IDF Ip Addresses$'].MDF " & _
	"FROM ['MDF-IDF Ip Addresses$'] WHERE ['MDF-IDF Ip Addresses$'].MDF LIKE '" & strStore & "%'", Connection

	Do While Not RecordSet.EOF

		arrSwitchperStore(Index) = arrSwitchperStore(Index) + 1
		arrSwitchperCSV(Index) = arrSwitchperCSV(Index) & arrCSVfile(0) 
		
		'txt = RecordSet.fields(0) & "," & RecordSet.fields(1) & ",,,"
		txt = RecordSet.fields(0) & ",,,,"
		objMDFTextFile.WriteLine(txt)
		objVerifyTextFile.Writeline(txt)  		
		RecordSet.MoveNext
	Loop

	RecordSet.Close
	
Next

objMDFTextFile.Close


For number = 2 To 11
	Wscript.Echo
	Wscript.Echo "Working on IDF" & number & " below"
	Set objIDFTextFile = objFSO.OpenTextFile _
   	(Path & "IDF" & number & ".csv", ForAppending, True)
	


	objIDFTextFile.WriteLine("primaryIPAddress,hostName,scriptField1,scriptField2,scriptField3")

	
	For Index = Lbound(arrStores) to Ubound(arrStores)
		strStore = UCase(arrStores(Index))
		Wscript.Echo "Finding IDF" & number & " Record for Store: " & strStore

		RecordSet.Open "SELECT  ['MDF-IDF Ip Addresses$'].IDF_" & number & "_IP_Addresses , ['MDF-IDF Ip Addresses$'].IDF_" & number & "_Switches " & _
		"FROM ['MDF-IDF Ip Addresses$'] WHERE ['MDF-IDF Ip Addresses$'].IDF_" & number & "_Switches LIKE '" & strStore & "%'", Connection

		
		Do While Not RecordSet.EOF
			arrSwitchperStore(Index) = arrSwitchperStore(Index) + 1
			arrSwitchperCSV(Index) = arrSwitchperCSV(Index) & ", " & arrCSVfile(number-1) 
			'txt = RecordSet.fields(0) & "," & RecordSet.fields(1) & ",,,"
			txt = RecordSet.fields(0) & ",,,,"
			objIDFTextFile.WriteLine(txt)
   			objVerifyTextFile.Writeline(txt) 
			RecordSet.MoveNext
		Loop
		
		RecordSet.Close
	
	Next
	
	objIDFTextFile.Close
	

Next
objVerifyTextFile.Close	
Connection.Close


WriteLog arrStores,arrSwitchperStore,arrSWitchperCSV,Logfile

End Sub



Sub TimeZone (sTimeZone,arrStores,arrSwitchperCSV,arrCSVfile,Logfile)

Set objRebootTextFile = objFSO.OpenTextFile _
    (Path & sTimeZone & "_Reboot.csv", ForAppending, True)

objRebootTextFile.WriteLine("primaryIPAddress,hostName,scriptField1,scriptField2,scriptField3")

Wscript.Echo "Working on " & sTimeZone & " Reboot list"

For Index = Lbound(arrStores) to Ubound(arrStores)
	strStore = UCase(arrStores(Index))
	
	RecordSet.Open "SELECT  ['MDF-IDF Ip Addresses$'].MDF_IP_addresses , ['MDF-IDF Ip Addresses$'].MDF " & _
	"FROM ['MDF-IDF Ip Addresses$'] WHERE ['MDF-IDF Ip Addresses$'].MDF LIKE '" & strStore & "%'", Connection

	Do While Not RecordSet.EOF

		arrSwitchperStore(Index) = arrSwitchperStore(Index) + 1
		arrSwitchperCSV(Index) = arrSwitchperCSV(Index) & arrCSVfile(0) 
		
		'txt = RecordSet.fields(0) & "," & RecordSet.fields(1) & ",,,"
		txt = RecordSet.fields(0) & ",,,,"
		objRebootTextFile.WriteLine(txt)
	  		
		RecordSet.MoveNext
	Loop

	RecordSet.Close
	
Next


For number = 2 To 11
	
	For Index = Lbound(arrStores) to Ubound(arrStores)
		strStore = UCase(arrStores(Index))


		RecordSet.Open "SELECT  ['MDF-IDF Ip Addresses$'].IDF_" & number & "_IP_Addresses , ['MDF-IDF Ip Addresses$'].IDF_" & number & "_Switches " & _
		"FROM ['MDF-IDF Ip Addresses$'] WHERE ['MDF-IDF Ip Addresses$'].IDF_" & number & "_Switches LIKE '" & strStore & "%'", Connection

		
		Do While Not RecordSet.EOF
			arrSwitchperStore(Index) = arrSwitchperStore(Index) + 1
			arrSwitchperCSV(Index) = arrSwitchperCSV(Index) & ", " & arrCSVfile(number-1) 
			'txt = RecordSet.fields(0) & "," & RecordSet.fields(1) & ",,,"
			txt = RecordSet.fields(0) & ",,,,"
			objRebootTextFile.WriteLine(txt)
   			 
			RecordSet.MoveNext
		Loop
		
		RecordSet.Close
	
	Next
	
	
	

Next
objRebootTextFile.Close	
Connection.Close

WriteLog arrStores,arrSwitchperStore,arrSWitchperCSV,Logfile

End Sub



Sub ResetStoreCounters(StoreArray)

For I = LBound(StoreArray) to UBound(StoreArray)
	StoreArray(I) = 0
	
Next 
End Sub




Sub ClearSwitchperCSV(sSwitchperCSV)

For I = LBound(sSwitchperCSV) to UBound(sSwitchperCSV)
	sSwitchperCSV(I) = ""
	
Next 
End Sub




Function TotalSwitchCounted(StoreArray)

Dim iTemp
iTemp = 0

For I = LBound(StoreArray) to UBound(StoreArray)
	iTemp = iTemp + StoreArray(I) 
	
Next 

TotalSwitchCounted = iTemp

End Function






Sub WriteLog(arrStores,arrSwitchperStore,arrSWitchperCSV,Logfile)

Set objOutputFile = objFSO.OpenTextFile(Path & LogFile, 8, True)

For Index = Lbound(arrStores) to Ubound(arrStores)
	strStore = UCase(arrStores(Index))


	
        objOutputFile.WriteLine(strStore & ": " & arrSwitchperStore(Index) & " of " & ReferenceMainList(strStore)  & vbTab & " :" & arrSwitchperCSV(Index))




Next

objOutputFile.WriteLine(vbCrLf & "Total number of Stacks is " & TotalSwitchCounted(arrSWitchperStore))
objOutputFile.Close

End Sub


Function ReferenceMainList(Store)

Dim strStore, nSwitches
Set Connection2 = CreateObject("ADODB.Connection")
Set RecordSet2 = CreateObject("ADODB.recordset")


With Connection2
   .Provider = "Microsoft.Jet.OLEDB.4.0"
   .ConnectionString = "Data Source=" & Path & "Store IOS Upgrade Rollout Schedule -9-20-07.xls;Extended Properties=Excel 8.0;"
   .Open
End With



	strStore = "T-" & UCase(Right(Store,4))


	RecordSet2.Open "SELECT  ['IOS Rollout Full Schedule$'].Store_Number , ['IOS Rollout Full Schedule$'].Number_SWT_Stacks " & _
	"FROM ['IOS Rollout Full Schedule$'] WHERE ['IOS Rollout Full Schedule$'].Store_Number LIKE '" & strStore & "%'", Connection2


			
		'nSwitches = RecordSet2.fields(1) & "," & RecordSet2.fields(0) 
		nSwitches = RecordSet2.fields(1)

	RecordSet2.Close
	Connection2.Close	
	ReferenceMainList = nSwitches





End Function
