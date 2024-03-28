Option Explicit

Dim T: Set T = New Trace
Dim FL: Set FL = New FileList
Dim TB



Sub InitLookup
	Set TB = New DirNameTable
	'TB.Add "Common64",""
	'TB.Add "ProgramFiles64Folder", "Program Files (x64)"
End Sub


Public isGUI, installer, database, message, compParam  'global variables access across functions

Const msiOpenDatabaseModeReadOnly     = 0
Const debugFlag = False


' Check if run from GUI script host, in order to modify display
If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "W" Then isGUI = True

' Connect to Windows Installer object

Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer")
If Err.Number > 0 Then
	WScript.Echo Err.Description
Else
	' Open database
	Dim databasePath:databasePath = Wscript.Arguments(0)
	Set database = installer.OpenDatabase(databasePath, msiOpenDatabaseModeReadOnly)
	InitLookup
	ListComponents True
	CreateReport
End If
Wscript.Quit 0

' List all components in database
Sub ListComponents(queryAll)
	Dim view, record, component, componentId, attributeVal, directory, keyPath, condition
	Dim checkAttr, fileName, fileLocation
	
	 Set view = database.OpenView("SELECT * FROM `Component` ORDER BY `Component`")
	view.Execute
	Do
		Set record = view.Fetch
		If record Is Nothing Then Exit Do
		component = record.StringData(1)
		componentId = record.StringData(2)
		directory = record.StringData(3)
		attributeVal = record.StringData(4)
		condition = record.StringData(5)
		keyPath = record.StringData(6)
		checkAttr = CInt(attributeVal)
		checkAttr = checkAttr And 4
        If checkAttr <> 4 and len(keyPath) <> 0 Then
         	fileName = QueryFileName(component, keyPath)
         	fileLocation = QueryDirectory(component, directory)
          'T.W "FileName: " & fileName & vbTab & "Location: " &  fileLocation & vbTab & "RawDir: " & directory & vbTab & "Condition: " & condition 
          FL.AddFile fileName, componentId, fileLocation, condition
		End If
		T.W "----------------------------------------"
	Loop
End Sub

Function QueryDirectory(component, directory)
	Dim view, record, directory_parent, defaultdir, path, temp, index, yatemp, tstr
		T.W "Query Comp: "& component &" Dir:  = " & directory
      Set view = database.OpenView("SELECT * FROM `Directory` WHERE `Directory` = ?")
      Set compParam = installer.CreateRecord(1)
      Dim TARGETDIR: TARGETDIR="TARGETDIR"
      Dim fcIndex, clIndex, eCount

	Do
        compParam.StringData(1) = directory
        view.Execute compParam
        Set record = view.Fetch
        T.W "SQLParam: Directory = " & directory
        
        If record Is Nothing Then Exit Do
        directory = record.StringData(1)
        directory_parent = record.StringData(2)
        defaultdir = record.StringData(3)
        T.W "SQLResult: " & directory_parent & "\" &  defaultdir
        
        If directory_parent <> TARGETDIR Then
        		T.W "Lookup Parent: " & directory_parent 
        		   		
        		'ignore empty String and "."
        		If Len(defaultdir) > 1 Then
        			temp = GetDirName(defaultdir)
        			T.W "GetDirName:" & temp
	        		If temp <> "" Then
	        			'special treatment for p prefixed dirname?
	        			If temp <> "" Then
	        				If Left(directory,1) = "p" Then
		        				T.W "directory started with with p"
		        				temp = QueryDirectory(component,temp)
		        			End If
		        			
		        		End If
		        		If Not TB.Exists(temp) Then
		        			path = temp & "\" & path
		        		End If
        			End If
'         			If temp <> "" Then
'         				path = temp & "\" & path	
'         			End If
        			T.W "PathAdjusted: " & path
        			
        		End If            
        		directory = directory_parent
        Else
    			path =GetLongDirName(defaultdir) & "\" & path
    			QueryDirectory = path
    			T.W "ResultDir:  = " & path
          Exit Do
        End If
    Loop
	Set view = Nothing
End Function

Function GetLongDirName(defDir)
	Dim  fcIndex
	fcIndex = InStr(defDir,"|")
	If fcIndex > 1 Then
		GetLongDirName = Mid(defDir,fcIndex+1)
	Else
		'remove the .: prefix from the defaultDir
		GetLongDirName = SubString(defDir,":")
	End If
End Function

Function SubString(str, sDelim)
	Dim  fcIndex
	fcIndex = InStr(str,sDelim)
	If fcIndex > 1 Then
		SubString = Mid(str,fcIndex+1)
	Else		
		SubString = str
	End If	
End Function


Function GetDirName(defDir)
	T.W "GetDirName: " & defDir
	Dim clIndex, fcIndex, temp, tstr, res
		res = ""
  		clIndex = InStr(defDir,":")
     If clIndex > 0 Then
     	T.W "GetDirName: ColonFound@" & clIndex
     	temp = Split(defDir, ":", -1, vbTextCompare)
      	tstr =  temp(UBound(temp))      	
      	If tstr <> "." Then
      		T.W "GetDirName: tstr=" & tstr
      		 res = GetLongDirName(tstr)
      	End If
		Else
     	T.W "GetDirName: ColonNotFound"
				res = GetLongDirName(defDir)
     End If
		GetDirName = res
		T.W "GetDirName: res=" & res
End Function

Function QueryFileName(component, keyPath)
	' Get component info and format output header
	Dim view, record, header, fileName, temp
	Set view = database.OpenView("SELECT `FileName` FROM `File` WHERE `File` = ?")
	Set compParam = installer.CreateRecord(1)
	compParam.StringData(1) = keyPath
	view.Execute compParam
	Set record = view.Fetch
	Set view = Nothing
	If record Is Nothing Then Fail "File is not in database: " & component
	fileName = record.StringData(1)
	Dim idx : idx = InStr(fileName,"|")
	If idx > 1 Then
		fileName = Mid(fileName,idx+1)
	End If
	QueryFileName = fileName
End Function

Sub CreateReport()
	Dim fN, fO, fLs
	Dim hw, str, idx
	Set hw = New HTMLWriter
	hw.NewTable "RetailFiles"
	hw.SetHeader Array("FileName", "ComponentID","InstallLocation","InstallCondition")
	
	idx = 0
	For Each fN In FL.GetFileNames
		hw.OpenRow ((idx Mod 2) = 1)
		Set fO = FL.GetFileInfo(fN)
		hw.AddRowData fN
		hw.AddRowData Join(fO.GetIDs,"<br>")
		hw.AddRowData2 Join(fO.GetLocations,"<br>"), "nowrap"
		hw.AddRowData2 Join(fO.GetConditions,"<br>"), "nowrap"
		hw.CloseRow
		idx = idx +1
	Next
	hw.CloseTable
	hw.W "<br>FileCount: " & idx
	hw.OpenFile
	Set hw = Nothing
End Sub


Class DirNameTable
	Dim dirDict
	Private Sub Class_Initialize()
		Set dirDict =CreateObject( "Scripting.Dictionary")
	End Sub
	
	Public Sub Add(dirName, dirActual)
		If Not dirDict.Exists(dirName) Then
			dirDict.Add dirName, dirActual
			T.W "TableLookup Added: " & dirName & "," & dirActual
		End If
	End Sub
	Public Function Exists(dirName)
		Exists = dirDict.Exists(dirName)
	End Function
	Public Function Name(dirName)
			GetName = dirDict.Item(dirName)
	End Function
	
	Private Sub Class_Terminate()
   	Set dirDict = Nothing
	End Sub
End Class

'**************************************
'RetailFileList 
'**************************************
Class FileList
	Dim filesDict
	Private Sub Class_Initialize()
   	Set filesDict =CreateObject( "Scripting.Dictionary")

	End Sub
	
	Public Sub AddFile(fileName, compId, fileLocation, fileCondition)
		Dim f
   	If filesDict.Exists(fileName) Then
   		Set f = filesDict.Item(fileName)
   		'T.W fileName & " exists: Adding " & fileLocation
   	Else
   		Set f = New FileInfo
   		f.Name = fileName
   		filesDict.Add fileName, f
   		'T.W fileName & " added. With fileLocation: " & fileLocation
   	End If		
   	f.AddLocation fileLocation
   	f.AddCondition fileCondition
   	f.AddId compId
	End Sub
	
	Public Function GetFileNames()
		GetFileNames = filesDict.Keys
	End Function
	
	Public Function GetFileInfo(fileName)
		Set GetFileInfo = filesDict.Item(fileName)
	End Function
	
	Private Sub Class_Terminate()
   	Set filesDict = Nothing
	End Sub
End Class

'**************************************
'FileInfo 
'**************************************
Class FileInfo
	Public Name
	Dim aLocation, aCondition, aID
	Private Sub Class_Initialize()
   	ReDim aLocation(0)
   	ReDim aCondition(0)
   	ReDim aID(0)
	End Sub
	Public Sub AddLocation(installLocation)
		AddToArray aLocation,installLocation
	End Sub
	Public Sub AddCondition(installCondition)
		AddToArray aCondition,installCondition
	End Sub
	Public Sub AddId(componentId)
		AddToArray aID,componentId
	End Sub

	Private Sub AddToArray(tgArray, nData)
		Dim n: n = UBound(tgArray)
		ReDim Preserve tgArray(n+1)
		tgArray(n) = nData 
	End Sub
	Public Function GetLocations()
		GetLocations = aLocation
	End Function
	Public Function GetConditions()
		GetConditions = aCondition
	End Function
  Public Function GetIDs()
		GetIDs = aID
	End Function
	
End Class

'**************************************
'Trace 
'**************************************
Class Trace
	Dim fso,  out, path
	
	Private Sub Class_Initialize()
   	Set fso =CreateObject( "Scripting.FileSystemObject")
	End Sub
	
	Public Sub SetPath(filePath)
		path = filePath
	End Sub
	
	Public Sub W(text)
		If IsEmpty(out) Then
			If IsEmpty(path) Then 
				path = fso.GetFolder("./") & "\" & WScript.ScriptName & ".log"
			End If
			Set out = fso.CreateTextFile(path , True)
		End If 		
		out.WriteLine Now() & vbTab & text
	End Sub
	
	Private Sub Class_Terminate()
		out.Close
   	Set fso = Nothing
   	Set out = Nothing
	End Sub
End Class


'**************************************
'HTMLWriter 
'**************************************
Class HTMLWriter
	Dim fso,  out, path
	
	Private Sub Class_Initialize()
   	Set fso =CreateObject( "Scripting.FileSystemObject")
	End Sub
	
	Private Sub Class_Terminate()
		If Not IsEmpty(out) Then
			W "</body></html>"
			out.Close
		End If
   	Set fso = Nothing
   	Set out = Nothing
	End Sub
	
	Public Property Get FilePath
   	FilePath = path
   End Property
   
   Public Property Let FilePath(fPath)
   	path = fPath
   End Property
	
	Public Sub W(text)
		If IsEmpty(out) Then
			If IsEmpty(path) Then 
				path = fso.GetFolder("./") & "\" & WScript.ScriptName & ".html"
			End If
			Set out = fso.CreateTextFile(path , True)
			OpenBody
		End If 		
		out.WriteLine text
	End Sub
	
	Public Sub SetHeader(txtArray)
		JT  "tr"
		Dim s
		For Each s In  txtArray
			ST "th", s
		Next
		CT "tr"
	End Sub

	Public Sub AddRowData(txt)
		ST "td", txt
	End Sub
	
	Public Sub AddRowData2(txt, thAttr)
		TT "td", thAttr, txt
	End Sub
	
	Public Sub OpenRow( bZebra)
		If bZebra Then
			OT  "tr","bgcolor='#333333'"
		Else
			JT  "tr"
		End If
	End Sub
	
	Public Sub CloseRow()
		CT "tr"
	End SUb
	
	Public Sub CloseTable()
		W "</table>"
	End Sub
   	
	Public Sub NewTable(tableName)
		W ToToggleLink(tableName)
		OT "table", "id='" & tableName & "' class='collapse'"
	End Sub	
	
	Private Sub OpenBody()
	JT "html"
	JT "head"
	ST "title", WScript.ScriptName
	OT "style", "type='text/css'"
	W "A{ border: thin solid;	font-family: Arial, Tahoma, 'Trebuchet MS';	clear: none;	font: bold;" 
	W "font-variant: small-caps;	color: #FF8C00;	border-top-width: 1px;	border-top: none;" 
	W "border-right: none;	border-left: none;	border-bottom: none;	text-decoration: none;	}"
	W "BODY{ PADDING-RIGHT: 0px; PADDING-LEFT: 10px; PADDING-BOTTOM: 0px; MARGIN: 0px;"
   W "COLOR: #F1F1F1; PADDING-TOP: 10px; BACKGROUND-COLOR: Black;  font-size: smaller;" 
   W " font-family: 'Arial Narrow', Tahoma;} "
  W "TABLE{	border-left: 1px solid;	border-right: 1px solid;	border-bottom: 1px solid; "
  	 W "	border-top: 1px solid;	background-color: #111111;}"
	W "TH {	font: Tahoma, 'Trebuchet MS', Arial; font-size: smaller;	border: none; " 
	W "background: #3F3F3F;	border-left: 0px;	border-right: 0px;}"
	W "TD {	font: Tahoma, 'Trebuchet MS', Arial;	font-size: smaller;	border: none;	padding-left: 4px;	padding-right: 4px;	}"
	W "h4{	background-color: #393939;}"
	W ".collapse { position: absolute; visibility: hidden; }"
	W ".expand { position: relative; visibility: visible; }"
	CT "style"
	JT "script"
 	W "function T(tableName) {"
  	W " var temp = document.getElementById(tableName);"
	W " state = temp.style.visibility;"
	W " if(state == 'visible' || state == 'show' ){"
  	W "  temp.style.position = 'absolute';"
  	W "  temp.style.visibility = 'hidden';"
  	W " }else{"
  	W "  temp.style.position = 'relative';"
  	W "  temp.style.visibility = 'visible';"
	W " }"
  	W "}"
	CT "script"
	CT "head"
	JT "body"
	End Sub

	Private Sub ST(tagName, txt)
		TT tagName, "", txt
	End Sub
	
	Private Sub TT(tagName, attr, txt)
		OT tagName, attr
		W txt
		CT tagName		
	End Sub
	
	Private Sub JT(tagName)
		W "<" & tagName & ">"
	End Sub
	
	Private Sub OT(tagName, attr)
		If attr = "" Then
			JT tagName
		Else
			JT tagName & " " & attr			
		End If
	End Sub
	
	Private Sub CT(tagName)
			W "</" & tagName & ">"
	End Sub
	
	Public Function ToToggleLink(dName)
		ToToggleLink= A(dName,"javascript:T('" & dName & "')") 
	End Function
	
	Function A(txt, link)
		A	 = "<a href=""" & link & """>" & txt & "</a>"
	End Function
	
	Public Sub OpenFile()
		RunCommand path, False
	End Sub

	Private Sub RunCommand(cmd,bWait)
		Dim SHL: Set SHL = CreateObject("WScript.Shell")
		SHL.Run cmd, 0, bWait
		Set SHL = Nothing
	End Sub 
End Class