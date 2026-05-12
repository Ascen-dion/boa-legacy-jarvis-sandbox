' --- UFT Driver Script (Action1) ---

' 0. HARDCODED ABSOLUTE PATH (Guarantees location based on your setup)
Dim fso, baseDir
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = "C:\Users\saisr\Downloads\Flight_Legacy_Framework\"

' 1. LOAD THE OBJECT REPOSITORY (.tsr)
RepositoriesCollection.Add baseDir & "ObjectRepository\Flight_OR.tsr"

' 2. LOAD LIBRARIES
ExecuteFile baseDir & "Libraries\Flight_App_Functions.qfl"

' 3. INITIALIZE FILE SYSTEM (For Execution Engine Tracking)
newReqPath = baseDir & "ExecutionEngine\NewRequest\"
workingPath = baseDir & "ExecutionEngine\Working\"
compPath = baseDir & "ExecutionEngine\Completed\"

' 4. CONNECT TO EXCEL TEST DATA
excelPath = baseDir & "TestData\Master_TestData.xlsx"
Set objExcel = CreateObject("Excel.Application")
Set objWB = objExcel.Workbooks.Open(excelPath)
Set objSheet = objWB.Sheets("MasterControl")
rowCount = objSheet.UsedRange.Rows.Count

' 5. MASTER EXECUTION LOOP
For i = 2 To rowCount
    testID = objSheet.Cells(i, 1).Value
    xmlTemplate = objSheet.Cells(i, 2).Value
    executionFlag = objSheet.Cells(i, 3).Value
    
    If UCase(executionFlag) = "YES" Then
        
        ' --- SIMULATE NEW REQUEST TO WORKING ---
        ' Create the tracker file
        Set trackerFile = fso.CreateTextFile(newReqPath & testID & ".txt", True)
        trackerFile.Close
        
        ' Move from NewRequest to Working
        If fso.FileExists(newReqPath & testID & ".txt") Then
            ' Safety check: Delete file in destination if it was left over from a previous run
            If fso.FileExists(workingPath & testID & ".txt") Then
                fso.DeleteFile workingPath & testID & ".txt", True
            End If
            fso.MoveFile newReqPath & testID & ".txt", workingPath & testID & ".txt"
        End If
        
        ' Fetch Test Data Variables
        uName = objSheet.Cells(i, 4).Value
        pWord = objSheet.Cells(i, 5).Value
        fCity = objSheet.Cells(i, 6).Value
        tCity = objSheet.Cells(i, 7).Value
        
        ' --- PARSE XML INPUT FLOW ---
        xmlPath = baseDir & "InputFlowTemplates\" & xmlTemplate
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.Async = False
        xmlDoc.Load(xmlPath)
        Set keywords = xmlDoc.SelectNodes("/TestFlow/Keyword")
        
        ' --- EXECUTE KEYWORDS DYNAMICALLY ---
        For Each kw In keywords
            keywordName = kw.GetAttribute("name")
            
            ' Inject Data into specific functions
            If keywordName = "FlightLogin" Then
                Eval(keywordName & "(""" & uName & """, """ & pWord & """)")
            ElseIf keywordName = "SearchFlight" Then
                Eval(keywordName & "(""" & fCity & """, """ & tCity & """)")
            Else
                Eval(keywordName & "()")
            End If
        Next
        
        ' --- SIMULATE WORKING TO COMPLETED ---
        If fso.FileExists(workingPath & testID & ".txt") Then
            ' Safety check: Delete file in destination if it was left over from a previous run
            If fso.FileExists(compPath & testID & ".txt") Then
                fso.DeleteFile compPath & testID & ".txt", True
            End If
            fso.MoveFile workingPath & testID & ".txt", compPath & testID & ".txt"
        End If
        
    End If
Next

' 6. CLEANUP
objWB.Close False
objExcel.Quit
Set objSheet = Nothing
Set objWB = Nothing
Set objExcel = Nothing
Set fso = Nothing
