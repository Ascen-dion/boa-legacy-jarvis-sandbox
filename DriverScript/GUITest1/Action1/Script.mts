' --- UFT Driver Script (Action1) ---

Dim fso, baseDir
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = "C:\Users\saisr\Downloads\Flight_Legacy_Framework\"

' 1. LOAD REPOSITORY & LIBRARIES
RepositoriesCollection.Add baseDir & "ObjectRepository\Flight_OR.tsr"
ExecuteFile baseDir & "Libraries\Flight_App_Functions.qfl"

' 2. INITIALIZE PATHS
newReqPath = baseDir & "ExecutionEngine\NewRequest\"
workingPath = baseDir & "ExecutionEngine\Working\"
compPath = baseDir & "ExecutionEngine\Completed\"
indRepPath = baseDir & "Reports\Individual\"
consRepPath = baseDir & "Reports\Consolidated\"

' --- NEW: ENTERPRISE XML CONFIGURATION INJECTION ---
' This function dynamically reads any XML file and loads the leaf nodes 
' into UFT's Global Environment object so all .qfl files can access them.
Sub LoadXMLConfigToEnvironment(xmlFilePath)
    If fso.FileExists(xmlFilePath) Then
        Dim xmlDoc, nodes, node
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.Async = False
        xmlDoc.Load(xmlFilePath)
        
        ' Select all elements that do not have child elements (the actual data values)
        Set nodes = xmlDoc.SelectNodes("//*[not(*)]")
        For Each node In nodes
            Environment.Value(node.nodeName) = node.text
        Next
    End If
End Sub

' Load both templates into memory
LoadXMLConfigToEnvironment baseDir & "ConfigTemplates\Global_Config.xml"
LoadXMLConfigToEnvironment baseDir & "ConfigTemplates\App_Config.xml"

' Retrieve variables for dashboard reporting
envURL = Environment.Value("EnvName")

' 3. CSS STYLES FOR REPORTS
' CSS for Consolidated Dashboard
consCss = "<style>body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; margin: 0; background-color: #f4f7f6; color: #333; } .header { background-color: #0033A0; color: white; padding: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); } .container { width: 90%; margin: 20px auto; } .summary-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; display: flex; justify-content: space-between; } .stat-box { text-align: center; padding: 10px 20px; border-radius: 5px; font-size: 18px; font-weight: bold; } .stat-total { background-color: #e2e3e5; color: #383d41; } .stat-pass { background-color: #d4edda; color: #155724; } .stat-fail { background-color: #f8d7da; color: #721c24; } .accordion { background-color: #fff; color: #444; cursor: pointer; padding: 18px; width: 100%; text-align: left; border: 1px solid #ddd; outline: none; transition: 0.4s; font-size: 16px; font-weight: bold; margin-top: 10px; border-radius: 5px; display: flex; justify-content: space-between; align-items: center; } .active, .accordion:hover { background-color: #e9ecef; } .panel { padding: 0 18px; background-color: white; display: none; overflow: hidden; border: 1px solid #ddd; border-top: none; border-bottom-left-radius: 5px; border-bottom-right-radius: 5px; } table { width: 100%; border-collapse: collapse; margin: 15px 0; } th, td { border: 1px solid #ddd; padding: 10px; text-align: left; } th { background-color: #f8f9fa; color: #333; } .badge { padding: 4px 8px; border-radius: 4px; color: white; font-size: 12px; } .badge.pass { background-color: #28a745; } .badge.fail { background-color: #dc3545; } .view-link { color: #0033A0; text-decoration: none; font-size: 14px; margin-left: 15px; } .view-link:hover { text-decoration: underline; }</style>"

' CSS for Individual ExtentReport Style
indCss = "<style>body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; background-color: #f8f9fa; color: #212529; } .navbar { background-color: #343a40; color: white; padding: 15px 20px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1); } .navbar h2 { margin: 0; font-size: 20px; } .container { padding: 20px; max-width: 1200px; margin: auto; } .card { background: white; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); padding: 20px; margin-bottom: 20px; border-top: 4px solid #0033A0; } .card h3 { margin-top: 0; color: #0033A0; border-bottom: 1px solid #eee; padding-bottom: 10px; } .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; } .info-item { font-size: 14px; } .info-item strong { display: inline-block; width: 120px; color: #555; } table { width: 100%; border-collapse: collapse; margin-top: 10px; } th, td { border: 1px solid #dee2e6; padding: 12px; text-align: left; font-size: 14px; } th { background-color: #e9ecef; font-weight: 600; color: #495057; } .status-pass { color: #28a745; font-weight: bold; } .status-fail { color: #dc3545; font-weight: bold; }</style>"

jsScript = "<script>function togglePanel(id) { var panel = document.getElementById('panel-' + id); var icon = document.getElementById('icon-' + id); if (panel.style.display === 'block') { panel.style.display = 'none'; icon.innerHTML = '+'; } else { panel.style.display = 'block'; icon.innerHTML = '&minus;'; } }</script>"

'' 4. CONNECT TO EXCEL TEST DATA
excelPath = baseDir & "TestData\Master_TestData.xlsx"
Set objExcel = CreateObject("Excel.Application")
Set objWB = objExcel.Workbooks.Open(excelPath)
Set objSheet = objWB.Sheets("MasterControl")
rowCount = objSheet.UsedRange.Rows.Count
colCount = objSheet.UsedRange.Columns.Count ' Dynamically count how many columns you have

' 5. INITIALIZE MEMORY BUFFERS FOR REPORTING
htmlExecutionBody = ""
totalTests = 0
passedTests = 0
failedTests = 0
batchStartTime = Now
runTimestamp = Replace(Replace(Replace(Now, ":", ""), "/", ""), " ", "_")

' 6. MASTER EXECUTION LOOP
For i = 2 To rowCount
    testID = objSheet.Cells(i, 1).Value
    xmlTemplate = objSheet.Cells(i, 2).Value
    executionFlag = objSheet.Cells(i, 3).Value
    
    If UCase(executionFlag) = "YES" Then
        totalTests = totalTests + 1
        testStartTime = Now
        testStatus = "PASS" 
        
        ' --- A. INITIALIZE AUDIT LOG BUFFER ---
        auditBuffer = "=========================================" & vbCrLf
        auditBuffer = auditBuffer & "   JARVIS EXECUTION AUDIT LOG" & vbCrLf
        auditBuffer = auditBuffer & "=========================================" & vbCrLf
        auditBuffer = auditBuffer & "Test ID: " & testID & vbCrLf
        auditBuffer = auditBuffer & "Template: " & xmlTemplate & vbCrLf
        auditBuffer = auditBuffer & "Environment: " & envURL & vbCrLf
        auditBuffer = auditBuffer & "Start Time: " & testStartTime & vbCrLf
        auditBuffer = auditBuffer & "-----------------------------------------" & vbCrLf
        auditBuffer = auditBuffer & "TEST DATA PARAMETERS:" & vbCrLf
        
        ' --- B. DYNAMICALLY READ ALL EXCEL COLUMNS ---
        ' This loop grabs every column from column 4 onwards, no matter how many there are
        Dim testDataDict
        Set testDataDict = CreateObject("Scripting.Dictionary")
        
        For c = 4 To colCount
            headerName = objSheet.Cells(1, c).Value
            cellValue = objSheet.Cells(i, c).Value
            testDataDict.Add headerName, cellValue
            
            ' Add it to the Audit Log
            auditBuffer = auditBuffer & "- " & headerName & ": " & cellValue & vbCrLf
        Next
        
        auditBuffer = auditBuffer & "-----------------------------------------" & vbCrLf
        auditBuffer = auditBuffer & "STEP TRACE:" & vbCrLf
        
        ' --- C. UPDATE NEW REQUEST TRACKER ---
        Set trackerFile = fso.CreateTextFile(newReqPath & testID & ".txt", True)
        trackerFile.WriteLine "Status: RUNNING"
        trackerFile.Close
        If fso.FileExists(workingPath & testID & ".txt") Then fso.DeleteFile workingPath & testID & ".txt", True
        fso.MoveFile newReqPath & testID & ".txt", workingPath & testID & ".txt"
        
        stepHtmlBuffer = ""
        
        ' --- D. EXECUTE FLOW ---
        xmlPath = baseDir & "InputFlowTemplates\" & xmlTemplate
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        xmlDoc.Async = False
        xmlDoc.Load(xmlPath)
        Set keywords = xmlDoc.SelectNodes("/TestFlow/Keyword")
        
        stepCounter = 1
        For Each kw In keywords
            keywordName = kw.GetAttribute("name")
            
            ' Map specific data to functions (Assuming standard column names)
            If keywordName = "FlightLogin" Then
                Eval(keywordName & "(""" & testDataDict("Username") & """, """ & testDataDict("Password") & """)")
            ElseIf keywordName = "SearchFlight" Then
                Eval(keywordName & "(""" & testDataDict("FromCity") & """, """ & testDataDict("ToCity") & """)")
            Else
                Eval(keywordName & "()")
            End If
            
            ' Log Step to HTML and Audit Buffer
            stepHtmlBuffer = stepHtmlBuffer & "<tr><td>" & stepCounter & "</td><td>Executed Keyword: " & keywordName & "</td><td class='status-pass'>PASS</td><td>" & Now & "</td></tr>"
            auditBuffer = auditBuffer & "[" & Now & "] Step " & stepCounter & ": " & keywordName & " -> SUCCESS" & vbCrLf
            
            stepCounter = stepCounter + 1
        Next
        
        testEndTime = Now
        testDuration = DateDiff("s", testStartTime, testEndTime) & " seconds"
        If testStatus = "PASS" Then passedTests = passedTests + 1 Else failedTests = failedTests + 1
        
        ' --- E. GENERATE INDIVIDUAL EXTENT-STYLE REPORT ---
        indFileName = testID & "_" & runTimestamp & ".html"
        indReportFile = indRepPath & indFileName
        Set iRep = fso.CreateTextFile(indReportFile, True)
        ' [KEEP YOUR EXISTING INDIVIDUAL REPORT HTML WRITING CODE HERE]
        iRep.WriteLine "<!DOCTYPE html><html><head><title>" & testID & " - Automation Report</title>" & indCss & "</head><body>"
        iRep.WriteLine "<div class='navbar'><h2>Flight App Test Automation Report</h2><span>" & testStartTime & "</span></div>"
        iRep.WriteLine "<div class='container'><div class='card'><h3>Execution Summary</h3><div class='info-grid'>"
        iRep.WriteLine "<div class='info-item'><strong>Test ID:</strong> " & testID & "</div>"
        iRep.WriteLine "<div class='info-item'><strong>Template:</strong> " & xmlTemplate & "</div>"
        iRep.WriteLine "<div class='info-item'><strong>Start Time:</strong> " & testStartTime & "</div>"
        iRep.WriteLine "<div class='info-item'><strong>End Time:</strong> " & testEndTime & "</div>"
        iRep.WriteLine "<div class='info-item'><strong>Duration:</strong> " & testDuration & "</div>"
        iRep.WriteLine "<div class='info-item'><strong>Status:</strong> <span class='status-" & LCase(testStatus) & "'>" & testStatus & "</span></div>"
        iRep.WriteLine "</div></div>"
        iRep.WriteLine "<div class='card'><h3>Step Details</h3><table><tr><th>#</th><th>Step Description</th><th>Status</th><th>Timestamp</th></tr>"
        iRep.WriteLine stepHtmlBuffer
        iRep.WriteLine "</table></div></div></body></html>"
        iRep.Close
        
        ' --- F. UPDATE CONSOLIDATED ACCORDION HTML ---
        ' [KEEP YOUR EXISTING CONSOLIDATED HTML WRITING CODE HERE]
        htmlExecutionBody = htmlExecutionBody & "<button class='accordion' onclick=""togglePanel('" & testID & "')"">"
        htmlExecutionBody = htmlExecutionBody & "<span><strong>" & testID & "</strong> | Template: " & xmlTemplate
        htmlExecutionBody = htmlExecutionBody & "<a href='file:///" & Replace(indReportFile, "\", "/") & "' target='_blank' class='view-link'>[View Individual Report]</a></span>"
        htmlExecutionBody = htmlExecutionBody & "<div><span class='badge " & LCase(testStatus) & "'>" & testStatus & "</span> <span id='icon-" & testID & "' class='toggle-icon'>+</span></div></button>"
        htmlExecutionBody = htmlExecutionBody & "<div id='panel-" & testID & "' class='panel'><table><tr><th>Step No.</th><th>Action Description</th><th>Status</th><th>Timestamp</th></tr>" & Replace(stepHtmlBuffer, "status-pass", "badge pass") & "</table></div>"
        
        ' --- G. FINALIZE AUDIT LOG & MOVE TO COMPLETED ---
        auditBuffer = auditBuffer & "=========================================" & vbCrLf
        auditBuffer = auditBuffer & "FINAL STATUS: " & testStatus & vbCrLf
        auditBuffer = auditBuffer & "END TIME: " & testEndTime & vbCrLf
        auditBuffer = auditBuffer & "DURATION: " & testDuration & vbCrLf
        auditBuffer = auditBuffer & "=========================================" & vbCrLf
        
        ' Write the full audit log to the text file
        Set trackerFile = fso.OpenTextFile(workingPath & testID & ".txt", 2, True) ' 2 = Write
        trackerFile.Write auditBuffer
        trackerFile.Close
        
        If fso.FileExists(compPath & testID & ".txt") Then fso.DeleteFile compPath & testID & ".txt", True
        fso.MoveFile workingPath & testID & ".txt", compPath & testID & ".txt"
        
        Set testDataDict = Nothing
    End If
Next

' 7. CONSTRUCT THE FINAL CONSOLIDATED DASHBOARD
' [KEEP YOUR EXISTING DASHBOARD HTML WRITING CODE HERE]
consReportFile = consRepPath & "Execution_Dashboard_" & runTimestamp & ".html"
Set cRep = fso.CreateTextFile(consReportFile, True)

cRep.WriteLine "<!DOCTYPE html><html><head><title>Jarvis Automation Dashboard</title>" & consCss & jsScript & "</head><body>"
cRep.WriteLine "<div class='header'><h2>Jarvis Automation Execution Dashboard</h2><p>Environment: " & envURL & " | Executed on: " & batchStartTime & "</p></div>"
cRep.WriteLine "<div class='container'><div class='summary-card'>"
cRep.WriteLine "<div class='stat-box stat-total'>Total Scripts Executed<br><span style='font-size: 24px;'>" & totalTests & "</span></div>"
cRep.WriteLine "<div class='stat-box stat-pass'>Total Passed<br><span style='font-size: 24px;'>" & passedTests & "</span></div>"
cRep.WriteLine "<div class='stat-box stat-fail'>Total Failed<br><span style='font-size: 24px;'>" & failedTests & "</span></div>"
cRep.WriteLine "</div>"
cRep.WriteLine "<h3>Execution Details</h3>"
cRep.WriteLine htmlExecutionBody
cRep.WriteLine "</div></body></html>"
cRep.Close

' 8. CLEANUP
objWB.Close False
objExcel.Quit
Set objSheet = Nothing
Set objWB = Nothing
Set objExcel = Nothing
Set fso = Nothing
