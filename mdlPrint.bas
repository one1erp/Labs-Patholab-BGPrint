Attribute VB_Name = "mdlPrint"
Option Explicit




Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'related to the Mutex handling:
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long

Public gCon As ADODB.Connection
Public gWordReport As WordReport.Report
Public gGetProcess As GetProcess.clsGetProcess

Public gSdgLog As Object 'SdgLog.CreateLog

Public Sub ExecutePrint()
    Dim lstSQL As String
    Dim lstComputerName As String
    Dim lloStrLen As Long
    Dim lrsPrints As New ADODB.Recordset
    Dim i As Integer

    'the answer returned from the waiting for the Mutex:
    Dim lWaitAnswer As Long
    
    On Error GoTo ErrHnd
    
    'wait to get the Mutex if taken:
    lWaitAnswer = WaitForSingleObject(frmBGPrint.lMutexHandle, 30000)
    
    'exit if didn't get the mutex:
    If lWaitAnswer <> 0 Then
        Exit Sub
    End If
    
    ' get the workstation name
    lloStrLen = 50
    lstComputerName = Space(50)
    Call GetComputerName(lstComputerName, lloStrLen)
    lstComputerName = Left(lstComputerName, lloStrLen)
    
    ' select the printing jobs for this workstation
    lstSQL = "SELECT * " & _
             "FROM lims.bg_print " & _
             "WHERE workstation_name  = '" & lstComputerName & "' " & _
             "ORDER BY created_on ASC,DOC_ID asc"
             'roy                    /\added doc id
    Set lrsPrints = gCon.Execute(lstSQL)
    
    ' loop over the printings and call the PrintSdg Procedure
    frmBGPrint.Caption = "מדפיס..."
    frmBGPrint.Refresh
    If Not lrsPrints.EOF Then
        Set gWordReport = New WordReport.Report
        Set gWordReport.Connection = gCon
        Set gGetProcess = New GetProcess.clsGetProcess
    End If
    i = 0
    While Not lrsPrints.EOF
                         
        Select Case UCase(nte(lrsPrints("REPORT_TYPE")))
        Case "WF", ""
            If PrintSdg(lrsPrints("SDG_ID"), lrsPrints("WORKFLOW_NODE_ID")) Then
            
                ' delete the printing immediatly after it executed
                lstSQL = "DELETE lims.bg_print " & _
                         "WHERE  (REPORT_TYPE = 'WF' OR  REPORT_TYPE IS NULL)" & _
                         "  AND  sdg_id = " & lrsPrints("SDG_ID") & _
                         "  AND  workflow_node_id = " & lrsPrints("WORKFLOW_NODE_ID") & _
                         "  AND  created_on = to_date('" & lrsPrints("CREATED_ON") & "','dd/mm/yyyy hh24:mi:ss') " & _
                         "  AND  workstation_name  = '" & lstComputerName & "' "
                gCon.Execute lstSQL
            End If
        Case "DIRECT"
            If PrintDirect(lrsPrints("DOC_ID"), lrsPrints("REPORT_ID")) Then
            
                ' delete the printing immediatly after it executed
                lstSQL = "DELETE lims.bg_print " & _
                         "WHERE  REPORT_TYPE = 'DIRECT' " & _
                         "  AND  report_id = " & lrsPrints("REPORT_ID") & _
                         "  AND  doc_id = " & lrsPrints("DOC_ID") & _
                         "  AND  created_on = to_date('" & lrsPrints("CREATED_ON") & "','dd/mm/yyyy hh24:mi:ss') " & _
                         "  AND  workstation_name  = '" & lstComputerName & "' "
                gCon.Execute lstSQL
                
                lstSQL = "DELETE lims.bg_print_params " & vbCrLf & _
                         "WHERE  doc_id = " & lrsPrints("DOC_ID")
                gCon.Execute lstSQL
            End If
        Case Else
            frmBGPrint.lstMsg.RemoveItem (0)
            frmBGPrint.lstMsg.AddItem "שגיאה בסוג דו""ח"
            frmBGPrint.lstMsg.ListIndex = frmBGPrint.lstMsg.ListCount - 1
            frmBGPrint.Refresh
        End Select

        lrsPrints.MoveNext
        i = i + 1
        If (i Mod 200) = 0 Then
            Open "C:\BG_Print_" & Format(Date, "yyyymmdd") & ".log" For Append As #4
            Print #4, Date & " " & Time
            Print #4, "Number Of Docs Printed = "; i
            Print #4, "Process Mem Usage = " & gGetProcess.GetProcesses("winword.exe")
            Close #4
        End If
    Wend
    Set gWordReport = Nothing
    Set gGetProcess = Nothing

    frmBGPrint.Caption = "סיים הדפסות."
    frmBGPrint.Refresh
                         
    lrsPrints.Close
    
    
    'let others have the Mutex
    ReleaseMutex (frmBGPrint.lMutexHandle)
     
    Exit Sub
    
ErrHnd:
    
    With frmBGPrint
        With .lstMsg
            .RemoveItem (0)
            .AddItem "שגיאה: " & Err.Description
            .ListIndex = .ListCount - 1
        End With
    End With
    
    Set gWordReport = Nothing
    Set gGetProcess = Nothing
    'let others have the Mutex
    ReleaseMutex (frmBGPrint.lMutexHandle)
End Sub

Private Function PrintSdg(pdoSdgID As Double, pdoWorkflowNodeID As Double) As Boolean
    
    Dim lrst As ADODB.Recordset
    Dim queries As ADODB.Recordset
    Dim wf As ADODB.Recordset
    Dim Dest As ADODB.Recordset
    Dim lstErrMsg As String
    Dim sql As String
    
    
    On Error GoTo ErrHnd
    
    PrintSdg = False
    
    frmBGPrint.lstMsg.RemoveItem (0)
    frmBGPrint.lstMsg.AddItem "דרישה: " & pdoSdgID & " (" & Time & ")"
    frmBGPrint.lstMsg.ListIndex = frmBGPrint.lstMsg.ListCount - 1
    frmBGPrint.Refresh
    
    Call gSdgLog.InsertLog(CLng(pdoSdgID), "BGP.SELECT", "")
    
    Set lrst = gCon.Execute("SELECT * FROM Sdg WHERE SDG_ID=" & pdoSdgID)
    
    sql = " select wf.name, wfn.long_name"
    sql = sql & " from lims_sys.workflow wf, "
    sql = sql & "      lims_sys.workflow_node wfn"
    sql = sql & " where wfn.WORKFLOW_ID=wf.WORKFLOW_ID"
    sql = sql & " and   wfn.WORKFLOW_NODE_ID=lims.Get_Parent_Event_Id('" & pdoWorkflowNodeID & "')"
    Set wf = gCon.Execute(sql)
    
'    Set wf = gCon.Execute("select wf.name, pwfn.long_name " & _
        "from lims_sys.workflow wf, lims_sys.workflow_node wfn, lims_sys.workflow_node pwfn " & _
        "where wf.workflow_id = wfn.workflow_id and " & _
        "pwfn.workflow_node_id = wfn.parent_id and " & _
        "wfn.workflow_node_id = " & pdoWorkflowNodeID)
        
    Set queries = gCon.Execute("select u_query, u_query_name, u_word_template, u_wreport_user.u_wreport_id " & _
        "from lims_sys.u_wreport_query_user, lims_sys.u_wreport_user " & _
        "where u_wreport_query_user.u_wreport_id = u_wreport_user.u_wreport_id and " & _
        "u_wreport_user.u_workflow_name = '" & wf("NAME") & "' and " & _
        "u_wreport_user.u_workflow_event = '" & wf("LONG_NAME") & "'")
    wf.Close
    Set Dest = gCon.Execute("select * from lims_sys.u_wrdestination_user du where du.u_wreport_id = " & queries("U_WREPORT_ID"))
    gWordReport.SetTemplate (queries("U_WORD_TEMPLATE"))
    gWordReport.RemoveQueries
    While Not queries.EOF
        Call gWordReport.AddQuery(queries("U_QUERY_NAME"), queries("U_QUERY"))
        queries.MoveNext
    Wend
    queries.Close
    gWordReport.RemoveParameters
    Call gWordReport.AddParameter("SDG_ID", lrst("SDG_ID"))
    While Not Dest.EOF
        If Dest("U_TYPE") = "F" Then
            Call gWordReport.SaveReport(nvl(Dest("U_DEVICE_NAME"), ""))
        Else
            Call gWordReport.PrintReport(nte(Dest("U_DEVICE_NAME")), nvl(Dest("U_COPIES"), 0))
        End If
        Dest.MoveNext
    Wend
    Dest.Close
    
    gWordReport.EndReport
    
    Call gSdgLog.InsertLog(CLng(pdoSdgID), "BGP.PRINT", "")
    
    PrintSdg = True
    Exit Function
    
ErrHnd:

    lstErrMsg = IIf(gWordReport.ErrMsg <> "", gWordReport.ErrMsg & vbCrLf, "") & _
                "SDG_ID: " & pdoSdgID & vbCrLf & _
                "SUB: PrintSdg" & vbCrLf & Err.Description
    
    gWordReport.EndReport
    gWordReport.StartWordAppl
    Open "C:\BG_Print_" & Format(Date, "yyyymmdd") & ".log" For Append As #4
    Print #4, Date & " " & Time
    Print #4, "SDG ID = " & pdoSdgID
    Print #4, lstErrMsg
    Close #4
    With frmBGPrint
        With .lstMsg
            .RemoveItem (0)
            .AddItem "שגיאה: " & lstErrMsg
            .ListIndex = .ListCount - 1
        End With
    End With
    PrintSdg = False
    
End Function

Private Function PrintDirect(pdoDocID As Double, pdoReportID As Double) As Boolean

    Dim sql As String
    Dim QueriesParams As ADODB.Recordset
    Dim queries As ADODB.Recordset
    Dim Dest As ADODB.Recordset
    Dim Template As String
    Dim i As Integer
    
    Dim lstErrMsg As String

    On Error GoTo ErrHnd
    
    PrintDirect = False
    
    frmBGPrint.lstMsg.RemoveItem (0)
    frmBGPrint.lstMsg.AddItem "דו""ח: " & pdoReportID & " (" & Time & ")"
    frmBGPrint.lstMsg.ListIndex = frmBGPrint.lstMsg.ListCount - 1
    frmBGPrint.Refresh
    
    Call gSdgLog.InsertLog(-1, "BGP.SELECT", "Direct report: " & CStr(pdoReportID))

    Set queries = gCon.Execute("select u_query, u_query_name, u_word_template, u_wreport_user.u_wreport_id " & _
        "from lims_sys.u_wreport_query_user, lims_sys.u_wreport_user " & _
        "where u_wreport_query_user.u_wreport_id = u_wreport_user.u_wreport_id and " & _
        "u_wreport_user.u_wreport_id = " & pdoReportID)

    Template = queries("U_WORD_TEMPLATE")
    
    gWordReport.RemoveQueries
    While Not queries.EOF
        Call gWordReport.AddQuery(queries("U_QUERY_NAME"), queries("U_QUERY"))
        queries.MoveNext
    Wend
    queries.Close
    
    sql = "SELECT param_name, param_value " & vbCrLf & _
          "FROM lims.bg_print_params " & vbCrLf & _
          "WHERE doc_id = " & pdoDocID

    Set QueriesParams = gCon.Execute(sql)
    
    If QueriesParams.EOF Then
        lstErrMsg = "אין פרמטרים" & vbCrLf & pdoReportID & " <-> " & pdoDocID
        GoTo ErrHnd
    End If
    
    gWordReport.RemoveParameters
    gWordReport.SetTemplate (Template)
    
    While Not QueriesParams.EOF
        Call gWordReport.AddParameter(QueriesParams("param_name"), QueriesParams("param_value"))
        QueriesParams.MoveNext
    Wend
    QueriesParams.Close
    Set Dest = gCon.Execute("select * from lims_sys.u_wrdestination_user du where du.u_wreport_id = " & pdoReportID)
    While Not Dest.EOF
        If Dest("U_TYPE") = "F" Then
            Call gWordReport.SaveReport(nvl(Dest("U_DEVICE_NAME"), ""))
        Else
            Call gWordReport.PrintReport(nvl(Dest("U_DEVICE_NAME"), ""), nvl(Dest("U_COPIES"), 0))
        End If
        Dest.MoveNext
    Wend
    Dest.Close
    
    gWordReport.EndReport
    
    Call gSdgLog.InsertLog(-1, "BGP.PRINT", "Direct report: " & CStr(pdoReportID))
    
    PrintDirect = True
    Exit Function
    
ErrHnd:
    
    lstErrMsg = IIf(gWordReport.ErrMsg <> "", gWordReport.ErrMsg & vbCrLf, "") & _
                "Doc ID = " & pdoDocID & " - ReportID = " & pdoReportID & vbCrLf & _
                "SUB: PrintDirect" & vbCrLf & Err.Description

    gWordReport.EndReport
    gWordReport.StartWordAppl
    Open "C:\BG_Print_" & Format(Date, "yyyymmdd") & ".log" For Append As #4
    Print #4, Date & " " & Time
    Print #4, "Doc ID = " & pdoDocID & " - ReportID = " & pdoReportID
    Print #4, lstErrMsg
    Close #4
    With frmBGPrint
        With .lstMsg
            .RemoveItem (0)
            .AddItem "שגיאה: " & lstErrMsg
            .ListIndex = .ListCount - 1
        End With
    End With
    PrintDirect = False

End Function

Public Sub DelPrints()
    Dim lstSQL As String
    Dim lstComputerName As String
    Dim lloStrLen As Long
    Dim lrsPrints As New ADODB.Recordset
    
    ' get the workstation name
    lloStrLen = 50
    lstComputerName = Space(50)
    Call GetComputerName(lstComputerName, lloStrLen)
    lstComputerName = Left(lstComputerName, lloStrLen)
    
    ' select the printing jobs for this workstation
    lstSQL = "DELETE lims.bg_print " & _
             "WHERE workstation_name  = '" & lstComputerName & "' "
    Call gCon.Execute(lstSQL)
End Sub

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

Private Function nvl(e As Variant, v As Variant) As Variant
    nvl = IIf(IsNull(e), v, e)
End Function


