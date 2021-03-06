Option Compare Database
Option Explicit

Public Sub frmControlPanel_Close(ByVal currentDash As String)
    Dim formstring As String
    
    formstring = "Display Form (" & currentDash & ")"
    Call CreateCumulativePlan
    On Error Resume Next 'ignore errors if the form we are trying to refresh isn't open
    Forms(formstring).Refresh
    Forms![Display Form (TEAM)].Refresh
    On Error GoTo 0
End Sub

'This function exports all the current part numbers as a text file
Public Sub ExportPNsToTxt()
    Dim rs As DAO.Recordset
    Dim SQL, PNsPath As String
    
    'Sql code for query to determine current pns
    SQL = "SELECT tblDemandInput.PN" & _
          " FROM tblDemandInput" & _
          " WHERE (((tblDemandInput.PN) Not Like '*N/A*'));"
    
    'set record set to query defined by sql
    Set rs = CurrentDb.OpenRecordset(SQL, dbOpenSnapshot)
    
    'determine path for PN export
    PNsPath = CurrentProject.path & "\Scripts\PartNumbers.txt"
    
    Open PNsPath For Output As #1
    
    Do While Not rs.EOF
        Print #1, rs!PN
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Close #1
    
End Sub




'This function waits for the readyforrefresh flag to be true
Public Function refreshformwait()
    Dim RefreshFlag As Boolean
    Dim Twait, Tnow As Date
    Dim i As Integer

    RefreshFlag = DLookup("ReadyForRefresh", "tblReadyForRefresh", "ID = 1")
    
    'While the RefreshFlag is false whait
    i = 0
    Do While RefreshFlag = False
        RefreshFlag = DLookup("ReadyForRefresh", "tblReadyForRefresh", "ID = 1")
    Loop
End Function


Public Sub TurnOffSubDataSheets()
    Dim MyDB As DAO.Database
    Dim MyProperty As DAO.Property
    Dim i, intChangedTables As Integer
    
    Dim propName As String
    Dim propType As Integer
    Dim propVal As String
    
    Dim strS As String
    
    Set MyDB = CurrentDb
    
    propName = "SubDataSheetName"
    propType = 10
    propVal = "[NONE]"
    
    On Error Resume Next
    
    For i = 0 To MyDB.TableDefs.Count - 1
    
        If (MyDB.TableDefs(i).Attributes And dbSystemObject) = 0 Then
        
            If MyDB.TableDefs(i).Properties(propName).Value < propVal Then
                MyDB.TableDefs(i).Properties(propName).Value = propVal
                intChangedTables = intChangedTables + 1
            End If
            
            If Err.Number = 3270 Then
                Set MyProperty = MyDB.TableDefs(i).CreateProperty(propName)
                MyProperty.Type = propType
                MyProperty.Value = propVal
                MyDB.TableDefs(i).Properties.Append MyProperty
            Else
            If Err.Number < 0 Then
                MsgBox "Error: " & Err.Number & " on Table " _
                & MyDB.TableDefs(i).Name & "."
                MyDB.Close
                Exit Sub
            End If
        End If
        
        End If
    Next i
    


End Sub

Public Function StoreSummaryStats(ByVal ProductLine As String)
    
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryStoreSummaryStats (" & ProductLine & ")"
    DoCmd.SetWarnings True

End Function


Public Function WeekDays(ByVal startDate As Date, ByVal endDate As Date) As Integer
   ' Returns the number of weekdays in the period from startDate
    ' to endDate not inclusive of start day. Returns blank if an error occurs.
    ' If your weekend days do not include Saturday and Sunday and
    ' do not total two per week in number, this function will
    ' require modification.
    On Error GoTo Weekdays_Error
    
    ' The number of weekend days per week.
    Const ncNumberOfWeekendDays As Integer = 2
    
    ' The number of days inclusive.
    Dim varDays As Variant
    
    ' The number of weekend days.
    Dim varWeekendDays As Variant
    
        
    ' Calculate the number of days not inclusive of start day
    varDays = DateDiff(Interval:="d", _
        date1:=startDate, _
        Date2:=endDate)
    
    
    ' Calculate the number of weekend days.
    varWeekendDays = (DateDiff(Interval:="ww", _
        date1:=startDate, _
        Date2:=endDate) _
        * ncNumberOfWeekendDays) _
        + IIf(DatePart(Interval:="w", _
        Date:=startDate) = vbSunday, 1, 0) _
        + IIf(DatePart(Interval:="w", _
        Date:=endDate) = vbSaturday, 1, 0)
    
    ' Calculate the number of weekdays.
    WeekDays = (varDays - varWeekendDays)
    
Weekdays_Exit:
    Exit Function
    
Weekdays_Error:
    WeekDays = ""
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Weekdays"
    Resume Weekdays_Exit
End Function

'==========================================================
' The DateAddW() function provides a workday substitute
' for DateAdd("w", number, date). This function performs
' error checking and ignores fractional Interval values.
'==========================================================
Function DateAddW(ByVal TheDate, ByVal Interval)
 
   Dim Weeks As Long, OddDays As Long, temp As String
 
   If VarType(TheDate) <> 7 Or VarType(Interval) < 2 Or _
              VarType(Interval) > 5 Then
      DateAddW = TheDate
   ElseIf Interval = 0 Then
      DateAddW = TheDate
   ElseIf Interval > 0 Then
      Interval = Int(Interval)
 
   ' Make sure TheDate is a workday (round down).
 
      temp = Format(TheDate, "ddd")
      If temp = "Sun" Then
         TheDate = TheDate - 2
      ElseIf temp = "Sat" Then
         TheDate = TheDate - 1
      End If
 
   ' Calculate Weeks and OddDays.
 
      Weeks = Int(Interval / 5)
      OddDays = Interval - (Weeks * 5)
      TheDate = TheDate + (Weeks * 7)
 
  ' Take OddDays weekend into account.
 
      If (DatePart("w", TheDate) + OddDays) > 6 Then
         TheDate = TheDate + OddDays + 2
      Else
         TheDate = TheDate + OddDays
      End If
 
      DateAddW = TheDate
    Else                         ' Interval is < 0
      Interval = Int(-Interval) ' Make positive & subtract later.
 
   ' Make sure TheDate is a workday (round up).
 
      temp = Format(TheDate, "ddd")
      If temp = "Sun" Then
         TheDate = TheDate + 1
      ElseIf temp = "Sat" Then
         TheDate = TheDate + 2
      End If
 
   ' Calculate Weeks and OddDays.
 
      Weeks = Int(Interval / 5)
      OddDays = Interval - (Weeks * 5)
      TheDate = TheDate - (Weeks * 7)
 
   ' Take OddDays weekend into account.
 
      If (DatePart("w", TheDate) - OddDays) > 2 Then
         TheDate = TheDate - OddDays - 2
      Else
         TheDate = TheDate - OddDays
      End If
 
      DateAddW = TheDate
    End If
 
End Function

Public Function DaysInMonth(Optional dtmDate As Date = 0) As Integer
    ' Return the number of days in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    DaysInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 1) - _
     DateSerial(Year(dtmDate), Month(dtmDate), 1)
End Function

'Creates the cumulative day month plan
Public Sub CreateCumulativePlan()
    Dim WorkingDays, totDays, DemandInput, ReqPerDay, _
        ReqPerDayRemainder, ReqPerDayNew, i, ExtraUnit, ReqDlvTot As Integer
    Dim DB As Object
    Dim rs As DAO.Recordset
    Dim SQL, Program, ShortDate As String
    
    Set DB = CurrentDb
    Set rs = DB.OpenRecordset("qryMonthPlan")
    
    
    'Delete records
    SQL = "DELETE tblCumulativePlan.* FROM tblCumulativePlan;"
    DB.Execute SQL

    
    totDays = DaysInMonth
    
    
    'Calc number of working days in month (number of weekdays between two dates plus one)
    WorkingDays = WeekDays(DateSerial(Year(Date), Month(Date), 1), DateSerial(Year(Date), Month(Date), totDays)) + 1
    
    'For each program in qryMonthPlan...
    With rs
        If Not .BOF And Not .EOF Then
            While (Not .EOF)
                Program = rs.Fields("Program")
                
                'DemandInput, ReqPerDay, ReqPerDayRemainder
                DemandInput = rs.Fields("PlanQTYTot")
                ReqPerDay = Int(DemandInput / WorkingDays)
                ReqPerDayRemainder = DemandInput Mod WorkingDays
                
                'For each day in the month...1 to totDays
                ReqDlvTot = 0 'set to zero
                For i = 1 To totDays
                    
                    'calc ShortDate
                    ShortDate = Month(Date) & "/" & (i) & "/" & Year(Date)
                    
                    'Enter required complete per day not including weekends
                    If Weekday(ShortDate) = "1" Or Weekday(ShortDate) = "7" Then
                        ReqPerDayNew = 0
                    Else
                        'we will add one additional unit every working day until the remainder is equal to zero
                        If ReqPerDayRemainder = 0 Then
                            ExtraUnit = 0
                        Else
                            ExtraUnit = 1
                            'subtract 1 from ShipPerDayRemainder
                            ReqPerDayRemainder = ReqPerDayRemainder - 1
                        End If
                        ReqPerDayNew = ReqPerDay + ExtraUnit
                    End If
                    ReqDlvTot = ReqDlvTot + ReqPerDayNew
                    'Send info to tblCumulativePlan
                    SQL = "INSERT INTO tblCumulativePlan (shortdate, ReqDlvTot, Program ) " & _
                          "SELECT #" & ShortDate & "#, " & ReqDlvTot & ", '" & Program & "';"
                    DB.Execute SQL
               Next i
                .MoveNext
            Wend
        End If
    End With
    'MsgBox WorkingDays
    
    'DB.Execute (sql)
    Set DB = Nothing
    Set rs = Nothing
    
End Sub


'**************  End of Code **************

'---------------------------------------------------------
' Create Query functions
'---------------------------------------------------------


'This function generates the qryOp queries that are used to sort the product locations
'This funciton is only called during creation or re-creation of the queries during
'development
Public Function createqryOp(ByVal opNum As Integer, ByVal opName As String, ByVal opProgram As String, ByVal IsHospital As Boolean, Optional OpNames As Variant)
    
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    Dim element As Variant
    
    'Define title for qryName
    'use "qryOp_opNum_opName (opProgram)"
    If IsHospital = True Then
        qryName = "qryOP_Hospital (" & opProgram & ")"
    Else
        qryName = "qryOP_" & opNum & "_" & opName & " (" & opProgram & ")"
    End If
    
    
    'Define SQL string
    SQL = "SELECT [qryRawData (" & opProgram & ")].Order," & _
        " [qryRawData (" & opProgram & ")].[SN]," & _
        " [qryRawData (" & opProgram & ")].Material," & _
        " [qryRawData (" & opProgram & ")].[Bscstart]," & _
        " [qryRawData (" & opProgram & ")].[Basicfin]," & _
        " [qryRawData (" & opProgram & ")].[Minutes Left Till Due]," & _
        " [qryRawData (" & opProgram & ")].[TimeRemaining]," & _
        " [qryRawData (" & opProgram & ")].[Flag Set]," & _
        " [Material] & ' ' & [SN] AS FormDisplay," & _
        " [qryRawData (" & opProgram & ")].[OrderType]," & _
        " [qryRawData (" & opProgram & ")].[ActRelease]," & _
        " Cdbl([qryRawData (" & opProgram & ")].[Age]) AS Age," & _
        " [qryRawData (" & opProgram & ")].[ReleaseDev]," & _
        " [qryRawData (" & opProgram & ")].[Operationshorttext]," & _
        " [qryRawData (" & opProgram & ")].[MfgFinishDate]," & _
        " [qryRawData (" & opProgram & ")].DisplayDashboard" & _
        " FROM [qryRawData (" & opProgram & ")]"
    If IsHospital = False Then
        SQL = SQL & " WHERE ((([qryRawData (" & opProgram & ")].[Operationshorttext])='" & opName & "'))" & _
                " AND (([qryRawData (" & opProgram & ")].DisplayDashboard)=True)" & _
                " ORDER BY Cdbl([qryRawData (" & opProgram & ")].[Age]) DESC;"
    Else
        SQL = SQL & " WHERE (("
        For Each element In OpNames
            SQL = SQL & "([qryRawData (" & opProgram & ")].[Operationshorttext]) Not Like '" & element & "' And"
        Next
        SQL = SQL & "([qryRawData (" & opProgram & ")].[Operationshorttext]) Not Like 'Finished Goods'))" & _
                    " AND (([qryRawData (" & opProgram & ")].DisplayHospital)=True)"
        SQL = SQL & "ORDER BY CDbl([qryRawData (" & opProgram & ")].[Age]) DESC;"
    End If
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function generates the qryOp SUM queries that are used to sum the qryOp queries
'This funciton is only called during creation or re-creation of the queries during
'development
Public Function createqryOpSum(ByVal opNum As Integer, ByVal opName As String, ByVal opProgram As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName, qrySumName As String
    Dim element As Variant
    
    'Define titles
    If opName = "Hospital" Then
        qryName = "qryOP_" & opName & " (" & opProgram & ")"
        qrySumName = "qryOPSum_" & opName & " (" & opProgram & ")"
    Else
        qryName = "qryOP_" & opNum & "_" & opName & " (" & opProgram & ")"
        qrySumName = "qryOPSum_" & opNum & "_" & opName & " (" & opProgram & ")"
    End If
    
    'Define SQL string
    SQL = "SELECT Count([" & qryName & "].[SN]) AS QTY " & _
        "FROM [" & qryName & "];"

    'Create query
    Set qdf = CurrentDb.CreateQueryDef(qrySumName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qrySumName
End Function


'This function creates the qryDemand (xxx) query using the program as the input
Public Function createqryFinishedGoods(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryFinishedGoods (" & Program & ")"
    
    SQL = "SELECT tblFinishedGoods.ExternalLongMaterialNumber AS PN," & _
          " Sum(Nz([tblFinishedGoods]![Unrestricted],0)) AS [On Hand]" & _
          " FROM tblFinishedGoods INNER JOIN tblDemandInput ON tblFinishedGoods.ExternalLongMaterialNumber = tblDemandInput.PN" & _
          " WHERE (((tblFinishedGoods.SLoc)<>'RNN1') AND ((tblDemandInput.Program)='" & Program & "'))" & _
          " GROUP BY tblFinishedGoods.ExternalLongMaterialNumber;"


    
    'Create query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
    
End Function

'This function creates the qry31DayDetail (xxx) query using the program as the input
Public Function createqry31DayDetail(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qry31DayMonthDetailed (" & Program & ")"
    

    
    SQL = "SELECT CumulativeDayMonth.[Short Date]," & _
        " CumulativeDayMonth.[Actual Shipments]," & _
        " CumulativeDayMonth.[Required Ship Per Day]," & _
        " CumulativeDayMonth.[Required Cumulative]," & _
        " CumulativeDayMonth.PN," & _
        " IIf([Short Date]=Date(),Nz([qryDemand (" & Program & ")]![On Hand]),0) AS [Finished Goods]," & _
        " Day([Short Date]) AS DayofMonth" & _
        " FROM CumulativeDayMonth LEFT JOIN [qryDemand (" & Program & ")] ON CumulativeDayMonth.PN = [qryDemand (" & Program & ")].PN" & _
        " WHERE (((CumulativeDayMonth.Program)='" & Program & "'));"

    'Create query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
    
End Function

'This function creates the qry31DayMonth (xxx) query using the program as the input
Public Function createqry31DayMonth(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qry31DayMonth (" & Program & ")"
    
    
    
    SQL = "SELECT tblCumulativePlan.shortdate," & _
          " tblCumulativePlan.ReqDlvTot," & _
          " tblCumulativePlan.Program," & _
          " [qryDlvOrdersPNQuantitiesSum (" & Program & ")].DLVTot" & _
          " FROM tblCumulativePlan INNER JOIN [qryDlvOrdersPNQuantitiesSum (" & Program & ")] ON tblCumulativePlan.DayOfMonth = [qryDlvOrdersPNQuantitiesSum (" & Program & ")].Dayy" & _
          " WHERE tblCumulativePlan.Program='" & Program & "';"



    'Create query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
    
End Function

'This function creates the qryPlanExecution query using the program as the input
Public Function createqryPlanExecution(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryPlanExecution (" & Program & ")"
    
    
    
    SQL = "SELECT [qry31DayMonth (" & Program & ")].Program," & _
         " [qry31DayMonth (" & Program & ")].ReqDlvTot," & _
         " [qry31DayMonth (" & Program & ")].DLVTot," & _
         " [qry31DayMonth (" & Program & ")].shortdate" & _
         " FROM [qry31DayMonth (" & Program & ")]" & _
         " WHERE ((([qry31DayMonth (" & Program & ")].shortdate)=Date()));"




    'Create query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
    
End Function


'This function creates the qry31RawData (xxx) query using the program as the input
Public Function createqryRawData(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryRawData (" & Program & ")"
    
    
    SQL = "SELECT [RAW DATA].[Order]," & _
        " [RAW DATA].SN," & _
        " [RAW DATA].ActRelease," & _
        " [RAW DATA].Material ," & _
        " [RAW DATA].Program," & _
        " [RAW DATA].Bscstart," & _
        " [RAW DATA].Basicfin," & _
        " [RAW DATA].Operationshorttext," & _
        " [RAW DATA].TimeRemaining," & _
        " [RAW DATA].OrderType," & _
        " [RAW DATA].ActStartDate," & _
        " [RAW DATA].PlanRelease," & _
        " [RAW DATA].KitDate," & _
        " [RAW DATA].KittingTAT," & _
        " [RAW DATA].Age," & _
        " [RAW DATA].ReleaseDev," & _
        " [RAW DATA].MfgFinishDate," & _
        " [RAW DATA].MRPctrlr," & _
        " [RAW DATA].PrSuperv," & _
        " [RAW DATA].Plant," & _
        " [RAW DATA].Targetqty," & _
        " [RAW DATA].Unit," & _
        " [RAW DATA].ActStartTime,"
    SQL = SQL & " [RAW DATA].ActfinishTime,"
    SQL = SQL & " [RAW DATA].[Minutes Left Till Due]," & _
        " [RAW DATA].[Flag Set]," & _
        " [RAW DATA].DisplayDashboard," & _
        " [RAW DATA].DisplayHospital" & _
        " FROM [RAW DATA]" & _
        " WHERE ([RAW DATA].Program='" & Program & "')"

    

    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryDlvOrdersPNQuantities (xxx) query using the program as the input
Public Function createqryDlvOrdersPNQuantities(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryDlvOrdersPNQuantities (" & Program & ")"
    
    
    SQL = "SELECT [qryDeliveredOrders (" & Program & ")].Material," & _
          " [qryDeliveredOrders (" & Program & ")].ActfinishDate_d," & _
          " Day([ActfinishDate_d]) AS DayOfMonth," & _
          " Sum([qryDeliveredOrders (" & Program & ")].Targetqty) AS DLVQTY" & _
          " FROM [qryDeliveredOrders (" & Program & ")]" & _
          " WHERE (((Month(CDate([qryDeliveredOrders (" & Program & ")].[ActfinishDate_d]))) = Month(Date())))" & _
          " GROUP BY [qryDeliveredOrders (" & Program & ")].Material, [qryDeliveredOrders (" & Program & ")].ActfinishDate_d;"



    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryDlvOrdersPNQuantitiesSum (xxx) query using the program as the input
Public Function createqryDlvOrdersPNQuantitiesSum(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryDlvOrdersPNQuantitiesSum (" & Program & ")"

    
    SQL = "SELECT [31 Days].[Day of Month] AS Dayy," & _
        " Sum(Nz([qryDlvOrdersPNQuantities (" & Program & ")]![DLVQTY],0)) AS DLVQTY," & _
        " Nz(DSum('[DLVQTY]','qryDlvOrdersPNQuantities (" & Program & ")','DayofMonth<=' & [Dayy]),0) AS RunningTot," & _
        " IIf([Dayy]>Day(Date()),0,[RunningTot]) AS DLVTot" & _
        " FROM [31 Days] LEFT JOIN [qryDlvOrdersPNQuantities (" & Program & ")] ON [31 Days].[Day of Month] = [qryDlvOrdersPNQuantities (" & Program & ")].DayofMonth" & _
        " GROUP BY [31 Days].[Day of Month];"
        

    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryPNandDescription (xxx) query using the program as the input
Public Function createqryPNandDescription(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryPNandDescription (" & Program & ")"
    
    
    SQL = "SELECT tblDemandInput.PN," & _
          " tblDemandInput.Program," & _
          " tblDemandInput.Description" & _
          " FROM tblDemandInput" & _
          " WHERE (((tblDemandInput.Program)='" & Program & "'));"

    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryDeliveredOrders (xxx) query using the program as the input
Public Function createqryDeliveredOrders(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryDeliveredOrders (" & Program & ")"
    
    
    SQL = "SELECT DeliveredOrders.Program," & _
        " DeliveredOrders.Targetqty," & _
        " DeliveredOrders.Order," & _
        " DeliveredOrders.Material," & _
        " DeliveredOrders.ActFinishDate_d" & _
        " FROM DeliveredOrders" & _
        " WHERE (((DeliveredOrders.Program)='" & Program & "'));"

    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryDeliveredOrders (xxx) query using the program as the input
Public Function createqryDailyCompletes(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryDailyCompletes (" & Program & ")"
    
    
    SQL = "SELECT Sum(Nz([Targetqty],0)) AS DeliveryQTY," & _
        " [Output Linearity Setup].Date2" & _
        " FROM [qryDeliveredOrders (" & Program & ")] RIGHT JOIN [Output Linearity Setup] ON [qryDeliveredOrders (" & Program & ")].[ActfinishDate_d] = [Output Linearity Setup].Date2" & _
        " GROUP BY [Output Linearity Setup].Date2;"


    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryOTDSetup (xxx) query using the program as the input
Public Function createqryOTDSetup(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOTDSetup (" & Program & ")"
    
    
    SQL = "SELECT tblOTD.Material," & _
          " CDate([ShpCmplDte]) AS ShpCmplDte_d," & _
          " tblOTD.CustReqDate," & _
          " tblOTD.ShippedQuantity," & _
          " (CDate([ShpCmplDte])-CDate([CustReqDate])) AS [Late Identifier]," & _
          " IIf([Late Identifier]<=0,'On Time','Late') AS [On Time Flag]," & _
          " tblDemandInput.Program" & _
          " FROM tblOTD INNER JOIN tblDemandInput ON tblOTD.Material = tblDemandInput.PN" & _
          " WHERE (((CDate([ShpCmplDte])) Between DateSerial(Year(Date()),Month(Date()),1) And DateSerial(Year(Date()),Month(Date()),31)) AND ((tblDemandInput.Program)='" & Program & "'));"
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryOTDOnTime (xxx) query using the program as the input
Public Function createqryOTDOnTime(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOTDOnTime (" & Program & ")"
    
    
    SQL = "SELECT [qryOTDSetup (" & Program & ")].[On Time Flag]," & _
        " [qryOTDSetup (" & Program & ")].[ShippedQuantity]" & _
        " FROM [qryOTDSetup (" & Program & ")]" & _
        " WHERE ((([qryOTDSetup (" & Program & ")].[On Time Flag])='On Time'));"
    
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function


'This function creates the qryOTDOnTimeSum (xxx) query using the program as the input
Public Function createqryOTDOnTimeSum(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOTDOnTimeSum (" & Program & ")"
    
    
    SQL = "SELECT CDbl(Nz(Sum([qryOTDOnTime (" & Program & ")].[ShippedQuantity]),0)) AS [SumOfShippedQuantity]" & _
        " FROM [qryOTDOnTime (" & Program & ")];"
    
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryOTDLate (xxx) query using the program as the input
Public Function createqryOTDLate(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOTDLate (" & Program & ")"
    
    
    SQL = "SELECT [qryOTDSetup (" & Program & ")].[On Time Flag]," & _
        " [qryOTDSetup (" & Program & ")].[ShippedQuantity]" & _
        " FROM [qryOTDSetup (" & Program & ")]" & _
        " WHERE ((([qryOTDSetup (" & Program & ")].[On Time Flag])='Late'));"
    
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function


'This function creates the qryOTDLateSum (xxx) query using the program as the input
Public Function createqryOTDLateSum(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOTDLateSum (" & Program & ")"
    
    
    SQL = "SELECT CDbl(Nz(Sum([qryOTDLate (" & Program & ")].[ShippedQuantity]),0)) AS [SumOfShippedQuantity]" & _
        " FROM [qryOTDLate (" & Program & ")];"
    
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryOTD (xxx) query using the program as the input
Public Function createqryOTD(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOTD (" & Program & ")"
    
    
    SQL = "SELECT IIf(([qryOTDOnTimeSum (" & Program & ")]![SumOfShippedQuantity]=0 And [qryOTDLateSum (" & Program & ")]![SumOfShippedQuantity]=0),'N/A',Round([qryOTDOnTimeSum (" & Program & ")]![SumOfShippedQuantity]/([qryOTDLateSum (" & Program & ")]![SumOfShippedQuantity]+[qryOTDOnTimeSum (" & Program & ")]![SumOfShippedQuantity])*100,0)) AS [OTD %]" & _
        " FROM [qryOTDLateSum (" & Program & ")], [qryOTDOnTimeSum (" & Program & ")];"

    
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryOutputLinearity (xxx) query using the program as the input
Public Function createqryOutputLinearity(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryOutputLinearity (" & Program & ")"
    
    
    SQL = "SELECT StDev([qryDailyCompletes (" & Program & ")]![DeliveryQTY]) AS [Standard Deviation]," & _
        " Avg([qryDailyCompletes (" & Program & ")]![DeliveryQTY]) AS [Average Units/Day]," & _
        " IIf([Average Units/Day]<1,'N/A',Round(IIf([Standard Deviation]=0,100,(1-(([Standard Deviation]/3)/[Average Units/Day]))*100),0)) AS [Output Linearity]" & _
        " FROM [qryDailyCompletes (" & Program & ")];"
    
    
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryShipments (xxx) query using the program as the input
Public Function createqryShipments(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryShipments (" & Program & ")"
    
    
    SQL = "SELECT tblShipments.Material," & _
          " tblShipments.MaterialDescription," & _
          " tblShipments.Serialnumber," & _
          " tblShipments.Deliveryquantity," & _
          " CDate([AcGIdate]) AS AcGIdate_d," & _
          " tblShipments.TRKID," & _
          " tblShipments.ForwardingagenttrackingID," & _
          " tblShipments.SLoc," & _
          " tblShipments.DlvTy," & _
          " tblDemandInput.Program" & _
          " FROM tblShipments INNER JOIN tblDemandInput ON tblShipments.Material = tblDemandInput.PN" & _
          " WHERE (((CDate([AcGIdate])) Between DateSerial(Year(Date()),Month(Date()),1) And DateSerial(Year(Date()),Month(Date()),31))" & _
          " AND (Not (tblShipments.DlvTy)='LR') " & _
          " AND ((tblDemandInput.Program)='" & Program & "'));"
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryShipmentsSum (xxx) query using the program as the input
Public Function createqryShipmentsSum(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryShipmentsSum (" & Program & ")"
    
    
    SQL = "SELECT Count(Nz([qryShipments (" & Program & ")]![Serialnumber],0)) AS [Total Shipped]," & _
        " [qryShipments (" & Program & ")].Material" & _
        " FROM [qryShipments (" & Program & ")]" & _
        " GROUP BY [qryShipments (" & Program & ")].Material;"

    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryDemandInput (xxx) query using the program as the input
Public Function createqryDemandInput(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryDemandInput (" & Program & ")"
    
    
    SQL = "SELECT [Demand Input (" & Program & ")].*" & _
        " FROM [Demand Input (" & Program & ")];"


    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryTotalCompleteDetailed (xxx) query using the program as the input
Public Function createqryTotalCompleteDetailed(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryTotalCompleteDetailed (" & Program & ")"
    
    
    SQL = "SELECT tblDemandInput.PN," & _
          " First(Nz([qryFinishedGoods (" & Program & ")].[On Hand],0)) AS [Finished Goods]," & _
          " Sum(Nz([qryDlvOrdersPNQuantities (" & Program & ")].[DLVQTY],0)) AS [Total Completed]," & _
          " First(tblDemandInput.Program) AS Program," & _
          " First(Nz([qryShipmentsSum (" & Program & ")].[Total Shipped],0)) AS [Total Shipped]" & _
          " FROM (([qryFinishedGoods (" & Program & ")] RIGHT JOIN tblDemandInput ON [qryFinishedGoods (" & Program & ")].PN = tblDemandInput.PN)" & _
          " LEFT JOIN [qryDlvOrdersPNQuantities (" & Program & ")] ON tblDemandInput.PN = [qryDlvOrdersPNQuantities (" & Program & ")].Material)" & _
          " LEFT JOIN [qryShipmentsSum (" & Program & ")] ON tblDemandInput.PN = [qryShipmentsSum (" & Program & ")].Material" & _
          " WHERE (((tblDemandInput.Program)='" & Program & "'))" & _
          " GROUP BY tblDemandInput.PN;"



    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryTotalComplete (xxx) query using the program as the input
Public Function createqryTotalComplete(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryTotalComplete (" & Program & ")"
    
    
    SQL = "SELECT [qryTotalCompleteDetailed (" & Program & ")].Program," & _
        " Sum(Nz([qryTotalCompleteDetailed (" & Program & ")].[Total Completed],0)) AS [Total Completed]," & _
        " Sum(NZ([qryTotalCompleteDetailed (" & Program & ")].[Total Shipped],0)) AS [Total Shipped]," & _
        " Sum(NZ([qryTotalCompleteDetailed (" & Program & ")].[Finished Goods],0)) AS [Finished Goods]" & _
        " FROM [qryTotalCompleteDetailed (" & Program & ")]" & _
        " GROUP BY [qryTotalCompleteDetailed (" & Program & ")].Program;"



    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryTotalWIP (xxx) query using the program as the input
Public Function createqryTotalWIP(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryTotalWIP (" & Program & ")"
    
    
    SQL = "SELECT [qryRawData (" & Program & ")].Program," & _
          " Count([qryRawData (" & Program & ")].SN) AS CountOfSN" & _
          " FROM [qryRawData (" & Program & ")]" & _
          " GROUP BY [qryRawData (" & Program & ")].Program;"



    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryTotalWIPDetailed (xxx) query using the program as the input
Public Function createqryTotalWIPDetailed(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryTotalWIPDetailed (" & Program & ")"
    
    
    SQL = "SELECT [qryRawData (" & Program & ")].Material," & _
          " Count([qryRawData (" & Program & ")].SN) AS CountOfSN" & _
          " FROM [qryRawData (" & Program & ")]" & _
          " GROUP BY [qryRawData (" & Program & ")].Material;"


    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryTotalWIPSWIP (xxx) query using the program as the input
Public Function createqryTotalWIPSWIP(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    Dim element As Variant
    
    'Define title
    qryName = "qryTotalWIPSWIP (" & Program & ")"
    
    
    SQL = "SELECT [qryTotalWIP (" & Program & ")].[CountOfSN] AS [Total WIP], " & _
          "  Sum(CInt(Nz([tblOpNames_All].[SWIP],0))) AS [Total SWIP]" & _
          " FROM [qryTotalWIP (" & Program & ")] INNER JOIN tblOpNames_All ON" & _
          " [qryTotalWIP (" & Program & ")].[Program] = tblOpNames_All.[Program]" & _
          " GROUP BY [qryTotalWIP (" & Program & ")].[CountOfSN];"



    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qrySummaryStats (xxx) query using the program as the input
Public Function createqrySummaryStats(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qrySummaryStats (" & Program & ")"
    
    
    SQL = "SELECT [qryOTD (" & Program & ")].[OTD %]," & _
        " [qryOutputLinearity (" & Program & ")].[Output Linearity]," & _
        " [qryTotalWIPSWIP (" & Program & ")].[Total SWIP]," & _
        " [qryTotalWIPSWIP (" & Program & ")].[Total WIP]," & _
        " Date() AS [Date]," & _
        " [qryTotalComplete (" & Program & ")].[Total Shipped]," & _
        " [qryTotalComplete (" & Program & ")].[Finished Goods]," & _
        " [qryTotalComplete (" & Program & ")].[Total Completed]," & _
        " [Standard Factory Metrics].Program," & _
        " Round([AvgOfAge],1) AS ShopOrderAge," & _
        " Round([AvgOfDaysToMFR],1) AS DaystoManufacture,"
    SQL = SQL & " Round((CDbl([DeliveredOrdersSummary.AvgOfReleaseDev])+CDbl([OpenOrdersReleaseSummary.AvgOfReleaseDev]))/2,1) AS ReleaseDevFromPlan," & _
        " Round((CDbl([DeliveredOrdersSummary.AvgOfDaysToStart])+CDbl([OpenOrdersStartSummary.AvgOfDaysToStart]))/2,1) AS ActStartAfterRelease," & _
        " [Standard Factory Metrics].OTRTotal," & _
        " [qryPlanExecution (" & Program & ")].[ReqDlvTot] AS [Required Cumulative], " & _
        " IIf((IsNull([DeliveredOrdersKitTATSummary].[AvgOfKittingTAT]) And IsNull([OpenOrdersKitTATSummary].[AvgOfKittingTAT])),'N/A',IIf((IsNull([DeliveredOrdersKitTATSummary].[AvgOfKittingTAT]) Or IsNull([OpenOrdersKitTATSummary].[AvgOfKittingTAT])),Round((CDbl(Nz([DeliveredOrdersKitTATSummary.AvgOfKittingTAT],0))+CDbl(Nz([OpenOrdersKitTATSummary.AvgOfKittingTAT],0))),1),Round((CDbl([DeliveredOrdersKitTATSummary.AvgOfKittingTAT])+CDbl([OpenOrdersKitTATSummary.AvgOfKittingTAT]))/2,1))) AS KittingTAT" & _
        " FROM [qryOTD (" & Program & ")]," & _
        " [qryOutputLinearity (" & Program & ")], [qryTotalWIPSWIP (" & Program & ")]," & _
        " [qryTotalComplete (" & Program & ")], [Standard Factory Metrics]," & _
        " [qryPlanExecution (" & Program & ")]" & _
        " WHERE ((([Standard Factory Metrics].Program)='" & Program & "'));"

    '    " IIf((IsNull([DeliveredOrdersKitTATSummary].[AvgOfKittingTAT]) And IsNull([OpenOrdersKitTATSummary].[AvgOfKittingTAT])),'N/A',IIf((IsNull([DeliveredOrdersKitTATSummary].[AvgOfKittingTAT]) Or IsNull([OpenOrdersKitTATSummary].[AvgOfKittingTAT])),Round((CDbl(Nz([DeliveredOrdersKitTATSummary.AvgOfKittingTAT],0))+CDbl(Nz([OpenOrdersKitTATSummary.AvgOfKittingTAT],0))),1),Round((CDbl([DeliveredOrdersKitTATSummary.AvgOfKittingTAT])+CDbl([OpenOrdersKitTATSummary.AvgOfKittingTAT]))/2,1))) AS KittingTAT"
    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qrySummaryStats (xxx) query using the program as the input
Public Function createqryStoreSummaryStats(ByVal Program As String)
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryStoreSummaryStats (" & Program & ")"
    
    
    SQL = "INSERT INTO SummaryStatsLog ( [OTD %]," & _
        " [Output Linearity]," & _
        " [Total SWIP]," & _
        " [Total WIP]," & _
        " [Date]," & _
        " [Total Shipped]," & _
        " [Total Finished Goods]," & _
        " [Total Complete]," & _
        " [Required Cumulative]," & _
        " Program," & _
        " ShopOrderAge," & _
        " ReleaseDevFromPlan," & _
        " ActStartAfterRelease," & _
        " KittingTAT," & _
        " DaystoManufacture )" & _
        " SELECT [qrySummaryStats (" & Program & ")].[OTD %]," & _
        " [qrySummaryStats (" & Program & ")].[Output Linearity]," & _
        " [qrySummaryStats (" & Program & ")].[Total SWIP]," & _
        " [qrySummaryStats (" & Program & ")].[Total WIP]," & _
        " [qrySummaryStats (" & Program & ")].Date," & _
        " [qrySummaryStats (" & Program & ")].[Total Shipped]," & _
        " [qrySummaryStats (" & Program & ")].[Finished Goods]," & _
        " [qrySummaryStats (" & Program & ")].[Total Completed]," & _
        " [qrySummaryStats (" & Program & ")].[Required Cumulative]," & _
        " [qrySummaryStats (" & Program & ")].Program,"
    SQL = SQL & " [qrySummaryStats (" & Program & ")].ShopOrderAge," & _
        " [qrySummaryStats (" & Program & ")].ReleaseDevFromPlan," & _
        " [qrySummaryStats (" & Program & ")].ActStartAfterRelease," & _
        " [qrySummaryStats (" & Program & ")].KittingTAT," & _
        " [qrySummaryStats (" & Program & ")].DaystoManufacture" & _
        " FROM [qrySummaryStats (" & Program & ")];"

    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function

'This function creates the qryRawData_RwrkOrders query using the program as the input
Public Function createqryRwrkOrders()
    Dim SQL As Variant
    Dim qdf As QueryDef
    Dim qryName As String
    
    'Define title
    qryName = "qryRawData_RwrkOrders"
    
    
    SQL = "SELECT [RAW DATA].Order," & _
        " [RAW DATA].[Serial number]," & _
        " [RAW DATA].[Actual release date]," & _
        " [RAW DATA].Material," & _
        " [RAW DATA].[Team Name]," & _
        " [RAW DATA].[Basic start date]," & _
        " [RAW DATA].[User status date]," & _
        " [RAW DATA].[Last move date]," & _
        " [RAW DATA].[Basic finish date]," & _
        " [RAW DATA].OpAc," & _
        " [RAW DATA].Operation," & _
        " [RAW DATA].Stat," & _
        " [RAW DATA].[Operations remaining in order]," & _
        " [RAW DATA].[Time Remaining]," & _
        " [RAW DATA].[Order Type]" & _
        " FROM [RAW DATA]" & _
        " WHERE ((([RAW DATA].[Order Type]) Like 'zrw*'));"


    'Create Query
    Set qdf = CurrentDb.CreateQueryDef(qryName, SQL)
    DoCmd.OpenQuery qdf.Name
    DoCmd.Close acQuery, qryName
End Function



'---------------------------------------------
' Call Functions
'---------------------------------------------

'This function loops over the sequences in the given table and calls the
'createqryOp function only used during delopemnt and dashboard creation
Public Function createqryOpCall(ByVal OpNamesTable As String)
    Dim DB As Object
    Dim rs As DAO.Recordset
    Dim opName, opProgram As String
    Dim opNum As Integer
    Dim OpNames() As String
    Dim element As Variant
    Dim i As Integer
    
    Set DB = CurrentDb
    Set rs = DB.OpenRecordset(OpNamesTable)
    
    i = 0
    'Check to see if the recordset actually contains rows
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True

    
            'Save values to var
            opName = rs("OperationName")
            opProgram = rs("Program")
            opNum = rs("OperationNum")
            
            'create array of all op names this will be used
            'to create the hospital call
            ReDim Preserve OpNames(i)
            OpNames(i) = opName
            
            'Call function to create select queries
            Call createqryOp(opNum, opName, opProgram, False)
            
            'Call function to create select sum queries
            Call createqryOpSum(opNum, opName, opProgram)
            
            
            i = i + 1
            
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
    
    
    'Call function for hospital query
    Call createqryOp(opNum, opName, opProgram, True, OpNames)
    'Call function to create select queries for hospital
    Call createqryOpSum("1", "Hospital", opProgram)
    
    
End Function

'This function calls all the program level create query functions
Public Function createqryProgramCall(ByVal Program As String)
    Call createqryRawData(Program)
    Call createqryPNandDescription(Program)
    Call createqryDeliveredOrders(Program)
    Call createqryDlvOrdersPNQuantities(Program)
    Call createqryDlvOrdersPNQuantitiesSum(Program)
    Call createqry31DayMonth(Program)
    Call createqryDailyCompletes(Program)
    Call createqryOTDSetup(Program)
    Call createqryOTDOnTime(Program)
    Call createqryOTDOnTimeSum(Program)
    Call createqryOTDLate(Program)
    Call createqryOTDLateSum(Program)
    Call createqryOTD(Program)
    Call createqryOutputLinearity(Program)
    Call createqryShipments(Program)
    Call createqryShipmentsSum(Program)
    Call createqryFinishedGoods(Program)
    Call createqryTotalCompleteDetailed(Program)
    Call createqryTotalComplete(Program)
    Call createqryTotalWIP(Program)
    Call createqryTotalWIPDetailed(Program)
    Call createqryOpCall("tblOpNames")
    Call createqryPlanExecution(Program)
    Call createqryTotalWIPSWIP(Program)
    Call createqrySummaryStats(Program)
    Call createqryStoreSummaryStats(Program)
    
    MsgBox "Operation Queries Created"
End Function

'Logs the error description, time, and function or sub that the error occured in
Public Sub LogErrorDesc(ErrTxt As String, ErrLoc As String)
    Dim DB As Object
    Dim SQL As String
    
    Set DB = CurrentDb
    
    SQL = "INSERT INTO ErrorLog ( [DateTime], ErrorDesc, LocOfError )" & _
          " SELECT '" & Now() & "', '" & ErrTxt & "', '" & ErrLoc & "';"
    
    DB.Execute (SQL)
    
    Set DB = Nothing
    
End Sub


'This function returns the program name of a give form
'by taking what is between the parenthesis in the form
'title
Public Function FindCurrentProgram(frmObj As Object) As String
    Dim CurrentFormName, Program, temp As String
    Dim paren1Loc, paren2Loc As Integer

    'Determine current form program name to pass to
    CurrentFormName = frmObj.Name
    paren1Loc = InStr(CurrentFormName, "(")
    paren2Loc = InStr(CurrentFormName, ")")
    Program = Left(CurrentFormName, paren2Loc - 1)
    Program = Right(Program, paren2Loc - paren1Loc - 1)
    
    FindCurrentProgram = Program
End Function
