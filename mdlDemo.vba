Option Compare Database
Option Explicit

' Module Constant values
Private Const INTNUMCUSTOMERS As Integer = 10
Private Const INTNUMITEMS As Integer = 30
Private Const INTCUSTITEMS As Integer = 5



Public Function AutoPopulate()
    
    clearTables

    '''''''''''''''''''''''''''''''''
    ' Here we get everything ready
    '''''''''''''''''''''''''''''''''
    

    ' Module Variables
    Dim rstItems As DAO.Recordset, _
    intLoanIDs() As Integer, _
    intAmounts() As Integer, _
    intAmountsLeft() As Integer, _
    intPickupDateID As Integer

    ReDim intAmounts(INTNUMITEMS - 1) As Integer
    ReDim intAmountsLeft(INTNUMITEMS - 1) As Integer
    
    
    ' Get the item IDs
    Set rstItems = CurrentDb.OpenRecordset("Select TOP " & INTNUMITEMS & " ID, Amount From Items " _
            & "ORDER BY Rnd(-(100000*ID)*Time())")

    ' Store the item amounts
    Do Until rstItems.EOF
        intAmounts(rstItems.AbsolutePosition()) = rstItems("Amount")
        rstItems.MoveNext
    Loop
    rstItems.MoveFirst

    intAmountsLeft = intAmounts

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Here we start with the creation of the data - for next PickupDate, last PickupDate and past PickupDate
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' First for the next pickup date - reservations
    intLoanIDs = initLoans(gintNextPickupDateID(), "reserved") ' initialize loans
    intAmountsLeft = initLoanItems(intLoanIDs, intAmountsLeft, rstItems) ' initialize loan items

    ' For last pickupDate - active loans
    intLoanIDs = initLoans(gintLastPickupDateID(), "active") 'initialize loans
    intAmountsLeft = initLoanItems(intLoanIDs, intAmountsLeft, rstItems) 'initialize loanItems
    setDemoDepositGiven (gintLastPickupDateID()) 'set deposit give for each loan (dependant on the loanItems, so waited for this)
    CurrentDb.Execute "Update Loans Set MaintenanceFeePaid = gcurLoanCleaningFee(ID) Where PickupDateID = " _
                        & gintLastPickupDateID(), dbFailOnError 'see note above
    

    ' Now a past pickupdate (one that doesn't conflict with most recent pickupDate) - mostly returned Loans
    intPickupDateID = DLookup("ID", "tblPickupDates", "ReturnDate <" & "#" & Now - gintCleanersDays() - 10 _
                                & "# And ReturnDate > " & "#" & Now - gintCleanersDays() - 20 & "#") 'Make sure cleaners doesn't conflict
    intLoanIDs = initLoans(intPickupDateID, "past")
    initLoanItems intLoanIDs, intAmounts, rstItems
    setDemoDepositGiven (intPickupDateID)
    CurrentDb.Execute "Update Loans Set MaintenanceFeePaid = gcurLoanCleaningFee(ID) Where PickupDateID = " & intPickupDateID, dbFailOnError
    
    ' For past pickupDate, need to initialize returns
    initLoanItemReturns (intPickupDateID) 'initialize returns for all the LoanItems
    initLoanItemsLostAndUnreturned rstItems, intAmountsLeft, intPickupDateID 'initialize Lost and Unreturned instances
    
    ' Now set status
    initLoanStatus (intPickupDateID)
    
    ' Now make the charges
    initLoanCharges (intPickupDateID)
    
    ' Now make the payments (and update status)
    initLoanPayments (intPickupDateID)
    
AutoPopulate_Exit:
    'rstCustomerIDs.Close
    rstItems.Close
    'Set rstCustomerIDs = Nothing
    Set rstItems = Nothing
    Exit Function
    
AutoPopulate_Err:
    Resume AutoPopulate_Exit
    
End Function


'''''''''''''''''''''''''''''''''''''''
' Initialization Helper functions
'''''''''''''''''''''''''''''''''''''''

' Delete old data, and revert inventory
Public Function clearTables()
    ' Clear Inventory stuff
    CurrentDb.Execute ("Delete * From InventoryChanges Where Type <> 1")
    CurrentDb.Execute ("Delete * From InventoryLog")
    ' Put back in basic inventory transactions, with a date that's plausible
    CurrentDb.Execute ("Insert Into InventoryLog " _
                     & "(ItemID, ChangeDate, ChangeAmount, UpdatedAmount, Type, ChangeRecordID) " _
                     & "Select ItemID, " _
                            & "ChangeDate, " _
                            & "ChangeAmount, " _
                            & "Amount, " _
                            & "Type, " _
                            & "InventoryChanges.ID " _
                     & "From Items " _
                     & "INNER JOIN InventoryChanges ON Items.ID = InventoryChanges.ItemID")
    CurrentDb.Execute ("Delete * from Loans")
    CurrentDb.Execute ("Delete * from LoanItems")
    
End Function


' Initialize Loans
Private Function initLoans(intPickupDateID As Integer, strLoanTypes As String)
    ''' The strLoanTypes determines what type of loans we'll create. Three options are "reserved", "active" or "past"
    ''' "reserved" means reserved, but not yet picked up
    ''' "active" means picked up already, but not yet due back
    ''' "past" means due back already
    
    ' Get the Customer IDs
    Dim rstCustomerIDs As DAO.Recordset
    Set rstCustomerIDs = CurrentDb.OpenRecordset("Select TOP " & INTNUMCUSTOMERS & " ID From Customers " _
            & "ORDER BY Rnd(-(100000*ID)*Time())")
    
    ' Initialize variable
    Dim intRandom As Integer, _
    strSQL As String, _
    intLoanIDs(INTNUMCUSTOMERS - 1) As Integer

    ' Loop over the customers
    Do Until rstCustomerIDs.EOF
        
        Select Case strLoanTypes
            Case "reserved"
        
                ' Get a random status of the loan
                intRandom = gintRandom(1, 3)
                intRandom = IIf(intRandom = 3, 7, intRandom) ' if it's three, switch to seven.
                
                strSQL = "INSERT INTO Loans (CustomerID, Status, ReminderSent, PickupDateID) " _
                        & "VALUES (" & rstCustomerIDs("ID") & "," & intRandom & "," & gbolRandom() & "," & intPickupDateID & ")"
        
            Case "active"
                ' Set the point-in-time fee data for the pickupdate
                setPickupDateFees (intPickupDateID)
                
                ' Create Random Value for deposit method and maintenance fee method
                intRandom = gintRandom(1, 3)
                
                strSQL = "INSERT INTO Loans (CustomerID, Status, DepositMethod, MaintenanceFeeMethod, AllDryCleaned, ReminderSent, ReturnReminderSent, PickupDateID, DepositStatus) " _
                        & "VALUES (" & rstCustomerIDs("ID") & ", 3," & intRandom & "," & intRandom & "," & gbolRandom() & "," & gbolRandom() & "," _
                        & gbolRandom() & "," & intPickupDateID & "," & 2 & ")"
            
            Case "Past"
                ' Same random data as 'active' loans, except status is fully returned
                
                ' Set the point-in-time fee data for the pickupdate
                setPickupDateFees (intPickupDateID)
                
                ' Create Random Value for deposit method and maintenance fee method
                intRandom = gintRandom(1, 3)
                
                strSQL = "INSERT INTO Loans (CustomerID, Status, DepositMethod, MaintenanceFeeMethod, AllDryCleaned, ReminderSent, ReturnReminderSent, PickupDateID, DepositStatus) " _
                        & "VALUES (" & rstCustomerIDs("ID") & ", 3," & intRandom & "," & intRandom & "," & gbolRandom() & "," & gbolRandom() & "," _
                        & gbolRandom() & "," & intPickupDateID & "," & 2 & ")"
                
        End Select
        
        ' Create the loan
        CurrentDb.Execute (strSQL) ' We don't fail on error, because there might be a duplicate that won't be allowed because of unique index
        
        ' Save the LoanID
        intLoanIDs(rstCustomerIDs.AbsolutePosition) = DMax("ID", "Loans")
        rstCustomerIDs.MoveNext
        
    Loop
    
    initLoans = intLoanIDs
    
initLoans_Exit:
    rstCustomerIDs.Close
    Set rstCustomerIDs = Nothing
    Exit Function

initLoans_Err:
    Resume initLoans_Exit

End Function


' Initialize Loan Items
Private Function initLoanItems(intLoanIDs() As Integer, intAmountsLeft() As Integer, rstItems As DAO.Recordset)

    Dim LoanIndex As Integer, intRandom As Integer, i As Integer, _
    intItemsUsed(INTCUSTITEMS - 1) As Integer 'Zero Indexing
    
    For LoanIndex = 0 To UBound(intLoanIDs)
        
        Erase intItemsUsed 'Reset for each Loan/Customer
        
        ' Add the items to that loan
        
        i = 0
        While i < INTCUSTITEMS

            ' Choose a random item to choose from the dataset
            intRandom = gintRandom(0, INTNUMITEMS - 1) ' Use zero based
            While IsInArray(intRandom + 1, intItemsUsed) ' Need the + 1, because array is initialized with zeros in VBA
                intRandom = gintRandom(0, INTNUMITEMS - 1)
            Wend
            rstItems.AbsolutePosition = intRandom
            intItemsUsed(i) = intRandom + 1 ' See comment above for the + 1
            
            ' Choose a random amount between 1 and amount available
            intRandom = gintRandom(0, intAmountsLeft(rstItems.AbsolutePosition))
            If LoanIndex = UBound(intLoanIDs) Then ' Make sure on the last iter we use all remaining items, so a few will be fully used
                intRandom = intAmountsLeft(rstItems.AbsolutePosition)
            End If
            If intRandom <> 0 Then ' We don't want loanItems with zero amounts
                CurrentDb.Execute "INSERT INTO LoanItems (LoanID, ItemID, Amount) " _
                    & "VALUES (" & intLoanIDs(LoanIndex) & "," & rstItems("ID") & "," & intRandom & ")", dbFailOnError
                intAmountsLeft(rstItems.AbsolutePosition) = intAmountsLeft(rstItems.AbsolutePosition) - intRandom
            End If
            
            i = i + 1
        
        Wend
        
        ' Reset Item position
        rstItems.MoveFirst
        
        initLoanItems = intAmountsLeft
    
    Next LoanIndex

End Function

Private Sub initLoanItemReturns(intPickupDateID As Integer)
    Dim datReturnDate As Date, rstLoanItems As DAO.Recordset, rstLoanItemReturns As DAO.Recordset, _
    rstLoanintAmountLeft As Integer, intReturnDates As Integer, intReturnAmount As Integer, intIndex As Integer, intAmountLeft As Integer
    
    ' Get the return Date
    datReturnDate = gdatReturnDate(intPickupDateID)
    
    ' get the LoanItems with the amounts
    Set rstLoanItems = CurrentDb.OpenRecordset("Select LoanItems.ID, LoanItems.Amount " _
                                                & "From LoanItems INNER JOIN Loans " _
                                                & "ON LoanItems.LoanID = Loans.ID " _
                                                & "WHERE Loans.PickupDateID = " & intPickupDateID)
        
    ' Open LoanItemReturns Table to make new records
    Set rstLoanItemReturns = CurrentDb.OpenRecordset("LoanItemReturns")
    
    While Not rstLoanItems.EOF
        
        ' how many of this item need to be returned
        intAmountLeft = rstLoanItems("Amount")
        intReturnDates = gintRandom(1, 3)
        For intIndex = 0 To (intReturnDates - 1) '0 indexed
            ' code to add new
            If intIndex = intReturnDates - 1 Then 'if it's the last day, return everything
                intReturnAmount = intAmountLeft
            Else
                intReturnAmount = gintRandom(0, intAmountLeft)
            End If
            ' code to add new and update how many left to be returned
            If intReturnAmount <> 0 Then
                rstLoanItemReturns.AddNew
                rstLoanItemReturns("LoanItemID") = rstLoanItems("ID")
                rstLoanItemReturns("ReturnDate") = datReturnDate + intIndex
                rstLoanItemReturns("ReturnAmount") = intReturnAmount
                rstLoanItemReturns.Update
                intAmountLeft = intAmountLeft - intReturnAmount
            End If
        Next intIndex
        rstLoanItems.MoveNext
    Wend
    rstLoanItems.MoveFirst
                
End Sub

' Randomly make Lost and Late items for past PickupDate in a way that creates conflicts for future reservations, and only future reservations
Private Sub initLoanItemsLostAndUnreturned(rstItems As DAO.Recordset, intAmountsLeft() As Integer, intPickupDateID As Integer)
    Dim rstLoanItemsPast As DAO.Recordset, rstLoanItemTotalsNext As DAO.Recordset, rstLoanItemReturns As DAO.Recordset, _
        rstInventoryChanges As DAO.Recordset, _
        wrkCurrent As DAO.Workspace
        
    Dim i As Integer, intDone As Integer, intItemID As Integer, intDeductable As Integer, intAmountLeft As Integer, _
        intRandomAmount As Integer, intRandomType As Integer, dblRandom As Double
    
    ''''''''''''''''''''''''
    ' For debugging
    ''''''''''''''''''''''''
    Dim intLostMade As Integer, intLateMade As Integer
    Dim strLostLoans As String, strLateLoans As String, strLateLoanItems, strLateLoanItemReturns
    strLostLoans = "Lost LoanIDs: "
    strLateLoans = "Late LoanIDs: "
    strLateLoanItems = "Late LoanItemIDs: "
    strLateLoanItemReturns = "Late LoanItemReturnIDs: "
    
    Dim strSQL As String
    strSQL = "LoanItems Inner Join (Select ID, PickupDateID From Loans) As Loans On LoanItems.LoanID = Loans.ID Where Loans.PickupDateID = "
    
    Set rstLoanItemsPast = CurrentDb.OpenRecordset("Select * from " & strSQL & intPickupDateID) 'get all LoanItems from Past PickupDate
    strSQL = "Select ItemID, sum(Amount) As TotalAmount from " & strSQL & gintNextPickupDateID() & " Group By ItemID" 'LoanItems Totals from Next PickupDate
    Set rstLoanItemTotalsNext = CurrentDb.OpenRecordset(strSQL) ' get LoanItems Totals for next PickupDate
    Set rstLoanItemReturns = CurrentDb.OpenRecordset("Select * from LoanItemReturns") ' We'll be able to look up the one we need by LoanItemID
    Set rstInventoryChanges = CurrentDb.OpenRecordset("InventoryChanges")
    Set wrkCurrent = DBEngine.Workspaces(0)
    
    intDone = 0
    
    For i = 0 To UBound(intAmountsLeft)
        If intDone = Int(INTNUMITEMS / 3) Then Exit For 'Maximum amount of Lost/unreturned items
        
        ' Get amount available for next PickupDate
        intAmountLeft = intAmountsLeft(i)
        ' get the itemID
        rstItems.AbsolutePosition = i
        intItemID = rstItems("ID")
        ' Get the record for totals for this ItemID for the next pickupDate
        rstLoanItemTotalsNext.FindFirst ("ItemID = " & intItemID)
        ' find a past loan item, (with a return), for this ID
        rstLoanItemsPast.FindFirst ("ItemID = " & intItemID)
        
        ' In order to be a good match, we need
        ' 1. That there are no more available for next PickupDate
        ' 2. That there are some booked for next PickupDate
        ' 3. That there are some booked for the PickupDate in the past
        If intAmountLeft = 0 And Not rstLoanItemTotalsNext.NoMatch And Not rstLoanItemsPast.NoMatch Then

            rstLoanItemReturns.FindFirst ("LoanItemID = " & rstLoanItemsPast("LoanItems.ID")) 'Find a past return for this item
            intDeductable = gMin(rstLoanItemTotalsNext("TotalAmount"), rstLoanItemReturns("ReturnAmount")) 'The amount we can lose/unreturn
            ' Lost/Damaged or Unreturned or both?
            dblRandom = Rnd()
            intRandomType = IIf(dblRandom < 0.45, 1, IIf(dblRandom < 0.9, 2, 3)) '1 is lost/damaged, 2 is unreturned, 3 is both
            If intRandomType = 1 Or intRandomType = 3 Then
                '''''''''''''''''
                ' For debugging
                '''''''''''''''''
                intLostMade = intLostMade + 1
                strLostLoans = strLostLoans & rstLoanItemsPast("LoanID") & " "
                
                ' Make a lost Amount, and deduct it from the return
                intRandomAmount = gintRandom(1, intDeductable)
                wrkCurrent.BeginTrans
                    ' Make the lost record
                    rstLoanItemsPast.Edit
                    rstLoanItemsPast("AmountLostDamaged") = rstLoanItemsPast("AmountLostDamaged") + intRandomAmount
                    rstLoanItemsPast.Update
                    
                    
                    
                    ' Whatever was lost, wasn't returned
                    If rstLoanItemReturns("ReturnAmount") = intRandomAmount Then
                        rstLoanItemReturns.Delete
                    Else
                        rstLoanItemReturns.Edit
                        rstLoanItemReturns("ReturnAmount") = rstLoanItemReturns("ReturnAmount") - intRandomAmount
                        rstLoanItemReturns.Update
                    End If
                                        
                    ' Create a inventory change record
                    rstInventoryChanges.AddNew
                    rstInventoryChanges("ItemID") = intItemID
                    rstInventoryChanges("ChangeDate") = gdatReturnDate(intPickupDateID)
                    rstInventoryChanges("ChangeAmount") = intRandomAmount * -1
                    rstInventoryChanges("Type") = 3
                    rstInventoryChanges("Description") = "LoanID: " & rstLoanItemsPast("LoanID")
                    rstInventoryChanges.Update
                                        
                wrkCurrent.CommitTrans
                
                ' Update amount deductable for "unreturned", if we're doing lost and unreturned
                intDeductable = intDeductable - intRandomAmount
            End If
                
            If (intRandomType = 2 Or intRandomType = 3) And intDeductable <> 0 Then
                '''''''''''''''''
                ' For debugging
                '''''''''''''''''
                intLateMade = intLateMade + 1
                strLateLoans = strLateLoans & rstLoanItemsPast("LoanID") & " "
                strLateLoanItems = strLateLoanItems & rstLoanItemsPast("LoanItems.ID") & " "
                strLateLoanItemReturns = strLateLoanItemReturns & rstLoanItemReturns("ID") & " "
                
                ' Make a unreturned amount - simply deduct from the return
                intRandomAmount = gintRandom(1, intDeductable)
    
                ' Make returned amount less wasn't returned
                If rstLoanItemReturns("ReturnAmount") = intRandomAmount Then
                    rstLoanItemReturns.Delete
                Else
                    rstLoanItemReturns.Edit
                    rstLoanItemReturns("ReturnAmount") = rstLoanItemReturns("ReturnAmount") - intRandomAmount
                    rstLoanItemReturns.Update
                End If
                            
            End If
                
            intDone = intDone + 1
        
        End If
    Next i
    
initLoanItemsLostAndUnreturned_Exit:
    rstLoanItemsPast.Close
    rstLoanItemReturns.Close
    rstLoanItemTotalsNext.Close
    rstInventoryChanges.Close
    
    Set rstLoanItemsPast = Nothing
    Set rstLoanItemReturns = Nothing
    Set rstLoanItemTotalsNext = Nothing
    Set rstInventoryChanges = Nothing
    
    rstItems.MoveFirst
    
    ' for debugging
    Debug.Print (strLostLoans & "Total: " & intLostMade & Chr(10) _
              & strLateLoans & "Total: " & intLateMade & Chr(10) _
              & strLateLoanItems & Chr(10) _
              & strLateLoanItemReturns & Chr(10))
              
    
    Exit Sub
    
initLoanItemsLostAndUnreturned_Err:
    Resume initLoanItemsLostAndUnreturned_Exit
    
End Sub

Public Sub initLoanStatus(intPickupDateID As Integer)
    ''''''''''''''''''''''
    ' For debugging
    ''''''''''''''''''''''
    Dim strActiveLoans As String, strPartialLoans As String, strClosedLoans As String
    strActiveLoans = "Active LoanIDs: "
    strPartialLoans = "Partially Returned LoansIDs: "
    strClosedLoans = "Fully Returned LoanIDs: "
    
    
    ' Get the real return date
    Dim datReturnDate As Date
    datReturnDate = gdatReturnDate(intPickupDateID)
    
    ' Create a recordset of Loans with amount not returned, amount lost/damaged, and latest date returned
    Dim rstLoans As DAO.Recordset, rstLoansStatus As DAO.Recordset, strSQL As String
    
    strSQL = "Select LoanItems.LoanID, LoanItems.ID, LoanItems.Amount, LoanItems.AmountLostDamaged, " _
                    & " Nz(sum(LoanItemReturns.ReturnAmount),0) As TotalReturnAmount " _
                & "From LoanItems Left Join LoanItemReturns " _
                & "ON LoanItems.ID = LoanItemReturns.LoanItemID " _
                & "Group By LoanItems.LoanID, LoanItems.ID, LoanItems.Amount, LoanItems.AmountLostDamaged"
    strSQL = "Select LoanItems.LoanID, " _
                    & "sum(LoanItems.Amount) As TotalAmount, " _
                    & "sum(LoanItems.Amount) - sum(LoanItems.AmountLostDamaged) - sum(TotalReturnAmount) As TotalLoanUnreturnedAmount From (" _
                    & strSQL & ") Where LoanItems.LoanID IN (Select ID From Loans where PickupDateID = " & intPickupDateID & ") " _
                    & "Group By LoanID"
                    
    Set rstLoansStatus = CurrentDb.OpenRecordset(strSQL)
    Set rstLoans = CurrentDb.OpenRecordset("Loans")
    ' For each record
    While Not rstLoansStatus.EOF
        ' If none returned and none lost/damaged, set as active
        rstLoans.FindFirst ("ID = " & rstLoansStatus("LoanID"))
        rstLoans.Edit
        If rstLoansStatus("TotalAmount") = rstLoansStatus("TotalLoanUnreturnedAmount") Then
            rstLoans("Status") = 3 'active
            '''''''''''''''''''''''
            'For Debugging
            '''''''''''''''''''''''
            strActiveLoans = strActiveLoans & rstLoansStatus("LoanID") & " "
        ElseIf rstLoansStatus("TotalLoanUnreturnedAmount") > 0 Then
            rstLoans("Status") = 4 'partially unreturned
            '''''''''''''''''''''''
            'For Debugging
            '''''''''''''''''''''''
            strPartialLoans = strPartialLoans & rstLoansStatus("LoanID") & " "
        Else
            rstLoans("Status") = 8 ' fully returned, deposit still held
            '''''''''''''''''''''''
            'For Debugging
            '''''''''''''''''''''''
            strClosedLoans = strClosedLoans & rstLoansStatus("LoanID") & " "
        End If
        rstLoans.Update
        rstLoansStatus.MoveNext
    Wend
    
    
initLoanStatus_Exit:
    rstLoansStatus.Close
    Set rstLoansStatus = Nothing
    rstLoans.Close
    Set rstLoans = Nothing
    '''''''''''''''''''''''
    'For Debugging
    '''''''''''''''''''''''
    Debug.Print (strActiveLoans & Chr(10) & strPartialLoans & Chr(10) & strClosedLoans & Chr(10))
        
    Exit Sub

End Sub

Public Sub initLoanCharges(intPickupDateID As Integer)
    ''''''''''''''''''''''
    ' For debugging
    ''''''''''''''''''''''
    Dim strLateChargeLoans As String, strLostChargeLoans As String
    strLateChargeLoans = "Late Charge LoanIDs: "
    strLostChargeLoans = "Lost Charge LoansIDs: "
    
    Dim strSQL As String, strSQL2 As String, _
        rstChargesQuery As DAO.Recordset, rstBalanceChanges As DAO.Recordset
        
    ' Part #1 - Get LoanItems Categories
    strSQL = "Select * " _
           & "From LoanItems " _
           & "Inner Join Items On LoanItems.ItemID = Items.ID "
    ' Add latest return date
    strSQL = "Select LoanItems.LoanID, " _
                  & "LoanItems.Amount, " _
                  & "LoanItems.AmountLostDamaged, " _
                  & "Items.Category, " _
                  & "max(LoanItemReturns.ReturnDate) As LastReturnDate " _
            & "From (" & strSQL & ") As LoanItemsWCategory " _
            & "Left Join LoanItemReturns ON LoanItemsWCategory.LoanItems.ID = LoanItemReturns.LoanItemID " _
            & "Group By LoanItems.LoanID, " _
                     & "LoanItems.Amount, " _
                     & "LoanItems.AmountLostDamaged, " _
                     & "Items.Category"
    ' Part #2 - Get Loans with PickupDate Charge info - (filter only this pickupDate and fully returned loans)
    strSQL2 = "Select Loans.ID, " _
                   & "Loans.PickupDateID, " _
                   & "tblPickupDates.Cat1Replacement, " _
                   & "tblPickupDates.Cat2Replacement, " _
                   & "tblPickupDates.DailyLateCharge, " _
                   & "tblPickupDates.MaxLateCharge, " _
                   & "tblPickupDates.ReturnDate " _
            & "From Loans " _
            & "Inner Join tblPickupDates ON Loans.PickupDateID = tblPickupDates.ID " _
            & "Where tblPickupDates.ID = " & intPickupDateID & " And Loans.Status = 8"
    ' Join #1 and #2 and calculate LostDamagedFee
    strSQL = "Select *, " _
                & "LoanItems.AmountLostDamaged * " _
                & "IIf(Items.Category = 1, tblPickupdates.Cat1Replacement, tblPickupDates.Cat2Replacement) As LostDamagedFee " _
           & "From (" & strSQL & ") As LoanItemsQuery " _
           & "Inner Join (" & strSQL2 & ") As LoansQuery ON LoanItemsQuery.LoanID = LoansQuery.ID"
    
    ' Sum LostDamageFee and calculate LateFee, and store relevant dates and Charge
    strSQL = "Select LoanItems.LoanID, " _
                  & "max(LastReturnDate) As LoanCloseDate, " _
                  & "tblPickupDates.ReturnDate, " _
                  & "tblPickupDates.MaxLateCharge, " _
                  & "max( " _
                        & "(Nz(LastReturnDate, tblPickupDates.ReturnDate) - tblPickupDates.ReturnDate) * tblPickupDates.DailyLateCharge " _
                    & ") As TotalLateCharge, " _
                    & "sum(LostDamagedFee) As TotalLostDamagedFee " _
           & "From (" & strSQL & ") " _
           & "Group By LoanItems.LoanID, " _
                    & "tblPickupDates.ReturnDate, " _
                    & "tblPickupDates.MaxLateCharge "

    

    ' Open the charges query (with the above SQL)
    Set rstChargesQuery = CurrentDb.OpenRecordset(strSQL)
    ' Open the BalanceChanges table
    Set rstBalanceChanges = CurrentDb.OpenRecordset("BalanceChanges")
    ' For each record in the charges query
    While Not rstChargesQuery.EOF
        ' If there's a lostDamagedCharge
        If rstChargesQuery("TotalLostDamagedFee") <> 0 Then
            '''''''''''''''''''''''
            'For Debugging
            '''''''''''''''''''''''
            strLostChargeLoans = strLostChargeLoans & rstChargesQuery("LoanID") & " "
            
            ' Put it in the BalanceChanges table
            rstBalanceChanges.AddNew
            rstBalanceChanges("LoanID") = rstChargesQuery("LoanID")
            rstBalanceChanges("TypeID") = 2
            rstBalanceChanges("ChangeDate") = rstChargesQuery("ReturnDate")
            rstBalanceChanges("Amount") = rstChargesQuery("TotalLostDamagedFee") * -1
            rstBalanceChanges.Update
        End If
        ' If there's a late charge
        If rstChargesQuery("TotalLateCharge") <> 0 Then
            '''''''''''''''''''''''
            'For Debugging
            '''''''''''''''''''''''
            strLateChargeLoans = strLateChargeLoans & rstChargesQuery("LoanID") & " "
            
            ' Put it in the balance Changes table
            rstBalanceChanges.AddNew
            rstBalanceChanges("LoanID") = rstChargesQuery("LoanID")
            rstBalanceChanges("TypeID") = 1
            rstBalanceChanges("ChangeDate") = rstChargesQuery("LoanCloseDate")
            rstBalanceChanges("Amount") = gMin(rstChargesQuery("TotalLateCharge"), rstChargesQuery("MaxLateCharge")) * -1
            rstBalanceChanges.Update
        End If
        rstChargesQuery.MoveNext
    Wend
    
    '''''''''''''''''''''''
    'For Debugging
    '''''''''''''''''''''''
    Debug.Print (strLateChargeLoans & Chr(10) & strLostChargeLoans & Chr(10))
    
End Sub

Private Sub initLoanPayments(intPickupDateID As Integer)
    Dim strSQL As String, _
        rstLoansWithCharges As DAO.Recordset, _
        intRandPayment As Integer, _
        dblRandomPercent As Double, _
        curPayment As Currency, _
        intStatus As Integer, _
        intDepositStatus As Integer, _
        datPaymentDate As Date, _
        rstLoans As DAO.Recordset, _
        rstBalanceChanges As DAO.Recordset, _
        wrkCurrent As DAO.Workspace
            
    Set wrkCurrent = DBEngine.Workspaces(0)
        
        
    ' Create a recordset of Loans that have charges, for that pickupDate
    strSQL = "Select BalanceChanges.LoanID, " _
                  & "sum(BalanceChanges.Amount) As TotalCharges, " _
                  & "max(BalanceChanges.ChangeDate) As LastChargeDate " _
           & "From Loans " _
           & "Inner Join BalanceChanges ON Loans.ID = BalanceChanges.LoanID " _
           & "Where Loans.PickupDateID = " & intPickupDateID & " " _
           & "Group By BalanceChanges.LoanID"
           
    Set rstLoansWithCharges = CurrentDb.OpenRecordset(strSQL)
    
    ' Open the recordsets for making the changes
    Set rstLoans = CurrentDb.OpenRecordset("Loans")
    Set rstBalanceChanges = CurrentDb.OpenRecordset("BalanceChanges")
    
    
    While Not rstLoansWithCharges.EOF
        ' Randomly choose payment method
        intRandPayment = gintRandom(1, 3)
        ' Get a random percentage of amount paid (50% chance the full, otherwise random)
        dblRandomPercent = Rnd()
        dblRandomPercent = IIf(dblRandomPercent > 0.5, 1, dblRandomPercent * 2)
        ' figure out the payment amount
        curPayment = (dblRandomPercent * rstLoansWithCharges("TotalCharges")) * -1
        curPayment = CCur(CInt(curPayment)) 'Drop the cents
        curPayment = IIf(curPayment = 0, 1, curPayment) 'Make sure there's a minimal payment
        
        ' Choose the status accordingly
        intStatus = IIf(dblRandomPercent = 1, 8, 5) ' 5 is fully returned, with outstanding charges
        If intStatus = 8 Then
            intStatus = gintRandom(1, 2)
            intStatus = IIf(intStatus = 1, 6, 8) ' 6 is closed, 8 is fully returned, just deposit unreturned
        End If
        
        ' Choose deposit Status accordingly
        intDepositStatus = IIf(intStatus = 6, 4, 2)
        ' Choose the payment date
        datPaymentDate = rstLoansWithCharges("LastChargeDate") + gintRandom(0, 2)
        
        ' Make the payment, and update the status
        wrkCurrent.BeginTrans
            rstBalanceChanges.AddNew
            rstBalanceChanges("LoanID") = rstLoansWithCharges("LoanID")
            rstBalanceChanges("TypeID") = 5
            rstBalanceChanges("ChangeDate") = datPaymentDate
            rstBalanceChanges("Amount") = curPayment
            rstBalanceChanges("MethodID") = intRandPayment
            rstBalanceChanges.Update
            
            rstLoans.FindFirst ("ID = " & rstLoansWithCharges("LoanID"))
            rstLoans.Edit
            rstLoans("Status") = intStatus
            rstLoans("DepositStatus") = intDepositStatus
            rstLoans.Update
        wrkCurrent.CommitTrans
        
        rstLoansWithCharges.MoveNext
    Wend
End Sub

Private Sub setDemoDepositGiven(intPickupDateID As Integer)
    Dim strSQLDepositNeeded
    Dim rstLoans As DAO.Recordset, rstFees As DAO.Recordset
    Dim curDepositNeeded As Currency
    
    On Error GoTo setDemoDepositGiven_Err
    
    strSQLDepositNeeded = gstrSQLDepositNeeded(intPickupDateID)
    
    Set rstLoans = CurrentDb.OpenRecordset("Select ID, DepositGiven From Loans Where PickupDateID = " & intPickupDateID)
    Set rstFees = CurrentDb.OpenRecordset(strSQLDepositNeeded)
    While Not rstFees.EOF
        rstLoans.FindFirst ("ID = " & rstFees("LoanID"))
        rstLoans.Edit
        rstLoans("DepositGiven") = rstFees("DepositNeeded")
        rstLoans.Update
        rstFees.MoveNext
    Wend
    
setDemoDepositGiven_Exit:
    rstLoans.Close
    rstFees.Close
    Set rstLoans = Nothing
    Set rstFees = Nothing
    Exit Sub
    
setDemoDepositGiven_Err:
    Resume setDemoDepositGiven_Exit

End Sub


Public Function gstrSQLDepositNeeded(intPickupDateID As Integer)
    Dim strSQL As String
    
    ' First Deposit
    ' Build the strSQL for the deposit amounts in a way we can see what's happening
    
    ' #1 - join LoanItems info with which category
    strSQL = "SELECT LoanItems.LoanID, items.Category, LoanItems.Amount " _
            & "FROM Items INNER JOIN LoanItems " _
            & "ON Items.ID = LoanItems.ItemID"
    
    ' #2 - get sum of LoanItem amounts by Loan and category, for relevant pickupDateID
    strSQL = "SELECT Loans.ID As LoanID, LoanItemsCategories.Category, Sum(LoanItemsCategories.Amount) AS LoanItemTotal " _
            & "FROM Loans INNER JOIN " _
            & "(" & strSQL & ")" & " AS LoanItemsCategories " _
            & "ON Loans.ID = LoanItemsCategories.LoanID " _
            & "GROUP BY Loans.PickupdateID, Loans.ID, LoanItemsCategories.Category " _
            & "HAVING Loans.PickupDateID = " & intPickupDateID
            
    ' #3 - select sum(Amount) Result inner Join Loans group by category, for this pickupDateID
    strSQL = "SELECT LoanItemTotals.LoanID, " _
            & "Sum((Int((LoanItemTotals.LoanItemTotal / Categories.DepositGrouping)*-1)*-1)*Categories.DepositFee) As DepositNeeded " _
            & "FROM Categories INNER JOIN " _
            & "(" & strSQL & ")" & " AS LoanItemTotals " _
            & "ON Categories.ID = LoanItemTotals.Category " _
            & "GROUP BY LoanItemTotals.LoanID "

    gstrSQLDepositNeeded = strSQL
End Function


'''''''''''''''''''''''''''''''''''''''
' Utility helper functions
'''''''''''''''''''''''''''''''''''''''


Public Function IsInArray(i As Integer, Arr() As Integer) As Boolean
    IsInArray = False
    
    Dim ind As Integer
    For ind = 0 To UBound(Arr)
        If Arr(ind) = i Then
            IsInArray = True
        End If
    Next ind

End Function

Public Function gintRandom(intStart As Integer, intEnd As Integer) As Integer
' Returns a random integer in the range between (and including) the two arguments

    ' Figure out which is the first
    Dim intLower As Integer, intHigher As Integer
    intLower = IIf(intStart < intEnd, intStart, intEnd)
    intHigher = IIf(intLower = intStart, intEnd, intStart)

    Dim intRange As Integer
    intRange = (intHigher - intLower) + 1 ' number of possible values
    gintRandom = Int(Rnd * intRange)      ' 0 <= gintRandom < intRange
    gintRandom = gintRandom + intLower    ' move it to the proper scale

End Function

Public Function gbolRandom() As Boolean
    gbolRandom = CBool(gintRandom(-1, 0))
End Function

Public Function gMin(intOne As Integer, intTwo As Integer)
    gMin = intOne
    If intTwo < intOne Then
        gMin = intTwo
    End If
End Function

