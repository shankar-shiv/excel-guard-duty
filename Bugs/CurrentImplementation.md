These are the bugs existing in the current implementation.

1. Under the duty committments sheet, the letter 'K' is missing.
2. In the SlotTransferrer macro
    ```Csharp
        // Original
        Dim destiCurrentRow as Integer

        // =============================================================

        // New
        Dim destiCurrentRow as Long
    ```
3. Under the DutyPersonnel Class module
    ``` Csharp
        // Original
        commitmentsNameCol = 4

        // =============================================================

        // New
        commitmentsNameCol = 2
    ```
4. Under DutySlots module
   ``` csharp
    // Original

    // This is changed because we don't want to use extra stack memory for 
    // useless storage.
    Dim personnel(255) as DutyPersonnel
    Dim slots(255) as DutySlot

    line 20 : // Wrong Variable name, should be date not day
    dayCol = 1

    line 38 : 
    If (DutySlots.getColHeader(sCol) = "ACTUAL (ARMED)") Then
        sArmed = True
    Else
        sArmed = False
    End If

    line 188 : 
    // This is not a logical mistake rather a grammatical mistake.
    currentDay = slots(i).day // It should be currentDate 
    // because the column points to DATE column : Integer.  
    
    // =============================================================

    // New
    Dim personnel(100) as DutyPersonnel

    Dim slots(100) as DutySlot

    line 20 : 
    dayCol = 2 // should be 2

    line 38 : 
    If (DutySlots.getColHeader(sCol) = "ACTUAL (W/ARMS)") Then
        sArmed = True
    Else
        sArmed = False
    End If

    line 188 : 
    // This is not a logical mistake rather a grammatical mistake.
    currentDate = slots(i).date // It should be currentDate 
    // because the column points to DATE column : Integer. Renaming to the 
    // proper variable may cause unexpected issues later which may be hard to debug. 
    // This is not a serious issue, can be avoided.     
   ```