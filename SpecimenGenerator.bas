Attribute VB_Name = "SpecimenGenerator"
Option Explicit
Sub GenerateSpecimen()
With FormSpecGen
    On Error GoTo ErrorHandler
    Dim HoldText      As String
    Dim path As String
    Dim objFolders As Object
    Set objFolders = CreateObject("WScript.Shell").SpecialFolders
    path = objFolders("mydocuments")
    path = path + "\SpecGen.ini"
    Set objFolders = Nothing
    
' Validate connectivity.  If not connected generate an error
CheckConnection:
    If Session.Connected = False Then
        MsgBox "You must be connected to ZBranch before running this macro"
        Exit Sub
    End If
    
' Initialize list boxes
    .lbFemaleNames.Clear
    .lbMaleNames.Clear
    .lbLastNames.Clear
    .lbPhyNames.Clear
    .lbTestNum.Clear
    .lbZipCode.Clear
    
' Initialize spin buttons
    .spbRedColor.Min = -8
    .spbRedColor.max = 264
    .spbRedColor.SmallChange = 8
    
    .spbBlueColor.Min = -8
    .spbBlueColor.max = 264
    .spbBlueColor.SmallChange = 8
    
    .spbGreenColor.Min = -8
    .spbGreenColor.max = 264
    .spbGreenColor.SmallChange = 8

' Read in user values
Open path For Input As #1
Do While (Not EOF(1))
    Line Input #1, HoldText 'TNG HoldText holds all of the text entered into the form
    Select Case Mid(HoldText, 1, 1)
        Case "1"
            .spbBlueColor.Value = val(Mid(HoldText, 2, 3))
            .spbGreenColor.Value = val(Mid(HoldText, 5, 3))
            .spbRedColor.Value = val(Mid(HoldText, 8, 3))
            
            .cbDOB.Tag = Format(Mid(HoldText, 11, 3), "###")
            .cbCollectTime.Tag = Format(Mid(HoldText, 14, 3), "###")
            .cbNPI.Tag = Format(Mid(HoldText, 17, 3), "###")
            .cbPhyID.Tag = Format(Mid(HoldText, 20, 3), "###")
            .cbSSN.Tag = Format(Mid(HoldText, 23, 3), "###")
            .cbPatID.Tag = Format(Mid(HoldText, 26, 3), "###")
            .cbVolume.Tag = Format(Mid(HoldText, 29, 3), "###")
            .cbPhyName.Tag = Format(Mid(HoldText, 32, 3), "###")
            .cbClinInfo.Tag = Format(Mid(HoldText, 35, 3), "###")
            .cbTubes.Tag = Format(Mid(HoldText, 38, 3), "###")
            .cbRespParty.Tag = Format(Mid(HoldText, 41, 3), "###")
            .cbCourtesy.Tag = Format(Mid(HoldText, 44, 3), "###")
            .cbControlNum.Tag = Format(Mid(HoldText, 47, 3), "###")
            
            If Len(HoldText) > 50 Then
                .cbHeight.Tag = Format(Mid(HoldText, 50, 3), "###")
                .cbWeight.Tag = Format(Mid(HoldText, 53, 3), "###")
                .tbEmail.Tag = Format(Mid(HoldText, 56, 15), "###############") 'TNG
            End If
        
        Case "L"
            .lbLastNames.AddItem (Right(HoldText, Len(HoldText) - 1))
        
        Case "M"
            .lbMaleNames.AddItem (Right(HoldText, Len(HoldText) - 1))
        
        Case "F"
            .lbFemaleNames.AddItem (Right(HoldText, Len(HoldText) - 1))
        
        Case "P"
            .lbPhyNames.AddItem (Right(HoldText, Len(HoldText) - 1))
        
        Case "T"
            .lbTestNum.AddItem (Right(HoldText, Len(HoldText) - 1))
    
        Case "Z"
            .lbZipCode.AddItem (Right(HoldText, Len(HoldText) - 1))
        
        Case "D"
            Select Case Mid(HoldText, 2, 1)
                Case "1"
                    .tbDOBValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "2"
                    .tbTimeValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "3"
                    .tbNPIValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "4"
                    .tbPhyIdValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "5"
                    .tbSSNValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "6"
                    .tbPatIdValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "7"
                    .tbVolumeValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "8"
                    .tbHeightValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "9"
                    .tbPhyNameValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "A"
                    .tbTubeValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "B"
                    .tbControlNumValue.Value = Right(HoldText, Len(HoldText) - 2)
                Case "C"
                    .tbWeightValue.Value = Right(HoldText, Len(HoldText) - 2)
            End Select

    End Select
Loop
Close #1
    If .lbZipCode.ListCount = -1 Then
        .lbZipCode.AddItem ("27215")
    End If

NoFile:
' Accept user input.
GetInput:
    .Show

Terminate:
    Unload FormSpecGen
    Exit Sub
        
End With
ErrorHandler:
    Select Case Err.Number
    Case 53
        FormSpecGen.spbBlueColor.Value = 255
        FormSpecGen.spbGreenColor.Value = 192
        FormSpecGen.spbRedColor.Value = 192
        FormSpecGen.cbDOB.Tag = ""
        FormSpecGen.cbCollectTime.Tag = ""
        FormSpecGen.cbNPI.Tag = ""
        FormSpecGen.cbPhyID.Tag = ""
        FormSpecGen.cbSSN.Tag = ""
        FormSpecGen.cbPatID.Tag = ""
        FormSpecGen.cbVolume.Tag = ""
        FormSpecGen.cbPhyName.Tag = ""
        FormSpecGen.cbClinInfo.Tag = ""
        FormSpecGen.cbTubes.Tag = ""
        FormSpecGen.cbRespParty.Tag = ""
        FormSpecGen.cbCourtesy.Tag = ""
        FormSpecGen.cbControlNum.Tag = ""
        FormSpecGen.tbEmail.Tag = ""    'TNG
        FormSpecGen.lbLastNames.AddItem ("Testing")
        FormSpecGen.lbMaleNames.AddItem ("Labcorp")
        FormSpecGen.lbFemaleNames.AddItem ("Labcorp")
        FormSpecGen.lbPhyNames.AddItem ("Physician")
        FormSpecGen.lbTestNum.AddItem ("001818")
        Resume NoFile
    
    Case Else
        Session.MsgBox Err.Description, vbExclamation + vbOKOnly
    End Select
End Sub

Sub GenerateSpecimens()
With FormSpecGen
    Dim HoldText      As String
    Dim TestVal       As Integer
    Dim a             As Variant

    Dim BillCode      As String * 2
    Dim CarrierCode   As String * 5
    Dim ClinInfo1     As String * 26
    Dim ClinInfo2     As String * 26
    Dim ClinInfo3     As String * 26
    Dim CollectDate   As String * 8
    Dim collectTime   As String * 4
    Dim CPU           As String * 1
    Dim CurrentTime
    Dim CurrentHr     As Integer
    Dim CurrentMn     As Integer
    Dim MinStart      As Integer
    Dim controlNum    As String * 10
    Dim DateOfBirth   As String * 8
    Dim DiagCode(6)   As String * 6
    Dim Fasting       As String * 1
    Dim HomeTray      As String * 8
    Dim MedicareNum   As String * 10
    Dim OrdTest       As String * 6
    Dim patID         As String * 20
    Dim PatLast       As String * 25
    Dim PatFirst      As String * 15
    Dim PatMiddle     As String * 15
    Dim phyID         As String * 10
    Dim PhyNameLast   As String * 9
    Dim PhyNameFirst  As String * 1
    Dim PhySign       As String * 1
    Dim PSTID         As String * 5
    Dim ReasonCode    As String * 1
    Dim ReasonText    As String * 6
    Dim RespPhone     As String * 10
    Dim RespAddress   As String * 35
    Dim sex           As String * 1
    Dim SplitSample   As String * 1
    Dim SSN           As String * 9
    Dim volume        As String * 4
    Dim UPIN          As String * 10
    Dim PatZip        As String * 5
    
    Dim ranNum As Integer
    Dim RanNum2 As Integer
    Dim Looper As Integer
    Dim currSpec As Integer
    Dim CurrLine As Integer
    Dim CurrZip As Integer
    Dim Response As Integer
    
'TNG enable sys1
' Check current CPU, make sure not on production
'    If checkCPU(.tbAccession.Text) = 1 Then
'        MsgBox "Invalid accession code, lives on system 1", vbOKOnly, "E R R O R !"
'        Exit Sub
'    End If
    
' Go to order entry screen
    GoToMainMenu
    GoToMenu "15", "Activate Physician-Upin"
    
' Set collection date to current date
    CollectDate = Mid(Date$, 1, 2) + Mid(Date$, 4, 2) + Right(Date$, 4)
    
' Initialize loop controllers for loop options.
    CurrLine = 0
    CurrZip = 0
    
' Initialize random number generator
    Randomize
            
' Loop through all specimens
    For currSpec = val(.tbSequence.Text) To val(.tbEndSeq.Text)
    
' Generate a random bill code.  Do this first so bill code specifc requirements can be met
        If .opBillCodeSpecify.Value = True Then
            BillCode = .tbBillCode
        Else
            BillCode = "03"
            'ranNum = Int(Rnd(1) * 4)
            'Select Case ranNum 'TNG commented this out bc never use it. Always want 03
            '    Case 0: BillCode = "03"
            '    Case 1: BillCode = "04"
            '    Case 2: BillCode = "05"
            '    Case 3: BillCode = "XI"
            'End Select
        End If
        
' Get email
        Dim Email As String
        Email = .tbEmail.Value
        
' Generate a control number
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbControlNumPercent.Text) Then
            If .tbControlNumValue.Value > "" Then
                controlNum = .tbControlNumValue.Value
            Else
                HoldText = ""
                TestVal = 0
                controlNum = "0" + Format(Str(Int(Rnd(1) * 99999999)), "00000000")
                If controlNum < "020000000" Then
                    For Looper = 1 To 9
                        If (Looper / 2) = Int(Looper / 2) Then
                            HoldText = HoldText + Format(Str(val(Mid(controlNum, Looper, 1))), "0")
                        Else
                            HoldText = HoldText + Format(Str(val(Mid(controlNum, Looper, 1) * 2)), "#0")
                        End If
                    Next Looper
                    
                    For Looper = 1 To Len(HoldText)
                        TestVal = TestVal + val(Mid(HoldText, Looper, 1))
                    Next Looper
                    TestVal = 10 - val(Right(Str(TestVal), 1))
                    controlNum = controlNum + Format(TestVal, "0")
                Else
                    controlNum = controlNum + Format(Str(val(controlNum Mod 7)), "0")
                End If
            End If
        Else
            controlNum = ""
        End If
        
' Determine patient gender
        If .opGenderRandom.Value = True Then
            ranNum = Int(Rnd(1) * 3) '*TNG this doesn't seem right.
        Else
            If .opGenderFemale.Value = True Then
                ranNum = 1
            Else
                ranNum = 0
            End If
        End If
        Select Case ranNum
            Case 0:  sex = "M"
            Case 1:  sex = "F"
            Case 2:  sex = "N"
        End Select
        
' Generate a random name, or use name specified in overide
        If .opSpecifyName.Value = True Then
            PatLast = .tbSpecifyLastName.Text
            PatFirst = .tbSpecifyFirstName.Text
            PatMiddle = .tbSpecifyMiddleName.Text
        Else
            Select Case sex
            Case "N":
                ranNum = Int(Rnd(1) * 2)
                If ranNum = 0 Then
                    PatFirst = PullFromList(.lbMaleNames)
                    'PatFirst = .lbMaleNames.List(Int(Rnd(1) * .lbMaleNames.ListCount))
                Else
                    'PatFirst = .lbFemaleNames.List(Int(Rnd(1) * .lbFemaleNames.ListCount))
                    PatFirst = PullFromList(.lbFemaleNames)
                End If
            
            Case "M":
                    PatFirst = PullFromList(.lbMaleNames)
            
            Case "F":
                    PatFirst = PullFromList(.lbFemaleNames)
            End Select
            PatLast = .lbLastNames.List(Int(Rnd(1) * .lbLastNames.ListCount))
            PatMiddle = Chr(Int(Rnd(1) * 26) + 65)
        End If
        
' Generate a random date of birth if option selected
        ranNum = Int(Rnd(1) * 99) + 1
        If (ranNum < val(.tbDOBPercent.Text)) Or (BillCode = "05") Then
            If .tbDOBValue.Value <> "" Then
                If Len(.tbDOBValue.Value) = 10 Then
                    DateOfBirth = Mid(.tbDOBValue.Value, 1, 2) + _
                                  Mid(.tbDOBValue.Value, 4, 2) + _
                                  Mid(.tbDOBValue.Value, 7, 4)
                Else
                    DateOfBirth = Mid(.tbDOBValue.Value, 1, 8)
                End If
            Else
                ranNum = Int(Rnd(1) * 12) + 1
                DateOfBirth = Format(ranNum, "00")
                If ranNum = 2 Then
                    ranNum = Int(Rnd(1) * 28) + 1
                ElseIf ranNum = 4 Or ranNum = 6 Or ranNum = 9 Or ranNum = 11 Then
                    ranNum = Int(Rnd(1) * 30) + 1
                Else
                    ranNum = Int(Rnd(1) * 31) + 1
                End If
                DateOfBirth = Mid(DateOfBirth, 1, 2) + Format(ranNum, "00")
                
                ranNum = Int(Rnd(1) * 100)
                If ranNum < 4 Then
                    DateOfBirth = Mid(DateOfBirth, 1, 4) + "20" + Format(ranNum, "00")
                Else
                    DateOfBirth = Mid(DateOfBirth, 1, 4) + "19" + Format(ranNum, "00")
                End If
            End If
        Else
            DateOfBirth = "00000000"
        End If
        
' Generate a random collection time, if option selected
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbCollectTimePercent.Text) Then
            If (.tbTimeValue.Value <> "") Then
                If Len(.tbTimeValue.Value) = 5 Then
                    collectTime = Mid(.tbTimeValue.Value, 1, 2) + Mid(.tbTimeValue.Value, 4, 2)
                Else
                    collectTime = Mid(.tbTimeValue.Value, 1, 4)
                End If
            Else
                CurrentTime = Time
                MinStart = InStr(1, CurrentTime, ":") + 1
                
                CurrentHr = val(Mid(CurrentTime, 1, 2))
                CurrentMn = val(Mid(CurrentTime, MinStart, 2))
                
                ranNum = Int(Rnd(1) * 24)
                If ranNum > CurrentHr Then
                    ranNum = CurrentHr
                End If
                
                RanNum2 = Int(Rnd(1) * 60)
                If ranNum = CurrentHr And RanNum2 > CurrentMn Then
                    RanNum2 = CurrentMn - 1
                End If
                collectTime = Format(ranNum, "00") + Format(RanNum2, "00")
            End If
        Else
            collectTime = "0000"
        End If
        
' Determine if patient is fasting
        If .opFastingRandom.Value = True Then
            ranNum = Int(Rnd(1) * 3)
            Select Case ranNum
               Case 0: Fasting = "Y"
               Case 1: Fasting = "N"
               Case 2: Fasting = " "
            End Select
        ElseIf .opFastingNo.Value = True Then
            Fasting = "N"
        Else
            Fasting = "Y"
        End If
        
' Generate a random UPIN number, 2 to 10 characters in length
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbNPIPercent.Text) Then
            If .tbNPIValue.Value <> "" Then
                UPIN = .tbNPIValue.Value
            Else
                ranNum = Int(Rnd(1) * 2)
                If ranNum < 1 Then
                    UPIN = Format(Int(Rnd(1) * 10), "0")
                Else
                    UPIN = Chr(Int(Rnd(1) * 26) + 65)
                End If
                Looper = 2
                Do
                    ranNum = Int(Rnd(1) * 10)
                    If ranNum > 8 Then
                        Looper = 11
                    ElseIf ranNum > 4 Then
                        UPIN = UPIN + Format(Int(Rnd(1) * 10), "0")
                    Else
                        UPIN = UPIN + Chr(Int(Rnd(1) * 26) + 65)
                    End If
                    Looper = Looper + 1
                Loop Until Looper > 10
            End If
        Else
            UPIN = ""
        End If
        
' Generate a random phyisician ID, 2 to 10 characters in length
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbPhyIDPercent.Text) Then
            If .tbPhyIdValue.Value <> "" Then
                phyID = .tbPhyIdValue.Value
            Else
                ranNum = Int(Rnd(1) * 2)
                If ranNum < 1 Then
                    phyID = Format(Int(Rnd(1) * 10), "0")
                Else
                    phyID = Chr(Int(Rnd(1) * 26) + 65)
                End If
                
                Looper = 2
                Do
                    ranNum = Int(Rnd(1) * 10)
                    If ranNum > 8 Then
                        Looper = 11
                    ElseIf ranNum > 4 Then
                        phyID = phyID + Format(Int(Rnd(1) * 10), "0")
                    Else
                        phyID = phyID + Chr(Int(Rnd(1) * 26) + 65)
                    End If
                    Looper = Looper + 1
                Loop Until Looper > 10
            End If
        Else
            phyID = ""
        End If
        
' Generate a random social security number
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbSSNPercent.Text) Then
            If .tbSSNValue.Value <> "" Then
                If Len(.tbSSNValue.Value) = 11 Then
                    SSN = Mid(.tbSSNValue.Value, 1, 3) + Mid(.tbSSNValue.Value, 5, 2) + Mid(.tbSSNValue.Value, 8, 4)
                Else
                    SSN = Format(.tbSSNValue.Value, "000000000")
                End If
            Else
                SSN = Format(Int(Rnd(1) * 999999999), "000000000")
            End If
        Else
            SSN = ""
        End If

' Generate a random patient ID
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbPatIDPercent.Text) Then
            If .tbPatIdValue.Value <> "" Then
                patID = .tbPatIdValue.Value
            Else
                ranNum = Int(Rnd(1) * 2)
                If ranNum < 1 Then
                    patID = Format(Int(Rnd(1) * 10), "0")
                Else
                    patID = Chr(Int(Rnd(1) * 26) + 65)
                End If
                
                Looper = 2
                Do
                    ranNum = Int(Rnd(1) * 10)
                    If ranNum > 8 Then
                        Looper = 21
                    ElseIf ranNum > 4 Then
                        patID = patID + Format(Int(Rnd(1) * 10), "0")
                    Else
                        patID = patID + Chr(Int(Rnd(1) * 26) + 65)
                    End If
                    Looper = Looper + 1
                Loop Until Looper > 20
            End If
        Else
            patID = ""
        End If
        
' Generate Total Volume
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbVolumePercent.Text) Then
            If .tbVolumeValue.Value <> "" Then
                volume = Format(.tbVolumeValue.Value, "0000")
            Else
                volume = Format(Int(Rnd(1) * 9999), "0000")
            End If
        Else
            volume = "0000"
        End If
        
' Generate Physician name
        ranNum = Int(Rnd(1) * 99) + 1
        If (ranNum < val(.tbPhyNamePercent.Text)) Or (BillCode = "05") Then
            If .tbPhyNameValue.Value <> "" Then
                a = split(.tbPhyNameValue.Value, " ")
                If UBound(a) = 0 Then
                    PhyNameLast = a(0)
                    PhyNameFirst = ""
                Else
                    If Len(a(0)) = 1 Then
                        PhyNameFirst = a(0)
                        PhyNameLast = a(1)
                    Else
                        PhyNameFirst = a(1)
                        PhyNameLast = a(0)
                    End If
                End If
            
            Else
                PhyNameLast = .lbPhyNames.List(Int(Rnd(1) * .lbPhyNames.ListCount))
                PhyNameFirst = Chr(Int(Rnd(1) * 26) + 65)
            End If
        Else
            PhyNameLast = ""
            PhyNameFirst = ""
        End If
        

        CarrierCode = Format(Int(Rnd(1) * 99999), "00000")
        MedicareNum = Format(Int(Rnd(1) * 9999999999#), "0000000000")
        
' Generate Bill code information
' Place courtesy copy first if requested
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbCourtesyPercent.Text) Then
            ClinInfo1 = "CC:" + .tbAccount.Text
        ElseIf ranNum < val(.tbClinInfoPercent.Text) Then
            ClinInfo1 = "Clinical information 1"
        Else
            ClinInfo1 = ""
        End If
        
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbClinInfoPercent.Text) Then
            ClinInfo2 = "Clinical information 2"
        Else
            ClinInfo2 = ""
        End If
                
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbClinInfoPercent.Text) Then
            ClinInfo3 = "Clin Info3"
        Else
            ClinInfo3 = ""
        End If
        
' Generate responsible party infomation
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbRespPartyPercent.Text) Or BillCode = "04" Or BillCode = "05" Then
            RespPhone = "555436" + Format(Int(Rnd(1) * 9999), "0000")
            RespAddress = Format(Int(Rnd(1) * 9999), "###0")
            ranNum = Int(Rnd(1) * 4)
            Select Case ranNum
                Case 0: RespAddress = RespAddress + " Somewhere "
                Case 1: RespAddress = RespAddress + " Nowhere "
                Case 2: RespAddress = RespAddress + " Everywhere "
                Case 3: RespAddress = RespAddress + " Random "
            End Select
            ranNum = Int(Rnd(1) * 4)
            Select Case ranNum
                Case 0: RespAddress = RespAddress + "Road"
                Case 1: RespAddress = RespAddress + "Court"
                Case 2: RespAddress = RespAddress + "Boulevard"
                Case 3: RespAddress = RespAddress + "Street"
            End Select
        Else
            RespPhone = "0000000000"
            RespAddress = ""
        End If
        
' Generate order item
        If .opTestLoop.Value = True Then
            OrdTest = .lbTestNum.List(CurrLine)
            CurrLine = CurrLine + 1
            If CurrLine = .lbTestNum.ListCount Then
                CurrLine = 0
            End If
        ElseIf .opSpecifyTest.Value = True Then
            OrdTest = .tbSpecifyTest.Text
        Else
            ranNum = Int(Rnd(1) * .lbTestNum.ListCount)
            OrdTest = .lbTestNum.List(ranNum)
        End If

' Generate a random zip code
        If Trim(RespAddress) = "" Then
            PatZip = ""
        Else
            If .opLoopZip.Value = True Then
                PatZip = .lbZipCode.List(CurrZip)
                CurrZip = CurrZip + 1
                If CurrZip = .lbZipCode.ListCount Then
                    CurrZip = 0
                End If
            ElseIf .opSpecifyZip = True Then
                PatZip = .tbSpecifyZip.Text
            Else
                ranNum = Int(Rnd(1) * .lbZipCode.ListCount)
                PatZip = .lbZipCode.List(ranNum)
            End If
        End If

' Enter general specimen information
        SendText .tbForm.Text
        If Len(.tbForm.Text) > 3 Then
            TabField 1
        Else
            TabField 2
        End If
        SendText controlNum
        TabField 2
        SendText .tbJul.Text + .tbAccession.Text + Format(currSpec, "0000") + .tbBrother.Text + .tbAccount.Text + "+"
        TestVal = WaitForResponse("Refer to Request", "Duplicate record")
        If TestVal = 2 Then
            GoTo nextSpecimen
        End If
        
' Generate form 841 layout '
        SendText PatLast
        SendText PatFirst
        SendText PatMiddle
        SendText sex
        SendText DateOfBirth
        SendText Fasting
        
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbHeightPercent.Text) Then
            If .tbHeightValue.Value <> "" Then
                SendText Format(.tbHeightValue.Value, "00000")
            Else
                ranNum = Int(Rnd(1) * 3) + 3
                SendText Format(ranNum, "0")
                ranNum = Int(Rnd(1) * 11) + 1
                SendText Format(ranNum, "00")
                ranNum = Int(Rnd(1) * 99) + 1
                SendText Format(ranNum, "00")
            End If
        Else
            TabField (3)
        End If
        
        
        ranNum = Int(Rnd(1) * 99) + 1
        If ranNum < val(.tbWeightPercent.Text) Then
            If .tbWeightValue.Value <> "" Then
                If (Len(.tbWeightValue) = 6) Then
                    SendText Mid(.tbWeightValue.Value, 1, 3) + "0" + Mid(.tbWeightValue.Value, 5, 2)
                Else
                    If (Len(.tbWeightValue.Value) = 5) Then
                        SendText Mid(.tbWeightValue.Value, 1, 3) + Mid(.tbWeightValue.Value, 5, 1)
                        TabField (1)
                    Else
                        SendText Format(.tbWeightValue.Value, "0000")
                        TabField (1)
                    End If
                End If
                    
            Else
                ranNum = Int(Rnd(1) * 150) + 50
                SendText Format(ranNum, "000")
                SendText "0"
                ranNum = Int(Rnd(1) * 16)
                SendText Format(ranNum, "00")
            End If
        Else
            TabField (3)
        End If
        
        TabField (5) ' Tab past waist, blood pressure and pulse
        
        SendText CollectDate
        SendText collectTime
        TabField (1) ' Tab past UPIN, place in NPI just a guess since they are now separate fields
        SendText UPIN
        SendText phyID
        SendText SSN
        
        TabField (3) ' Tab past the new received date
        SendText patID
        SendText volume
        SendText PhyNameLast
        SendText PhyNameFirst
        SendText "Y"
        SendText BillCode
        'If Trim(.tbForm) <> "1500" Then
        '    SendText ClinInfo1
        'Else: SendText Email
        'End If
        SendText Email
        SendText ClinInfo1
        SendText ClinInfo2
        SendText ClinInfo3
        SendText RespPhone
        SendText "N"
        SendText OrdTest
        Response = WaitForResponse("Diagnosis Free Text", "Blood Lead Data", "Specimen Order Entry")
        
        If Response = 3 Then
            GoTo nextSpecimen
        End If
        
        If Response = 2 Then
            Response = WaitForResponse("Diagnosis Free Text", "Specimen Order Entry")
        End If
                
        If Response = 2 Then
            GoTo nextSpecimen
        End If
                
        DownLine 3
        If Trim(RespAddress) <> "" Then
            TabField 2
            SendText "1" + RespAddress
            TabField 3
            SendText PatZip
        Else
            DownLine 2
            TabField 2
        End If
        
        If BillCode = "XI" Then
            TabField 6
            SendText CarrierCode
        ElseIf BillCode = "05" Then
            SendText MedicareNum
        End If

tryagain:
        TestVal = WaitForResponse("Zip Duplication Screen", "ABN", "Order has been processed", "SUSPEND REASONS")
        Select Case TestVal
            Case 1:
                SendText "01"
                WaitForResponse "Diagnosis Free Text"
                GoTo tryagain
                
            Case 2:
                SendText "Y"
                GoTo tryagain
        End Select
            
nextSpecimen:
    Next currSpec
    ' Close (7)  What is this?
    
End With
End Sub

Sub ResetSpecimens()
With FormSpecGen
    Dim MidInit  As String * 1
    Dim currSpec As Integer
    Dim CheckError As String
        
    GoToMainMenu
    GoToMenu "11", "3.  Delete a Specimen"
            
    For currSpec = val(.tbSequence.Text) To val(.tbEndSeq.Text)
        SendText .tbJul.Text + .tbAccession.Text + Format(currSpec, "0000") + .tbBrother.Text + "2"
        WaitForReturn
        
        MidInit = ReadScreen(2, 62, 15)
        If MidInit >= "Z" Or MidInit < "A" Then
            MidInit = "A"
        Else
            MidInit = Chr(Asc(MidInit) + 1)
        End If
        
        TabField 2
        SendText MidInit
        WaitForReturn
        
        CheckError = ReadScreen(23, 28, 8)
        If CheckError = "WARNING!" Then
            WaitForReturn
        End If
        WaitForReturn
        BackOneScreen "Delete a Specimen"
    Next currSpec

End With
End Sub

Sub ResultSpecimens()
With FormSpecGen

    Dim currResult As String * 7
    Dim CurrTest   As String * 6
    Dim CurrLab    As String * 2
     
    Dim currSpec   As Integer
    Dim CurrLine   As Integer
    Dim MoreTests  As String * 1
    
    Dim resultType As String
    Dim NumResult  As Integer
    Dim result     As String * 7
           
    Randomize
    GoToMenu "10", "Comm/Abrv"
    For currSpec = val(.tbSequence.Text) To val(.tbEndSeq.Text)
        SendText .tbJul.Text + .tbAccession.Text + Format(currSpec, "0000") + .tbBrother.Text
        WaitForResponse "Order  Test"
        MoreTests = "Y"
        
        While MoreTests = "Y"
            For CurrLine = 2 To 22
                CurrTest = ReadScreen(CurrLine, 7, 6)
                resultType = ReadScreen(CurrLine, 32, 1)
                currResult = ReadScreen(CurrLine, 21, 7)
                CurrLab = ReadScreen(CurrLine, 71, 2)
                
                If CurrTest <> "      " Then
                    If ((.cbOverwrite.Value = True) Or (currResult = "       ")) And (CurrLab <> "  ") Then
                        Select Case resultType
                            Case "A": result = "N      "
                            Case "L": result = ">5000  "
                            Case "C": result = "COMMNT "
                            Case "0": NumResult = Int(Rnd(1) * 99): result = Format(NumResult, "######0")
                            Case "1": NumResult = Int(Rnd(1) * 99): result = Format(NumResult / 10, "####0.0")
                            Case "2": NumResult = Int(Rnd(1) * 999): result = Format(NumResult / 100, "###0.00")
                            Case "3": NumResult = Int(Rnd(1) * 9999): result = Format(NumResult / 1000, "##0.000")
                            Case "4": NumResult = Int(Rnd(1) * 99999): result = Format(NumResult / 10000, "#0.0000")
                        End Select
                        ClearField
                        SendText result
                        DownLine 1
                        backtab 1
                    Else
                        DownLine 1
                    End If
                Else
                    CurrLine = 22
                End If
            Next CurrLine
            
            MoreTests = ReadScreen(0, 75, 1)
            If MoreTests = "Y" Then
                WaitForReturn
            End If
        Wend
        WaitForReturn
        BackOneScreen "RESULT/INQUIRY MAINTENANCE"
    Next currSpec
End With
End Sub

Function PullFromList(lb As ListBox) As String
    PullFromList = lb.List(Int(Rnd(1) * lb.ListCount))
End Function


Function checkCPU(acc As String)
With Session
    GoToMainMenu
    GoToMenu "08", "DIFF/DRUG/NORM/TEST/TIXI/"
    SendText "ACCS" + acc
    WaitForReturn
    checkCPU = Session.GetText(23, 20, 23, 20)
End With
End Function

