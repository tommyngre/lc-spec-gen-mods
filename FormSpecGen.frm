VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSpecGen 
   Caption         =   "Specimen Generator Version 2.0.6"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15315
   OleObjectBlob   =   "FormSpecGen.frx":0000
End
Attribute VB_Name = "FormSpecGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Const VERSION = "2.0.7"

Private Sub bnCancel_Click()
With FormSpecGen
    Select Case .Tag
        Case ""
            .Tag = "C"
            .Hide
            Exit Sub
                        
        Case "D"
            .frameAccession.Visible = False
                        
        Case "L"
            .frameAccession.Visible = False
            .cbOverwrite.Visible = False
                                   
        Case "O"
            .frameColor.Visible = False
            .framePercents.Visible = False
            .frameNames.Visible = False
            .imgLogo.Visible = True
                        
        Case "R"
            .frameAccession.Visible = False
               
        Case "S"
            .frameOEInfo.Visible = False
            .frameAccession.Visible = False
            .frameControls.Visible = False
            .framePercents.Visible = False
    End Select

    .caption = "Specimen Generator Version " + VERSION
    .Tag = ""
    .frameOption.Visible = True
    
    .bnCancel.caption = "Exit"
    .bnCancel.Accelerator = "X"
    .bnOk.Visible = False
End With
End Sub

Private Sub bnDelete_Click()
With FormSpecGen
    .Tag = "D"
    .caption = "Specimen Generator Version " + VERSION + "-- Delete Specimens"
    
    .bnOk.caption = "Delete"
    .bnOk.Accelerator = "D"
    .bnOk.Visible = True

    .bnCancel.caption = "Cancel"
    .bnCancel.Accelerator = "C"
    
    .frameOption.Visible = False
    
    .frameAccession.Top = 198
    .frameAccession.Left = 340
    .frameAccession.Visible = True
   
    .tbJul.SetFocus
    
End With
End Sub


Private Sub bnOk_Click()
With FormSpecGen
    Const ValidBill = "03 04 05 XI XR"
    Dim Looper As Integer
    Dim Buffer As String
        
    If (Not IsNumeric(.tbClinInfoPercent.Text)) Or (val(.tbClinInfoPercent.Text) > 100) Then
        Session.MsgBox "Invalid clinical info percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbCollectTimePercent.Text)) Or (val(.tbCollectTimePercent.Text) > 100) Then
        Session.MsgBox "Invalid clinical info percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
        
    If (Not IsNumeric(.tbControlNumPercent.Text)) Or (val(.tbControlNumPercent.Text) > 100) Then
        Session.MsgBox "Invalid control number percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
        
    If (Not IsNumeric(.tbCourtesyPercent.Text)) Or (val(.tbCourtesyPercent.Text) > 100) Then
        Session.MsgBox "Invalid courtesy copy percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbDOBPercent.Text)) Or (val(.tbDOBPercent.Text) > 100) Then
        Session.MsgBox "Invalid date of birth percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbNPIPercent.Text)) Or (val(.tbNPIPercent.Text) > 100) Then
        Session.MsgBox "Invalid NPI percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbPatIDPercent.Text)) Or (val(.tbPatIDPercent.Text) > 100) Then
        Session.MsgBox "Invalid patient ID percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbPhyIDPercent.Text)) Or (val(.tbPhyIDPercent.Text) > 100) Then
        Session.MsgBox "Invalid physician ID percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbPhyNamePercent.Text)) Or (val(.tbPhyNamePercent.Text) > 100) Then
        Session.MsgBox "Invalid physician name percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbRespPartyPercent.Text)) Or (val(.tbRespPartyPercent.Text) > 100) Then
        Session.MsgBox "Invalid responsible party percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If

    If (Not IsNumeric(.tbSSNPercent.Text)) Or (val(.tbSSNPercent.Text) > 100) Then
        Session.MsgBox "Invalid social security percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If

    If (Not IsNumeric(.tbTubesPercent.Text)) Or (val(.tbTubesPercent.Text) > 100) Then
        Session.MsgBox "Invalid tube percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbVolumePercent.Text)) Or (val(.tbVolumePercent.Text) > 100) Then
        Session.MsgBox "Invalid volume percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbHeightPercent.Text)) Or (val(.tbHeightPercent.Text) > 100) Then
        Session.MsgBox "Invalid height percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    If (Not IsNumeric(.tbWeightPercent.Text)) Or (val(.tbWeightPercent.Text) > 100) Then
        Session.MsgBox "Invalid weight percentage", vbOKOnly, "E R R O R"
        Exit Sub
    End If
    
    Select Case Mid(.Tag, 1, 1)
        Case "O"
            If .lbTestNum.ListCount = 0 Then
                Session.MsgBox "You must have at least one test number", vbOKOnly, "E R R O R"
                Exit Sub
            End If
            
            If .lbLastNames.ListCount = 0 Then
                Session.MsgBox "You must have at least one last name", vbOKOnly, "E R R O R"
                Exit Sub
            End If
            
            If .lbMaleNames.ListCount = 0 Then
                Session.MsgBox "You must have at least one male name", vbOKOnly, "E R R O R"
                Exit Sub
            End If
            
            If .lbFemaleNames.ListCount = 0 Then
                Session.MsgBox "You must have at least one female name", vbOKOnly, "E R R O R"
                Exit Sub
            End If
            
            If .lbPhyNames.ListCount = 0 Then
                Session.MsgBox "You must have at least one physician name", vbOKOnly, "E R R O R"
                Exit Sub
            End If
            
            Dim path As String
            Dim objFolders As Object
            Set objFolders = CreateObject("WScript.Shell").SpecialFolders
            path = objFolders("mydocuments")
            path = path + "\SpecGen.ini"
            Set objFolders = Nothing
            .cbClinInfo.Tag = .tbClinInfoPercent.Text
            .cbCollectTime.Tag = .tbCollectTimePercent.Text
            .cbControlNum.Tag = .tbControlNumPercent.Text
            .cbCourtesy.Tag = .tbCourtesyPercent.Text
            .cbDOB.Tag = .tbDOBPercent.Text
            '.cbEmail.Tag = .tbEmailPercent.Text ' TNG
            .cbNPI.Tag = .tbNPIPercent.Text
            .cbPatID.Tag = .tbPatIDPercent.Text
            .cbPhyID.Tag = .tbPhyIDPercent.Text
            .cbPhyName.Tag = .tbPhyNamePercent.Text
            .cbRespParty.Tag = .tbRespPartyPercent.Text
            .cbSSN.Tag = .tbSSNPercent.Text
            .cbTubes.Tag = .tbTubesPercent.Text
            .cbVolume.Tag = .tbVolumePercent.Text
            .cbHeight.Tag = .tbHeightPercent.Text
            .cbWeight.Tag = .tbWeightPercent.Text
        
            Buffer = "1"
            Open path For Output Shared As #1
            Buffer = Buffer + Format(.lbBlue.caption, "000")
            Buffer = Buffer + Format(.lbGreen.caption, "000")
            Buffer = Buffer + Format(.lbRed.caption, "000")
            
            If .cbDOB.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbDOB.Tag, "000")
            End If
            
            ' TNG
            'If .cbEmail.Tag = "" Then
            '    Buffer = Buffer + "000"
            'Else
            '    Buffer = Buffer + Format(.cbEmail.Tag, "000")
            'End If
            
            If .cbCollectTime.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbCollectTime.Tag, "000")
            End If
            
            If .cbNPI.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbNPI.Tag, "000")
            End If
            
            If .cbPhyID.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbPhyID.Tag, "000")
            End If
            
            If .cbSSN.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbSSN.Tag, "000")
            End If
            
            If .cbPatID.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbPatID.Tag, "000")
            End If
            
            If .cbVolume.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbVolume.Tag, "000")
            End If
            
            If .cbPhyName.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbPhyName.Tag, "000")
            End If
            
            If .cbClinInfo.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbClinInfo.Tag, "000")
            End If
            
            If .cbTubes.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbTubes.Tag, "000")
            End If
            
            If .cbRespParty.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbRespParty.Tag, "000")
            End If
            
            If .cbCourtesy.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbCourtesy.Tag, "000")
            End If
            
            If .cbControlNum.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbControlNum.Tag, "000")
            End If
            
            If .cbHeight.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbHeight.Tag, "000")
            End If
            
            If .cbWeight.Tag = "" Then
                Buffer = Buffer + "000"
            Else
                Buffer = Buffer + Format(.cbWeight.Tag, "000")
            End If
            Print #1, Buffer
            
            For Looper = 0 To .lbTestNum.ListCount - 1
                Print #1, "T" & .lbTestNum.List(Looper)
            Next Looper
            
            For Looper = 0 To .lbLastNames.ListCount - 1
                Print #1, "L" & .lbLastNames.List(Looper)
            Next Looper
            
            For Looper = 0 To .lbMaleNames.ListCount - 1
                Print #1, "M" & .lbMaleNames.List(Looper)
            Next Looper
            
            For Looper = 0 To .lbFemaleNames.ListCount - 1
                Print #1, "F" & .lbFemaleNames.List(Looper)
            Next Looper
            
            For Looper = 0 To .lbPhyNames.ListCount - 1
                Print #1, "P" & .lbPhyNames.List(Looper)
            Next Looper
            
            For Looper = 0 To .lbZipCode.ListCount - 1
                Print #1, "Z" & .lbZipCode.List(Looper)
            Next Looper
            
            If .tbDOBValue.Value > "" Then
                Print #1, "D1" & .tbDOBValue.Value
            End If
            
            'TNG
            'If .tbEmailValue.Value > "" Then
            '    Print #1, "D1" & .tbEmailValue.Value
            'End If
            
            If .tbTimeValue.Value > "" Then
                Print #1, "D2" & .tbTimeValue.Value
            End If
            
            If .tbNPIValue.Value > "" Then
                Print #1, "D3" & .tbNPIValue.Value
            End If
            
            If .tbPhyIdValue.Value > "" Then
                Print #1, "D4" & .tbPhyIdValue.Value
            End If
            
            If .tbSSNValue.Value > "" Then
                Print #1, "D5" & .tbSSNValue.Value
            End If
            
            If .tbPatIdValue.Value > "" Then
                Print #1, "D6" & .tbPatIdValue.Value
            End If
            
            If .tbVolumeValue.Value > "" Then
                Print #1, "D7" & .tbVolumeValue.Value
            End If
            
            If .tbHeightValue.Value > "" Then
                Print #1, "D8" & .tbHeightValue.Value
            End If
            
            If .tbPhyNameValue.Value > "" Then
                Print #1, "D9" & .tbPhyNameValue.Value
            End If
            
            If .tbTubeValue.Value > "" Then
                Print #1, "DA" & .tbTubeValue.Value
            End If
            
            If .tbControlNumValue.Value > "" Then
                Print #1, "DB" & .tbControlNumValue.Value
            End If
            
            If .tbWeightValue.Value > "" Then
                Print #1, "DC" & .tbWeightValue.Value
            End If
            
            Close #1
            .imgLogo.Visible = True
            Call bnCancel_Click
        
        Case "L"
            If (Len(.tbAccession.Text) < 3) Then
                Session.MsgBox "Invalid accession number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbJul.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
            
            If Not IsNumeric(.tbSequence.Text) Then
                Session.MsgBox "Invalid sequence number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbAccession.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                                                    
            If Not IsNumeric(.tbBrother.Text) Then
                Session.MsgBox "Invalid brother number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbSequence.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                
            If (Not IsNumeric(.tbEndSeq.Text)) Or (.tbEndSeq.Text < .tbSequence.Text) Then
                Session.MsgBox "Invalid ending sequence", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbBrother.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            Else
                .tbNumSpec = Format((val(.tbEndSeq.Text) - val(.tbSequence.Text) + 1), "0000")
            End If
            
            .frameAccession.Visible = False
            .cbOverwrite.Visible = False
            .Hide
            
            Call ResultSpecimens
            
        Case "D"
            Session.MsgBox "Function not implemented", vbOKOnly, "E R R O R"
                
        Case "R"
            If (Len(.tbAccession.Text) < 3) Then
                Session.MsgBox "Invalid accession number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbJul.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
            
            If Not IsNumeric(.tbSequence.Text) Then
                Session.MsgBox "Invalid sequence number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbAccession.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                                                    
            If Not IsNumeric(.tbBrother.Text) Then
                Session.MsgBox "Invalid brother number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbSequence.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                
            If (Not IsNumeric(.tbEndSeq.Text)) Or (.tbEndSeq.Text < .tbSequence.Text) Then
                Session.MsgBox "Invalid ending sequence", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbBrother.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
            
            .frameAccession.Visible = False
            .Hide
            
            Call ResetSpecimens
            
        Case "S"
            If (Len(.tbAccount.Text) < 8) Or (Not IsNumeric(.tbAccount.Text)) Then
                Session.MsgBox "Invalid account number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbForm.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                        
            If (Len(.tbAccession.Text) < 3) Then
                Session.MsgBox "Invalid accession number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbJul.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
            
            If Not IsNumeric(.tbSequence.Text) Then
                Session.MsgBox "Invalid sequence number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbAccession.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                                                    
            If Not IsNumeric(.tbBrother.Text) Then
                Session.MsgBox "Invalid brother number", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbSequence.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
                
            If (Not IsNumeric(.tbEndSeq.Text)) Or (.tbEndSeq.Text < .tbSequence.Text) Then
                Session.MsgBox "Invalid ending sequence", vbOKOnly, "E R R O R"
                .frameAccession.SetFocus
                .tbBrother.SetFocus
                SendKeys "{TAB}"
                Exit Sub
            End If
            
            If .opBillCodeSpecify.Value = True Then
                If Len(.tbBillCode.Text) < 2 Or InStr(1, ValidBill, .tbBillCode.Text, vbTextCompare) = 0 Then
                    Session.MsgBox "Invalid bill code", vbOKOnly, "E R R O R"
                    .frameControls.SetFocus
                    .frameBillCode.SetFocus
                    .opBillCodeSpecify.SetFocus
                    SendKeys "{TAB}"
                    Exit Sub
                End If
            End If
            
            .frameAccession.Visible = False
            .frameControls.Visible = False
            .framePercents.Visible = False
            .Hide
            Call GenerateSpecimens
    End Select

End With
End Sub

Private Sub bnResetSpecimens_Click()
With FormSpecGen
    .Tag = "R"
    .caption = "Specimen Generator Version " + VERSION + " -- Reset Specimens"
    
    .bnOk.caption = "Start"
    .bnOk.Accelerator = "S"
    .bnOk.Visible = True

    .bnCancel.caption = "Cancel"
    .bnCancel.Accelerator = "C"
    
    .frameOption.Visible = False
    .frameAccession.Top = 198
    .frameAccession.Left = 340
    .frameAccession.Visible = True
    
    .tbJul.SetFocus
    
End With
End Sub

Private Sub bnResultSpecimens_Click()
With FormSpecGen
    .Tag = "L"
    .caption = "Specimen Generator Version " + VERSION + " -- Result Specimens"
    
    .bnOk.caption = "Start"
    .bnOk.Accelerator = "S"
    .bnOk.Visible = True

    .bnCancel.caption = "Cancel"
    .bnCancel.Accelerator = "C"
    
    .frameOption.Visible = False
    .frameAccession.Top = 198
    .frameAccession.Left = 340
    
    .cbOverwrite.Visible = True
    .frameAccession.Visible = True
    .tbJul.SetFocus

End With
End Sub

Private Sub bnSetDefaults_Click()
With FormSpecGen
    .Tag = "O"
    .caption = "Specimen Generator Version " + VERSION + " -- Choose Default Values"
    
    .bnOk.caption = "Apply"
    .bnOk.Accelerator = "A"
    .bnCancel.caption = "Cancel"
    .bnCancel.Accelerator = "C"
    
    .bnOk.Visible = True
    
    .frameColor.lbBlue.caption = Str(.spbBlueColor.Value)
    .frameColor.lbGreen.caption = Str(.spbGreenColor.Value)
    .frameColor.lbRed.caption = Str(.spbRedColor.Value)
        
    .frameOption.Visible = False
    
    .imgLogo.Visible = False
    
    .framePercents.Visible = True
    .frameColor.Visible = True
    .frameNames.Visible = True
    .frameNames.SetFocus
    
End With
End Sub

Private Sub bnSingleSpec_Click()
With FormSpecGen
    .Tag = "S"
    .caption = "Specimen Generator Version " + VERSION + " -- Single Specimen processing"
    
    .frameOption.Visible = False
    
    .bnOk.caption = "Generate"
    .bnOk.Accelerator = "G"
    .bnOk.Visible = True
    
    .bnCancel.caption = "Cancel"
    .bnCancel.Accelerator = "C"
    
    .frameAccession.Top = 30
    .frameAccession.Left = 204
    
    .frameAccession.Visible = True
    .frameControls.Visible = True
    .framePercents.Visible = True
    .frameOEInfo.Visible = True
    
    .tbForm.SetFocus
    
End With
End Sub

Private Sub cbDOB_Click()
With FormSpecGen
    If .cbDOB.Value = True Then
        .tbDOBPercent.BackColor = &H80000005
        .tbDOBPercent.Enabled = True
        If .cbDOB.Tag = "" Then
            .tbDOBPercent.Value = 100
        Else
            .tbDOBPercent.Value = val(.cbDOB.Tag)
        End If
        .tbDOBValue = .tbDOBValue.Tag
    Else
        .tbDOBPercent.BackColor = &H8000000B
        .tbDOBPercent.Enabled = False
        .cbDOB.Tag = .tbDOBPercent.Value
        .tbDOBPercent.Value = 0
        .tbDOBValue.Tag = .tbDOBValue
        .tbDOBValue.Value = ""
    End If
    .tbDOBValue.BackColor = .tbDOBPercent.BackColor
    .tbDOBValue.Enabled = .tbDOBPercent.Enabled

    If .Tag <> "" And .tbDOBPercent.Enabled Then
        .tbDOBPercent.SetFocus
    End If

End With
End Sub
Private Sub cbEmail_Click() ' TNG
With FormSpecGen
    If .cbEmail.Value = True Then
        .tbEmailPercent.BackColor = &H80000005
        .tbEmailPercent.Enabled = True
        If .cbEmail.Tag = "" Then
            .tbEmailPercent.Value = 100
        Else
            .tbEmailPercent.Value = val(.cbEmail.Tag)
        End If
        .tbEmailValue = .tbEmailValue.Tag
    Else
        .tbEmailPercent.BackColor = &H8000000B
        .tbEmailPercent.Enabled = False
        .cbEmail.Tag = .tbEmailPercent.Value
        .tbEmailPercent.Value = 0
        .tbEmailValue.Tag = .tbEmailValue
        .tbEmailValue.Value = ""
    End If
    .tbEmailValue.BackColor = .tbEmailPercent.BackColor
    .tbEmailValue.Enabled = .tbEmailPercent.Enabled

    If .Tag <> "" And .tbEmailPercent.Enabled Then
        .tbEmailPercent.SetFocus
    End If

End With
End Sub
Private Sub cbCollectTime_Click()
With FormSpecGen
    If .cbCollectTime.Value = True Then
        .tbCollectTimePercent.BackColor = &H80000005
        .tbCollectTimePercent.Enabled = True
        If .cbCollectTime.Tag = "" Then
            .tbCollectTimePercent.Value = 100
        Else
            .tbCollectTimePercent.Value = val(.cbCollectTime.Tag)
        End If
        .tbTimeValue = .tbTimeValue.Tag
    Else
        .tbCollectTimePercent.BackColor = &H8000000B
        .tbCollectTimePercent.Enabled = False
        .cbCollectTime.Tag = .tbCollectTimePercent.Value
        .tbCollectTimePercent.Value = 0
        .tbTimeValue.Tag = .tbTimeValue
        .tbTimeValue = ""
    End If
    
    .tbTimeValue.BackColor = .tbCollectTimePercent.BackColor
    .tbTimeValue.Enabled = .tbCollectTimePercent.Enabled
        
    If .Tag <> "" And .tbCollectTimePercent.Enabled Then
        .tbCollectTimePercent.SetFocus
    End If
    
End With
End Sub



Private Sub cbNPI_Click()
With FormSpecGen
    If .cbNPI.Value = True Then
        .tbNPIPercent.BackColor = &H80000005
        .tbNPIPercent.Enabled = True
        If .cbNPI.Tag = "" Then
            .tbNPIPercent.Value = 100
        Else
            .tbNPIPercent.Value = val(.cbNPI.Tag)
        End If
        .tbNPIValue.Value = .tbNPIValue.Tag
    Else
        .tbNPIPercent.BackColor = &H8000000B
        .tbNPIPercent.Enabled = False
        .cbNPI.Tag = tbNPIPercent.Value
        .tbNPIPercent.Value = 0
        .tbNPIValue.Tag = tbNPIValue.Value
        .tbNPIValue.Value = ""
    End If
    .tbNPIValue.BackColor = .tbNPIPercent.BackColor
    .tbNPIValue.Enabled = .tbNPIPercent.Enabled
    
    If .Tag <> "" And .tbNPIPercent.Enabled Then
        .tbNPIPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbPhyID_Click()
With FormSpecGen
    If .cbPhyID.Value = True Then
        .tbPhyIDPercent.BackColor = &H80000005
        .tbPhyIDPercent.Enabled = True
        If .cbPhyID.Tag = "" Then
            .tbPhyIDPercent.Value = 100
        Else
            .tbPhyIDPercent.Value = val(.cbPhyID.Tag)
        End If
        .tbPhyIdValue.Value = .tbPhyIdValue.Tag
    Else
        .tbPhyIDPercent.BackColor = &H8000000B
        .tbPhyIDPercent.Enabled = False
        .cbPhyID.Tag = .tbPhyIDPercent.Value
        .tbPhyIDPercent.Value = 0
        .tbPhyIdValue.Tag = .tbPhyIdValue.Value
        .tbPhyIdValue.Value = ""
    End If
    
    .tbPhyIdValue.BackColor = .tbPatIDPercent.BackColor
    .tbPhyIdValue.Enabled = .tbPhyIDPercent.Enabled
    
    If .Tag <> "" And .tbPhyIDPercent.Enabled Then
        .tbPhyIDPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbSSN_Click()
With FormSpecGen
    If .cbSSN.Value = True Then
        .tbSSNPercent.BackColor = &H80000005
        .tbSSNPercent.Enabled = True
        If .cbSSN.Tag = "" Then
            .tbSSNPercent.Value = 100
        Else
            .tbSSNPercent.Value = val(.cbSSN.Tag)
        End If
        .tbSSNValue.Value = .tbSSNValue.Tag
    Else
        .tbSSNPercent.BackColor = &H8000000B
        .tbSSNPercent.Enabled = False
        .cbSSN.Tag = .tbSSNPercent.Value
        .tbSSNPercent.Value = 0
        .tbSSNValue.Tag = .tbSSNValue.Value
        .tbSSNValue.Value = ""
    End If
    
    .tbSSNValue.BackColor = .tbSSNPercent.BackColor
    .tbSSNValue.Enabled = .tbSSNPercent.Enabled

    If .Tag <> "" And .tbSSNPercent.Enabled Then
        .tbSSNPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbPatID_Click()
With FormSpecGen
    If .cbPatID.Value = True Then
        .tbPatIDPercent.BackColor = &H80000005
        .tbPatIDPercent.Enabled = True
        If .cbPatID.Tag = "" Then
            .tbPatIDPercent.Value = 100
        Else
            .tbPatIDPercent.Value = val(.cbPatID.Tag)
        End If
        .tbPatIdValue.Value = .tbPatIdValue.Tag
    Else
        .tbPatIDPercent.BackColor = &H8000000B
        .tbPatIDPercent.Enabled = False
        .cbPatID.Tag = .tbPatIDPercent.Value
        .tbPatIDPercent.Value = 0
        .tbPatIdValue.Tag = .tbPatIdValue.Value
        .tbPatIdValue.Value = ""
    End If
    
    .tbPatIdValue.BackColor = tbPatIDPercent.BackColor
    .tbPatIdValue.Enabled = tbPatIDPercent.Enabled
    
    If .Tag <> "" And .tbPatIDPercent.Enabled Then
        .tbPatIDPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbVolume_Click()
With FormSpecGen
    If .cbVolume.Value = True Then
        .tbVolumePercent.BackColor = &H80000005
        .tbVolumePercent.Enabled = True
        If .cbVolume.Tag = "" Then
            .tbVolumePercent.Value = 100
        Else
            .tbVolumePercent.Value = val(.cbVolume.Tag)
        End If
        .tbVolumeValue.Value = .tbVolumeValue.Tag
    Else
        .tbVolumePercent.BackColor = &H8000000B
        .tbVolumePercent.Enabled = False
        .cbVolume.Tag = .tbVolumePercent.Value
        .tbVolumePercent.Value = 0
        .tbVolumeValue.Tag = tbVolumeValue.Value
        .tbVolumeValue.Value = ""
    End If
    
    .tbVolumeValue.BackColor = tbVolumePercent.BackColor
    .tbVolumeValue.Enabled = tbVolumePercent.Enabled
    
    If .Tag <> "" And .tbVolumePercent.Enabled Then
        .tbVolumePercent.SetFocus
    End If
        
End With
End Sub

Private Sub cbHeight_Click()
With FormSpecGen
    If .cbHeight.Value = True Then
        .tbHeightPercent.BackColor = &H80000005
        .tbHeightPercent.Enabled = True
        If .cbHeight.Tag = "" Then
            .tbHeightPercent.Value = 100
        Else
            .tbHeightPercent.Value = val(.cbHeight.Tag)
        End If
        .tbHeightValue.Value = .tbHeightValue.Tag
    Else
        .tbHeightPercent.BackColor = &H8000000B
        .tbHeightPercent.Enabled = False
        .cbHeight.Tag = .tbHeightPercent.Value
        .tbHeightPercent.Value = 0
        .tbHeightValue.Tag = .tbHeightValue
        .tbHeightValue.Value = ""
    End If
    
    .tbHeightValue.BackColor = tbHeightPercent.BackColor
    .tbHeightValue.Enabled = tbHeightPercent.Enabled
    
    If .Tag <> "" And .tbHeightPercent.Enabled Then
        .tbHeightPercent.SetFocus
    End If
        
End With
End Sub

Private Sub cbPhyName_Click()
With FormSpecGen
    If .cbPhyName.Value = True Then
        .tbPhyNamePercent.BackColor = &H80000005
        .tbPhyNamePercent.Enabled = True
        If .cbPhyName.Tag = "" Then
            .tbPhyNamePercent.Value = 100
        Else
            .tbPhyNamePercent.Value = val(.cbPhyName.Tag)
        End If
        .tbPhyNameValue.Value = .tbPhyNameValue.Tag
    Else
        .tbPhyNamePercent.BackColor = &H8000000B
        .tbPhyNamePercent.Enabled = False
        .cbPhyName.Tag = .tbPhyNamePercent.Value
        .tbPhyNamePercent.Value = 0
        .tbPhyNameValue.Tag = .tbPhyNameValue.Value
        .tbPhyNameValue.Value = ""
    End If
    
    .tbPhyNameValue.BackColor = .tbPhyNamePercent.BackColor
    .tbPhyNameValue.Enabled = .tbPhyNamePercent.Enabled
    
    If .Tag <> "" And .tbPhyNamePercent.Enabled Then
        .tbPhyNamePercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbClinInfo_Click()
With FormSpecGen
    If .cbClinInfo.Value = True Then
        .tbClinInfoPercent.BackColor = &H80000005
        .tbClinInfoPercent.Enabled = True
        If .cbClinInfo.Tag = "" Then
            .tbClinInfoPercent.Value = 100
        Else
            .tbClinInfoPercent.Value = val(.cbClinInfo.Tag)
        End If
    Else
        .tbClinInfoPercent.BackColor = &H8000000B
        .tbClinInfoPercent.Enabled = False
        .cbClinInfo.Tag = .tbClinInfoPercent.Value
        .tbClinInfoPercent.Value = 0
    End If
    
    If .Tag <> "" And .tbClinInfoPercent.Enabled Then
        .tbClinInfoPercent.SetFocus
    End If
        
End With
End Sub

Private Sub cbTubes_Click()
With FormSpecGen
    If .cbTubes.Value = True Then
        .tbTubesPercent.BackColor = &H80000005
        .tbTubesPercent.Enabled = True
        If .cbTubes.Tag = "" Then
            .tbTubesPercent.Value = 100
        Else
            .tbTubesPercent.Value = val(.cbTubes.Tag)
        End If
        .tbTubeValue.Value = .tbTubeValue.Tag
    Else
        .tbTubesPercent.BackColor = &H8000000B
        .tbTubesPercent.Enabled = False
        .cbTubes.Tag = .tbTubesPercent.Value
        .tbTubesPercent.Value = 0
        .tbTubeValue.Tag = .tbTubeValue.Value
        .tbTubeValue.Value = ""
    End If
    
    .tbTubeValue.BackColor = .tbTubesPercent.BackColor
    .tbTubeValue.Enabled = .tbTubesPercent.Enabled
    
    If .Tag <> "" And .tbTubesPercent.Enabled Then
        .tbTubesPercent.SetFocus
    End If

End With
End Sub

Private Sub cbRespParty_Click()
With FormSpecGen
    If .cbRespParty.Value = True Then
        .tbRespPartyPercent.BackColor = &H80000005
        .tbRespPartyPercent.Enabled = True
        If .cbRespParty.Tag = "" Then
            .tbRespPartyPercent.Value = 100
        Else
            .tbRespPartyPercent.Value = val(.cbRespParty.Tag)
        End If
    Else
        .tbRespPartyPercent.BackColor = &H8000000B
        .tbRespPartyPercent.Enabled = False
        .cbRespParty.Tag = .tbRespPartyPercent.Value
        .tbRespPartyPercent.Value = 0
    End If
    
    If .Tag <> "" And .tbRespPartyPercent.Enabled Then
        .tbRespPartyPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbCourtesy_Click()
With FormSpecGen
    If .cbCourtesy.Value = True Then
        .tbCourtesyPercent.BackColor = &H80000005
        .tbCourtesyPercent.Enabled = True
        If .cbCourtesy.Tag = "" Then
            .tbCourtesyPercent.Value = 100
        Else
            .tbCourtesyPercent.Value = val(.cbCourtesy.Tag)
        End If
    Else
        .tbCourtesyPercent.BackColor = &H8000000B
        .tbCourtesyPercent.Enabled = False
        .cbCourtesy.Tag = .tbCourtesyPercent.Value
        .tbCourtesyPercent.Value = 0
    End If
    
    If .Tag <> "" And .tbCourtesyPercent.Enabled Then
        .tbCourtesyPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbControlNum_Click()
With FormSpecGen
    If .cbControlNum.Value = True Then
        .tbControlNumPercent.BackColor = &H80000005
        .tbControlNumPercent.Enabled = True
        If .cbControlNum.Tag = "" Then
            .tbControlNumPercent.Value = 100
        Else
            .tbControlNumPercent.Value = val(.cbControlNum.Tag)
        End If
        .tbControlNumValue.Value = .tbControlNumValue.Tag
    Else
        .tbControlNumPercent.BackColor = &H8000000B
        .tbControlNumPercent.Enabled = False
        .cbControlNum.Tag = .tbControlNumPercent.Value
        .tbControlNumPercent.Value = 0
        .tbControlNumValue.Tag = .tbControlNumValue.Value
        .tbControlNumValue.Value = ""
    End If
        
    
    .tbControlNumValue.BackColor = .tbControlNumPercent.BackColor
    .tbControlNumValue.Enabled = .tbControlNumPercent.Enabled
            
    If .Tag <> "" And .tbControlNumPercent.Enabled Then
        .tbControlNumPercent.SetFocus
    End If
    
End With
End Sub

Private Sub cbWeight_Click()
With FormSpecGen
    If .cbWeight.Value = True Then
        .tbWeightPercent.BackColor = &H80000005
        .tbWeightPercent.Enabled = True
        If .cbWeight.Tag = "" Then
            .tbWeightPercent.Value = 100
        Else
            .tbWeightPercent.Value = val(.cbWeight.Tag)
        End If
        .tbWeightValue.Value = .tbWeightValue.Tag
    Else
        .tbWeightPercent.BackColor = &H8000000B
        .tbWeightPercent.Enabled = False
        .cbWeight.Tag = .tbWeightPercent.Value
        .tbWeightPercent.Value = 0
        .tbWeightValue.Tag = .tbWeightValue
        .tbWeightValue.Value = ""
    End If
    
    .tbWeightValue.BackColor = tbWeightPercent.BackColor
    .tbWeightValue.Enabled = tbWeightPercent.Enabled
    
    If .Tag <> "" And .tbWeightPercent.Enabled Then
        .tbWeightPercent.SetFocus
    End If
        
End With
End Sub

Private Sub frmEmail_Click()

End Sub

Private Sub frameControls_Click()

End Sub

Private Sub frameFasting_Click()

End Sub

Private Sub frmZipCode_Click()

End Sub

Private Sub lbFemaleNames_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lbFemaleNames.ListIndex = -1
End Sub

Private Sub lbFemaleNames_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyDelete And .lbFemaleNames.ListIndex > -1 Then
        .lbFemaleNames.RemoveItem (.lbFemaleNames.ListIndex)
    End If
    If .lbFemaleNames.ListCount = 0 Then
        .bnOk.Visible = False
        .bnOk.Enabled = False
    End If
End With
End Sub

Private Sub lbLastNames_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lbLastNames.ListIndex = -1
End Sub

Private Sub lbLastNames_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyDelete And .lbLastNames.ListIndex > -1 Then
        .lbLastNames.RemoveItem (.lbLastNames.ListIndex)
    End If

    If .lbLastNames.ListCount = 0 Then
        .bnOk.Visible = False
        .bnOk.Enabled = False
    End If
End With
End Sub

Private Sub lbMaleNames_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lbMaleNames.ListIndex = -1
End Sub

Private Sub lbMaleNames_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyDelete And .lbMaleNames.ListIndex > -1 Then
        .lbMaleNames.RemoveItem (.lbMaleNames.ListIndex)
    End If
    
    If .lbMaleNames.ListCount = 0 Then
        .bnOk.Visible = False
        .bnOk.Enabled = False
    End If
End With
End Sub

Private Sub lbPhyNames_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lbPhyNames.ListIndex = -1
End Sub

Private Sub lbTestNum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lbTestNum.ListIndex = -1
End Sub

Private Sub lbTestNum_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyDelete And .lbTestNum.ListIndex > -1 Then
        .lbTestNum.RemoveItem (.lbTestNum.ListIndex)
    End If
    
    If .lbTestNum.ListCount = 0 Then
        .bnOk.Visible = False
        .bnOk.Enabled = False
    End If
End With
End Sub

Private Sub lbPhyNames_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyDelete And .lbPhyNames.ListIndex > -1 Then
        .lbPhyNames.RemoveItem (.lbPhyNames.ListIndex)
    End If
    
    If .lbPhyNames.ListCount = 0 Then
        .bnOk.Visible = False
        .bnOk.Enabled = False
    End If
End With
End Sub

Private Sub lbZipCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    FormSpecGen.lbZipCode.ListIndex = -1
End Sub

Private Sub lbZipCode_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyDelete And .lbZipCode.ListIndex > -1 Then
        .lbZipCode.RemoveItem (.lbZipCode.ListIndex)
    End If
End With
End Sub

Private Sub opBillCodeRandom_Click()
With FormSpecGen
    .tbBillCode.Enabled = False
    .tbBillCode.BackColor = &H8000000F
    .tbBillCode.Text = ""
End With

End Sub

Private Sub opBillCodeSpecify_Click()
With FormSpecGen
    .tbBillCode.Enabled = True
    .tbBillCode.BackColor = &H80000005
    .tbBillCode.SetFocus
End With
End Sub

Private Sub opGenderFemale_Change()
    If FormSpecGen.tbForm.Text = "120" Then
        opGenderFemale.Value = True
    End If
End Sub

Private Sub opLoopZip_Click()
With FormSpecGen
    .tbSpecifyZip.Enabled = False
    .tbSpecifyZip.Text = ""
    .tbSpecifyZip.BackColor = &H8000000F
End With
End Sub

Private Sub opRandomName_Click()
With FormSpecGen
    .tbSpecifyFirstName.Enabled = False
    .tbSpecifyLastName.Enabled = False
    .tbSpecifyMiddleName.Enabled = False
    
    .tbSpecifyFirstName.BackColor = &H8000000F
    .tbSpecifyLastName.BackColor = &H8000000F
    .tbSpecifyMiddleName.BackColor = &H8000000F
    
    .tbSpecifyFirstName.Text = ""
    .tbSpecifyLastName.Text = ""
    .tbSpecifyMiddleName.Text = ""
End With
End Sub

Private Sub opRandomTest_Click()
With FormSpecGen
    .tbSpecifyTest.Text = ""
    .tbSpecifyTest.Enabled = False
    .tbSpecifyTest.BackColor = &H8000000F
End With
End Sub

Private Sub opRandomZip_Click()
With FormSpecGen
    .tbSpecifyZip.Enabled = False
    .tbSpecifyZip.Text = ""
    .tbSpecifyZip.BackColor = &H8000000F
End With
End Sub

Private Sub opSpecifyName_Click()
With FormSpecGen
    .tbSpecifyFirstName.Enabled = True
    .tbSpecifyLastName.Enabled = True
    .tbSpecifyMiddleName.Enabled = True
    
    .tbSpecifyFirstName.BackColor = &H80000005
    .tbSpecifyLastName.BackColor = &H80000005
    .tbSpecifyMiddleName.BackColor = &H80000005
    .tbSpecifyFirstName.SetFocus
End With
End Sub

Private Sub opSpecifyTest_Click()
With FormSpecGen
    .tbSpecifyTest.Text = ""
    .tbSpecifyTest.Enabled = True
    .tbSpecifyTest.BackColor = &H80000005
    .tbSpecifyTest.SetFocus
End With
End Sub

Private Sub opSpecifyZip_Click()
With FormSpecGen
    .tbSpecifyZip.Text = ""
    .tbSpecifyZip.Enabled = True
    .tbSpecifyZip.BackColor = &H80000005
    .tbSpecifyZip.SetFocus
End With
End Sub

Private Sub opTestLoop_Click()
With FormSpecGen
    .tbSpecifyTest.Text = ""
    .tbSpecifyTest.Enabled = False
    .tbSpecifyTest.BackColor = &H8000000F
End With
End Sub

Private Sub spbBlueColor_Change()
With FormSpecGen
    If .spbBlueColor.Value < 0 Then .spbBlueColor = 255
    .bnBackGround.BackColor = RGB(.spbRedColor.Value, .spbGreenColor.Value, .spbBlueColor.Value)
    .lbBlue.caption = .spbBlueColor.Value
End With
End Sub

Private Sub spbBlueColor_SpinDown()
With FormSpecGen
    If .spbBlueColor.Value = 247 Then .spbBlueColor.Value = 248
End With
End Sub

Private Sub spbBlueColor_SpinUp()
With FormSpecGen
    If .spbBlueColor.Value > 260 Then .spbBlueColor.Value = 0
    If .spbBlueColor.Value > 255 Then .spbBlueColor.Value = 255
End With
End Sub

Private Sub spbGreenColor_Change()
With FormSpecGen
    If .spbGreenColor.Value < 0 Then .spbGreenColor = 255
    .bnBackGround.BackColor = RGB(.spbRedColor.Value, .spbGreenColor.Value, .spbBlueColor.Value)
    .lbGreen.caption = .spbGreenColor.Value
End With
End Sub

Private Sub spbGreenColor_SpinDown()
With FormSpecGen
    If .spbGreenColor.Value = 247 Then .spbGreenColor.Value = 248
End With
End Sub

Private Sub spbGreenColor_SpinUp()
With FormSpecGen
    If .spbGreenColor.Value > 260 Then .spbGreenColor.Value = 0
    If .spbGreenColor.Value > 255 Then .spbGreenColor.Value = 255
End With
End Sub

Private Sub spbRedColor_Change()
With FormSpecGen
    If .spbRedColor.Value < 0 Then .spbRedColor = 255
    .bnBackGround.BackColor = RGB(.spbRedColor.Value, .spbGreenColor.Value, .spbBlueColor.Value)
    .lbRed.caption = .spbRedColor.Value
End With
End Sub

Private Sub spbRedColor_SpinDown()
With FormSpecGen
    If .spbRedColor.Value = 247 Then .spbRedColor.Value = 248
End With
End Sub

Private Sub spbRedColor_SpinUp()
With FormSpecGen
    If .spbRedColor.Value > 260 Then .spbRedColor.Value = 0
    If .spbRedColor.Value > 255 Then .spbRedColor.Value = 255
End With
End Sub
Private Sub tbEmail_AfterUpdate()
    tbEmail.Text = UCase(tbEmail.Text)
End Sub
Private Sub tbBillCode_AfterUpdate()
    tbBillCode.Text = UCase(tbBillCode.Text)
End Sub

Private Sub tbControlNumPercent_AfterUpdate()
With FormSpecGen
    If .tbForm.Text = "CCU" Then
        .tbControlNumPercent.Value = 100
    End If
End With
End Sub

Private Sub tbEndSeq_AfterUpdate()
With FormSpecGen
    If Len(.tbEndSeq.Text) < 4 Then
        .tbEndSeq.Text = Format(.tbEndSeq.Text, "0000")
    End If
    
    If .tbEndSeq.Text >= .tbSequence.Text Then
        .tbNumSpec.Text = Format(val(.tbEndSeq.Text) - val(.tbSequence.Text) + 1, "0000")
    Else
        .tbEndSeq.Text = .tbSequence.Text
        .tbNumSpec.Text = "0001"
        SendKeys "+{TAB}+{TAB}"
    End If
End With
End Sub

Private Sub tbFemaleName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyReturn And .tbFemaleName.Text <> "" Then
        .lbFemaleNames.AddItem (.tbFemaleName.Text)
        .tbFemaleName.Text = ""
        SendKeys "+{TAB}"
        If (.bnOk.Visible = False) Then
            Call CheckNames
        End If
    End If
    
End With
End Sub

Private Sub tbForm_AfterUpdate()
With FormSpecGen
    .tbForm.Text = UCase(.tbForm.Text)
    If .tbForm.Text = "120" Then
        .opGenderFemale.Value = True
    End If
End With
End Sub

Private Sub tbMaleName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyReturn And .tbMaleName.Text <> "" Then
        .lbMaleNames.AddItem (.tbMaleName.Text)
        .tbMaleName.Text = ""
        SendKeys "+{TAB}"
        
        If (.bnOk.Visible = False) Then
            Call CheckNames
        End If
    End If
End With
End Sub

Private Sub tbLastName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyReturn And .tbLastName.Text <> "" Then
        .lbLastNames.AddItem (.tbLastName.Text)
        .tbLastName.Text = ""
        SendKeys "+{TAB}"
        If (.bnOk.Visible = False) Then
            Call CheckNames
        End If
    End If
End With
End Sub

Private Sub tbNumSpec_AfterUpdate()
With FormSpecGen
    If Len(tbNumSpec.Text) < 4 Then
        tbNumSpec.Text = Format(val(tbNumSpec.Text), "0000")
    End If
    
    If val(.tbNumSpec.Text) + val(.tbSequence.Text) > 9999 Then
        .tbNumSpec.Text = Format(9999 - val(.tbSequence.Text))
    End If

    If val(.tbNumSpec.Text) > 0 Then
        .tbEndSeq.Text = Format(val(.tbSequence.Text) + val(.tbNumSpec.Text) - 1, "0000")
    Else
        .tbEndSeq.Text = .tbSequence.Text
    End If
End With
End Sub

Private Sub tbPhyName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyReturn And .tbPhyName.Text <> "" Then
        .lbPhyNames.AddItem (.tbPhyName.Text)
        .tbPhyName.Text = ""
        SendKeys "+{TAB}"
        If (.bnOk.Visible = False) Then
            Call CheckNames
        End If
    End If
End With
End Sub

Private Sub tbSequence_AfterUpdate()
With FormSpecGen
    If Len(tbSequence.Text) < 4 Then
        tbSequence.Text = Format(tbSequence.Text, "0000")
    End If

    If val(.tbEndSeq.Text) < val(.tbSequence.Text) Then
        .tbEndSeq.Text = .tbSequence.Text
        .tbNumSpec.Text = "0001"
    Else
        .tbNumSpec.Text = Format((val(.tbEndSeq.Text) - val(.tbSequence.Text) + 1), "0000")
    End If
End With
End Sub

Private Sub tbSpecifyLastName_Change()

End Sub

Private Sub tbTestNum_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyReturn Then
        If Len(tbTestNum.Text) = 6 Then
            .lbTestNum.AddItem (.tbTestNum.Text)
            .tbTestNum.Text = ""
        End If
        SendKeys "+{TAB}"
        
        If (.bnOk.Visible = False) Then
            Call CheckNames
        End If
    End If
End With
End Sub

Private Sub tbZipCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
With FormSpecGen
    If KeyCode = vbKeyReturn Then
        If Len(tbZipCode.Text) = 5 Then
            .lbZipCode.AddItem (.tbZipCode.Text)
            .tbZipCode.Text = ""
        End If
        SendKeys "+{TAB}"
    End If
End With
End Sub

Private Sub UserForm_Activate()
With FormSpecGen
    .bnCancel.caption = "Exit"
    .bnCancel.Accelerator = "X"
    .caption = "Specimen Generator Version " + VERSION

    '.cbEmail.Value = val(.cbEmail.Tag) > 0 'TNG
    .cbDOB.Value = val(.cbDOB.Tag) > 0
    .cbCollectTime.Value = val(.cbCollectTime.Tag) > 0
    .cbNPI.Value = val(.cbNPI.Tag) > 0
    .cbPhyID.Value = val(.cbPhyID.Tag) > 0
    .cbSSN.Value = val(.cbSSN.Tag) > 0
    .cbPatID.Value = val(.cbPatID.Tag) > 0
    .cbVolume.Value = val(.cbVolume.Tag) > 0
    .cbHeight.Value = val(.cbHeight.Tag) > 0
    .cbPhyName.Value = val(.cbPhyName.Tag) > 0
    .cbClinInfo.Value = val(.cbClinInfo.Tag) > 0
    .cbTubes.Value = val(.cbTubes.Tag) > 0
    .cbRespParty.Value = val(.cbRespParty.Tag) > 0
    .cbCourtesy.Value = val(.cbCourtesy.Tag) > 0
    .cbControlNum.Value = val(.cbControlNum.Tag) > 0
    .cbWeight.Value = val(.cbWeight.Tag) > 0
        
    .ScrollBars = fmScrollBarsBoth
    .KeepScrollBarsVisible = fmScrollBarsNone
    .ScrollHeight = .Height / 2
    .ScrollWidth = .Width / 2
    .ScrollLeft = 0
    .ScrollTop = 0
        
    .bnCancel.SetFocus
    Call setcolors

End With
End Sub

Sub setcolors()
With FormSpecGen
    .BackColor = RGB(.spbRedColor.Value, .spbGreenColor.Value, .spbBlueColor.Value)
    
    .frameAccession.BackColor = .BackColor
    .frameAccession.BorderColor = .BackColor
    
    .frameOEInfo.BackColor = .BackColor
    .frameOEInfo.BorderColor = .BackColor
    
    .frameOption.BackColor = .BackColor
    .frameOption.BorderColor = .BackColor
    
    .framePercents.BackColor = .BackColor
    .framePercents.BorderColor = .BackColor
    
    .frameColor.BackColor = .BackColor
    .frameColor.BorderColor = .BackColor
    
    .frameNames.BackColor = .BackColor
    .frameNames.BorderColor = .BackColor
    
    .frameControls.BackColor = .BackColor
    .frameControls.BorderColor = .BackColor
        
    .frameOverideName.BackColor = .BackColor
    .frameOverideTest.BackColor = .BackColor
    
    .frameGender.BackColor = .BackColor
    .frameFasting.BackColor = .BackColor
    .frameBillCode.BackColor = .BackColor
    .frmZipCode.BackColor = .BackColor
    '.frmEmailCode.BackColor = .BackColor   'TNG
    
    .bnBackGround.BackColor = .BackColor
End With
End Sub

Private Sub CheckNames()
With FormSpecGen
    If (.lbFemaleNames.ListCount > 0) And (.lbMaleNames.ListCount > 0) And (.lbLastNames.ListCount > 0) And (.lbPhyNames.ListCount > 0) And (.lbTestNum.ListCount > 0) Then
        .bnOk.Visible = True
        .bnOk.Enabled = True
    End If
End With
End Sub
