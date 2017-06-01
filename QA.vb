'-----------------------------------------------------------------------------------------------------------------------
' An Excel Macro to read and interpret narcotic dispensing data from Omnicell PCR reports, and compare it to AIMS anesthesia records.
' Both programs write out data in a spreadsheet that I would conservatively describe as "whack," so I built a workbook to read them
' for me. Omnicell report goes in a sheet named "Ocra," AIMS report goes in a sheet named "Aims," Button and grid go on a "Solutions"
' sheet, incomplete PCRs and issues on the "Solutions" sheet are written to the "Review" sheet, and Metadata gets written to a "Stats" 
' sheet. I'd just upload the file, but there's patient data in there and that's an actual crime.
'
'
' Kind of a niche project TBH 
'-----------------------------------------------------------------------------------------------------------------------
' I assume somebody else will have to look at this code eventually (sorry, by the way, if it's you) so here's how this
' works:
'   1) Do/Loop to run through every line of the OCRA sheet, looking for transactions coded "W" for a waste transaction
'   2) Takes Patient data, drug name, drug dose, and some other stuff from that line. Drug dose is totalled to account for
'      multiple vial sizes (i.e. user takes out a 1000mcg fent and a 100mcg fent to give 1100mcg)
'   3) It writes that data to the 'Solutions' sheet (because a: I don't even know if VBA can do dictionaries like python, and
'      b: excel literally is a data matrix software. Why fight it, you know?)
'   4) Nested Do/Loops. This is where it gets kind of ugly - Loop through the solutions sheet, reading in MRN. For each line,
'      Loop through the entire AIMS sheet looking for a match. When found, check drug name to exclude non-narcotics. Write
'       matches to the sheet
'   5) The easy part: Compare data within the 'Solutions' sheet, write desired results.
'   6) Metadata: analysis of data is pulled together as that data is handled. For example, PCR completeness checks are run as
'      the OCRA sheet is read the first time, but matches and percentages are pulled after AIMS has been read


Private Sub CommandButton1_Click()
'clear sheet
    
    With Sheets("Solutions")
        .Rows("4:" & .Rows.Count).Value = ""
    End With
    
    With Sheets("Review")
        .Rows("3:" & .Rows.Count).Value = ""
    End With
    
    ThisWorkbook.Sheets("Stats").Range("B2:B10").Value = ""
    ThisWorkbook.Sheets("Stats").Range("E2:E10").Value = ""
    
    
'Drug Admin Variables
    Dim strDrugNameOCRA As String
    Dim strDrugNameAIMS As String
    
    Dim OCRAFent As Single
    OCRAFent = 0
    Dim OCRAHydrom As Single
    OCRAHydrom = 0
    Dim OCRAMidaz As Single
    OCRAMidaz = 0
    Dim OCRAKetamine As Single
    OCRAKetamine = 0
    Dim OCRAMorphine As Single
    OCRAMorphine = 0
    Dim OCRARemi As Single
    OCRARemi = 0
    Dim OCRAMethohex As Single
    OCRAMethohex = 0

    Dim AIMSFent As Single
    AIMSFent = 0
    Dim AIMSHydrom As Single
    AIMSHydrom = 0
    Dim AIMSMidaz As Single
    AIMSMidaz = 0
    Dim AIMSKetamine As Single
    AIMSKetamine = 0
    Dim AIMSMorphine As Single
    AIMSMorphine = 0
    Dim AIMSRemi As Single
    AIMSRemi = 0
    Dim AIMSMethohex As Single
    AIMSMethohex = 0
    
'PCR check variables
    Dim PCRTotal As Long
    Dim PCRClosed As Long
    Dim PCRPercent As Single
    Dim AIMSTotalChecked As Integer
    Dim AIMSMissing As Integer
    Dim AIMSNoMatch As Integer
    Dim AIMSPercent As Single
    
'Incrementation stuff
    Dim intCount As Integer
    Dim intSubCount As Integer
    Dim intOutput As Long
    Dim strCurrentOCRAMRN As String
    Dim strCurrentOCRAName As String
    strCurrentOCRAName = "false"
    strCurrentOCRAMRN = "false"
    Dim strCurrentAIMSMRN As String
    Dim strCurrentAIMSName As String
    strCurrentAIMSName = "false"
    strCurrentAIMSMRN = "false"
    Dim intReviewOutput As Integer
    intReviewOutput = 2     'Will start writing to 3, but need to check previous rows whenever we right for duplicate data

'Patient Data
    Dim strOCRAMRN As String
    Dim strOCRAMRNNoZero As String
    Dim strAIMSMRN As String

'Misc Variables
    Dim strOCRADate As String
    
    
    PCRTotal = 0
    PCRClosed = 0
    AIMSTotalChecked = 0
    AIMSMissing = 0
    AIMSNoMatch = 0
    intCount = 2        'start reading OCRA on the first data row, skipping labels
    intOutput = 4       'start writing on the first empty row, skipping button/labels
    
    'Read PCR report (OCRA)
    Do While (ThisWorkbook.Sheets("OCRA").Range("A" & (intCount)).Value) <> ""
        strOCRAMRN = ThisWorkbook.Sheets("OCRA").Range("B" & intCount).Value
        strOCRAMRNNoZero = ThisWorkbook.Sheets("OCRA").Range("B" & intCount).Value
        strOCRANAME = ThisWorkbook.Sheets("OCRA").Range("D" & intCount).Value
        'PCR closure stuff for stats and review sheets
        PCRTotal = PCRTotal + 1
        If ThisWorkbook.Sheets("OCRA").Range("R" & intCount).Value = 0 Then
            PCRClosed = PCRClosed + 1
        Else
            'peel off leading zeroes
            Do While Left(strOCRAMRNNoZero, 1) = "0"
                strOCRAMRNNoZero = Mid(strOCRAMRNNoZero, 2)
            Loop
            
            If ThisWorkbook.Sheets("Review").Range("B" & intReviewOutput).Value <> strOCRAMRNNoZero Then
                intReviewOutput = intReviewOutput + 1
            End If
            ThisWorkbook.Sheets("Review").Range("B" & intReviewOutput).Value = ThisWorkbook.Sheets("OCRA").Range("B" & intCount).Value
            ThisWorkbook.Sheets("Review").Range("A" & intReviewOutput).Value = ThisWorkbook.Sheets("OCRA").Range("D" & intCount).Value
            ThisWorkbook.Sheets("Review").Range("C" & intReviewOutput).Value = ThisWorkbook.Sheets("OCRA").Range("AD" & intCount).Value
            strOCRADate = Left(ThisWorkbook.Sheets("OCRA").Range("A" & intCount).Value, 8)
            ThisWorkbook.Sheets("Review").Range("D" & intReviewOutput).Value = Mid(strOCRADate, 5, 2) & "/" & Right(strOCRADate, 2) & "/" & Left(strOCRADate, 4)
            ThisWorkbook.Sheets("Review").Range("E" & intReviewOutput).Value = "PCR Not Closed"
        End If
        'Pull foundational data from OCRA, write to form
        If ThisWorkbook.Sheets("OCRA").Range("AN" & intCount).Value = "W" And Left(ThisWorkbook.Sheets("OCRA").Range("AE" & intCount).Value, 7) <> "BIW_PRO" Then

        
                If strCurrentOCRAMRN = "false" Then
                    strCurrentOCRAMRN = strOCRAMRN
                    strCurrentOCRAName = strOCRANAME
                End If
            
                If strCurrentOCRAMRN = strOCRAMRN Then
                    strDrugNameOCRA = ThisWorkbook.Sheets("OCRA").Range("G" & intCount).Value
                    Select Case strDrugNameOCRA
                        Case "Fentanyl"
                            OCRAFent = OCRAFent + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "HYDROmorphone PF Dilaudid)"
                            OCRAHydrom = OCRAHydrom + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "Midazolam"
                            OCRAMidaz = OCRAMidaz + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "Ketamine"
                            OCRAKetamine = OCRAKetamine + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "Methohexital"
                            OCRAMethohex = OCRAMethohex + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "Remifentanil"
                            OCRARemi = OCRARemi + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "Morphine Sulfate"
                            OCRAMorphine = OCRAMorphine + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                        Case "Morphine PF"
                            OCRAMorphine = OCRAMorphine + ThisWorkbook.Sheets("OCRA").Range("AK" & intCount).Value
                    End Select
                    
                    Range("A" & intOutput).Value = strCurrentOCRAName
                    Range("B" & intOutput).Value = strCurrentOCRAMRN
                    Range("S" & intOutput).Value = ThisWorkbook.Sheets("OCRA").Range("AD" & intCount).Value
    
                    If OCRAFent <> 0 Then
                        Range("D" & intOutput).Value = OCRAFent
                    End If
                    If OCRAHydrom <> 0 Then
                        Range("F" & intOutput).Value = Left(str(OCRAHydrom), 5)
                    End If
                    If OCRAMidaz <> 0 Then
                        Range("H" & intOutput).Value = OCRAMidaz
                    End If
                    If OCRAKetamine <> 0 Then
                        Range("J" & intOutput).Value = OCRAKetamine
                    End If
                    If OCRAMethohex <> 0 Then
                        Range("L" & intOutput).Value = OCRAMethohex
                    End If
                    If OCRARemi <> 0 Then
                        Range("N" & intOutput).Value = OCRARemi
                    End If
                    If OCRAMorphine <> 0 Then
                        Range("P" & intOutput).Value = OCRAMorphine
                    End If
                    

                    intCount = intCount + 1
                

                Else

                    
                    Range("A" & intOutput).Value = strCurrentOCRAName
                    Range("B" & intOutput).Value = strCurrentOCRAMRN
                    If OCRAFent <> 0 Then
                        Range("D" & intOutput).Value = OCRAFent
                    End If
                    If OCRAHydrom <> 0 Then
                        Range("F" & intOutput).Value = Left(str(OCRAHydrom), 5)
                    End If
                    If OCRAMidaz <> 0 Then
                        Range("H" & intOutput).Value = OCRAMidaz
                    End If
                    If OCRAKetamine <> 0 Then
                        Range("J" & intOutput).Value = OCRAKetamine
                    End If
                    If OCRAMethohex <> 0 Then
                        Range("L" & intOutput).Value = OCRAMethohex
                    End If
                    If OCRARemi <> 0 Then
                        Range("N" & intOutput).Value = OCRARemi
                    End If
                    If OCRAMorphine <> 0 Then
                        Range("P" & intOutput).Value = OCRAMorphine
                    End If
        
                    'reset currentMRN
                    strCurrentOCRAMRN = strOCRAMRN
                    strCurrentOCRAName = strOCRANAME
                    
                    OCRAFent = 0
                    OCRAHydrom = 0
                    OCRAMidaz = 0
                    OCRAKetamine = 0
                    OCRAMorphine = 0
                    OCRARemi = 0
                    OCRAMethohex = 0
                    intOutput = intOutput + 1
                End If
                

                    
                       
        Else
            intCount = intCount + 1
        End If
                
    Loop
    'Read through the pulled data, find matching records in AIMS
    intCount = 4
    intSubCount = 2
    Do While Range("A" & intCount).Value <> ""
        strOCRAMRN = Range("B" & intCount).Value
        Do While (ThisWorkbook.Sheets("AIMS").Range("A" & intSubCount).Value) <> ""
            If strOCRAMRN = ThisWorkbook.Sheets("AIMS").Range("C" & intSubCount).Value Then
                strDrugNameAIMS = ThisWorkbook.Sheets("AIMS").Range("L" & intSubCount).Value
                Select Case Left(strDrugNameAIMS, 7)
                    Case "Fentany"
                        AIMSFent = AIMSFent + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Hydromo"
                        AIMSHydro = AIMSHydro + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Midazol"
                        AIMSMidaz = AIMSMidaz + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Ketamin"
                        AIMSKetamine = AIMSKetamine + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Methohe"
                        AIMSMethohex = AIMSMethohex + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Remifen"
                        AIMSRemi = AIMSRemi + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Morphin"
                        AIMSMorphine = AIMSMorphine + ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                End Select
                Range("Q" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("B" & intSubCount).Value
            End If
            intSubCount = intSubCount + 1

        Loop
        'Write comparison data
        If AIMSFent <> 0 Then
            Range("C" & intCount).Value = AIMSFent
        End If
        If AIMSHydro <> 0 Then
            Range("E" & intCount).Value = AIMSHydro
        End If
        If AIMSMidaz <> 0 Then
            Range("G" & intCount).Value = AIMSMidaz
        End If
        If AIMSKetamine <> 0 Then
            Range("I" & intCount).Value = AIMSKetamine
        End If
        If AIMSMethohex <> 0 Then
            Range("K" & intCount).Value = AIMSMethohex
        End If
        If AIMSRemi <> 0 Then
            Range("M" & intCount).Value = AIMSRemi
        End If
        If AIMSMorphine <> 0 Then
            Range("O" & intCount).Value = AIMSMorphine
        End If
        
        If Range("Q" & intCount).Value = "" Then
            Range("Q" & intCount).Value = "No AIMS Data"
            AIMSMissing = AIMSMissing + 1
        End If
        
        'reset for next loop
        AIMSFent = 0
        AIMSHydro = 0
        AIMSMidaz = 0
        AIMSKetamine = 0
        AIMSMethohex = 0
        AIMSRemi = 0
        AIMSMorphine = 0
        intCount = intCount + 1
        intSubCount = 2
    Loop
    'Read output sheet, look for mismatched AIMS/OCRA Data
    intCount = 4
    Do While Range("A" & intCount).Value <> ""
        AIMSTotalChecked = AIMSTotalChecked + 1
        If Range("C" & intCount).Value = Range("D" & intCount).Value And Range("E" & intCount).Value = Range("F" & intCount).Value And Range("G" & intCount).Value = Range("H" & intCount).Value And Range("I" & intCount).Value = Range("J" & intCount).Value And Range("K" & intCount).Value = Range("L" & intCount).Value And Range("M" & intCount).Value = Range("N" & intCount).Value And Range("O" & intCount).Value = Range("P" & intCount).Value Then
            Range("R" & intCount).Value = "Yes"
            Range("R" & intCount).Interior.Color = RGB(100, 200, 100)
        Else
            Range("R" & intCount).Value = "No"
            Range("R" & intCount).Interior.Color = RGB(200, 100, 100)
            AIMSNoMatch = AIMSNoMatch + 1
            
            intReviewOutput = intReviewOutput + 1
            ThisWorkbook.Sheets("Review").Range("B" & intReviewOutput).Value = Range("B" & intCount).Value
            ThisWorkbook.Sheets("Review").Range("A" & intReviewOutput).Value = Range("A" & intCount).Value
            ThisWorkbook.Sheets("Review").Range("C" & intReviewOutput).Value = Range("S" & intCount).Value
            ThisWorkbook.Sheets("Review").Range("D" & intReviewOutput).Value = Range("Q" & intCount).Value
            If Range("Q" & intCount).Value = "No AIMS Data" Then
                ThisWorkbook.Sheets("Review").Range("E" & intReviewOutput).Value = "No AIMS Record Found"
            Else
                ThisWorkbook.Sheets("Review").Range("E" & intReviewOutput).Value = "AIMS and PCR Do Not Match"
            End If
        End If
        intCount = intCount + 1
    Loop
                
    ' Write to stats sheet
    ThisWorkbook.Sheets("Stats").Range("B1").Value = PCRTotal
    ThisWorkbook.Sheets("Stats").Range("B2").Value = PCRClosed
    If PCRTotal = 0 Then
        ThisWorkbook.Sheets("Stats").Range("B3").Value = "No Data"
    Else
        ThisWorkbook.Sheets("Stats").Range("B3").Value = Int((PCRClosed / PCRTotal) * 100) & "%"
    End If
    
    ThisWorkbook.Sheets("Stats").Range("E1").Value = AIMSTotalChecked
    ThisWorkbook.Sheets("Stats").Range("E2").Value = AIMSNoMatch
    ThisWorkbook.Sheets("Stats").Range("E3").Value = AIMSMissing
    If AIMSTotalChecked = 0 Then
        ThisWorkbook.Sheets("Stats").Range("E4").Value = "No Data"
    Else
        ThisWorkbook.Sheets("Stats").Range("E4").Value = Int((AIMSNoMatch / AIMSTotalChecked) * 100) & "%"
    End If
                

        
        

        

End Sub




