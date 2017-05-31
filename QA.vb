'-----------------------------------------------------------------------------------------------------------------------
' An Excel Macro to read and interpret narcotic dispensing data from Omnicell PCR reports, and compare it to AIMS anesthesia records.
' Both programs write out data in a spreadsheet that I would conservatively describe as "whack," so I built a workbook to read them
' for me. Omnicell report goes in a sheet named "Ocra," AIMS report goes in a sheet named "Aims," Button and grid go on a "Solutions"
' sheet, and Metadata gets written to a "Stats" sheet. I'd just upload the file, but there's patient data in there and that's an actual
' crime.
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
'      matches to the solutions sheet
'   5) The easy part: Compare data within the 'Solutions' sheet, write desired results.
'   6) Metadata: In progress, gotta get a list of stuff to present


Private Sub CommandButton1_Click()
'clear sheet
    With Sheets("Solutions")
        .Rows("4:" & .Rows.Count).Delete
    End With
    ThisWorkbook.Sheets("Stats").Range("B1:B10").Value = ""
    
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
    
'Incrementation stuff
    Dim intCount As Integer
    Dim intSubCount As Integer
    Dim intOutput As Long
    Dim strCurrentOCRAMRN As String
    Dim strCurrentOCRAName As String
    strCurrentOCRAName = "false"
    strCurrentOCRAMRN = "false"

'Patient Data
    Dim strOCRAMRN As String
    Dim strAIMSMRN As String

    
    PCRTotal = 0
    PCRClosed = 0
    intCount = 2        'start reading OCRA on the first data row, skipping labels
    intOutput = 4       'start writing on the first empty row, skipping button/labels
    
    'Read PCR report (OCRA)
    Do While (ThisWorkbook.Sheets("OCRA").Range("A" & (intCount)).Value) <> ""
        'PCR closure stuff for stats sheet
        PCRTotal = PCRTotal + 1
        If ThisWorkbook.Sheets("OCRA").Range("R" & intCount).Value = 0 Then
            PCRClosed = PCRClosed + 1
        End If
        'Patient identifiers and transaction info
        If ThisWorkbook.Sheets("OCRA").Range("AN" & intCount).Value = "W" And Left(ThisWorkbook.Sheets("OCRA").Range("AE" & intCount).Value, 7) <> "BIW_PRO" Then
            strOCRAMRN = ThisWorkbook.Sheets("OCRA").Range("B" & intCount).Value
            strOCRANAME = ThisWorkbook.Sheets("OCRA").Range("D" & intCount).Value
        
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
                Select Case Left(strDrugNameAIMS, 7) 'read only first 7 characters to ignore special routes of admin i.e. "morphine - spinal"
                    Case "Fentany"
                        Range("C" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Hydromo"
                        Range("E" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Midazol"
                        Range("G" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Ketamin"
                        Range("I" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Methohe"
                        Range("K" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Remifen"
                        Range("M" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                    Case "Morphin"
                        Range("O" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("M" & intSubCount).Value
                End Select
                Range("Q" & intCount).Value = ThisWorkbook.Sheets("AIMS").Range("B" & intSubCount).Value
            End If
            intSubCount = intSubCount + 1
            If Range("Q" & intCount).Value = "" Then
                Range("Q" & intCount).Value = "No AIMS Data"
            End If
        Loop
        intCount = intCount + 1
        intSubCount = 2
    Loop
    'Read Solutions sheet, look for mismatched AIMS/OCRA Data
    intCount = 4
    Do While Range("A" & intCount).Value <> ""
        If Range("C" & intCount).Value = Range("D" & intCount).Value And Range("E" & intCount).Value = Range("F" & intCount).Value And Range("G" & intCount).Value = Range("H" & intCount).Value And Range("I" & intCount).Value = Range("J" & intCount).Value And Range("K" & intCount).Value = Range("L" & intCount).Value And Range("M" & intCount).Value = Range("N" & intCount).Value And Range("O" & intCount).Value = Range("P" & intCount).Value Then
            Range("R" & intCount).Value = "Yes"
            Range("R" & intCount).Interior.Color = RGB(100, 200, 100)
        Else
            Range("R" & intCount).Value = "No"
            Range("R" & intCount).Interior.Color = RGB(200, 100, 100)
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

End Sub


