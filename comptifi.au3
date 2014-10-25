#include <Date.au3>

;Local $SourceFolder = "C:\temp"
Local Const $EndOfReport = " ***************    END OF REPORT    **************************"
Local $message = "Choose TIFI for Comparison"
$SourceFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE"
Local $TIFI1 = FileOpenDialog("TIFI 1", $SourceFolder , "Production TIFIs (TIFI.C053.*)", 1 + 2 )
if @error = 1 Then
	MsgBox(0,"Error","File not chosen")
	Exit
EndIf

$SourceFolder = "\\Fsmon001\ece\IS\EOD\Testing\Regression Testing"
Local $TIFI2 = FileOpenDialog("TIFI 2", $SourceFolder , "QA TIFIs (TIFI.Q053.*)", 1 + 2 )
if @error = 1 Then
	MsgBox(0,"Error","File not chosen")
	Exit
EndIf

; ProgressOn("Comparing 2 TIFIs", $TIFI1 & " and " & $TIFI2, "Working..."-1,-1,1+2+16)

; Put in heading for DIffs FileChangeDir
Local $ResultsFileName
Local $DT = _NowCalcDate() & _NowTime(4)
$DT = StringReplace($DT,"/","")
$DT = StringReplace($DT,":","")
;
; Prepare file to write all errors to
Local $DiffsFileName
$DiffsFileName = $SourceFolder & "\" & "Diffs" & $DT & ".txt"
;
Global $DiffsFile = FileOpen($DiffsFileName,1)
;
;
; Write headings

Local $OutString = "COMPTIFI.AU3, " & $DiffsFileName &  ": List of Differences" & @CR & @LF & @CR & @LF
FileWrite($DiffsFile, $OutString)
;
Local $err = CompTIFI($SourceFolder, $TIFI1, $TIFI2)

FileWrite($DiffsFile,      $EndOfReport)
FileClose($DiffsFile)

Run("notepad.exe " & $DiffsFileName)

Msgbox(0," Result of comparison: ", "error(s) : " & $err & @CR & @LF & "Results in " & $DiffsFileName )

;-------------------------------------------------------------
; COMPTIFI
;
;Compare each record except first and last
; Don't compare daily transaction code column
;
; Diffs file must be opened and closed from the callig program
FUNC CompTIFI($WorkFolder, $TIFI1, $TIFI2)

	Local $recnum = 0
	Local $numerr = 0
	; Local $Progress = "Comparing..."
	Local $LastRec   ; indicates last record of file (only for regular TIFI files with the string "REC-CNT=" in the last record
	Local $LEOF = "F"
	Local $file1, $file2
	Local $line1,$line2  ; records read form file
	Local $cline1, $cline2   ; records processed for comparison
	Local $MaxErr = 50   ; Maximum number of errors to report
	Local Const $NullStr = ""
	Local Const $fmt  = "%22s | %17s | %4u | %4u | %10s |"  ; format for diffs report
	Local Const $fmtb = "%22s | %17s | %4s | %4s | %10s |"  ; format for diffs report blank line between errors
	Local Const $OutStringB = stringformat($fmtb,$NullStr,$NullStr,$NullStr,$NullStr,$NullStr) &@CR & @LF
	Local $OutString1
	Local $OutString2

	Local Const $OneBlank    = " "
	;						    123456
	Local Const $SixBlanks   = "      "
	;						    12345678
	Local Const $EightBlanks = "        "
	;						    1234567890
	Local Const $TenBlanks   = "          "

	Local Const $FidessaNames = "ALTA,FIRE,MCQC"
	;
	Local $FileSizePR = FileGetSize ($TIFI1)
	Local $FileSizeQA = FileGetSize ($TIFI2)
	;
	if $FileSizePR <> $FileSizeQA Then
		$MaxErr = 10
	Else
		$Maxerr = 50
	Endif

	;
	Local $OutString = "Production TIFI: " & $TIFI1 & ",Size: " & $FileSizePR & @CR & @LF
	FileWrite($DiffsFile, $OutString)

	Local $OutString = "QA         TIFI: " & $TIFI2 & ",Size: " & $FileSizeQA & @CR & @LF & @CR & @LF
	FileWrite($DiffsFile, $OutString)
	;
	;
	$file1 = FileOpen($TIFI1)
	$file2 = FileOpen($TIFI2)
	;
	While $LEOF = "F" and $numerr <= $MaxErr

		; Read Production TIFI
		$line1 = FileReadLine($file1)
		Local $Len = StringLen( $line1 )
		If @error = -1 OR StringLen( $line1 )= 0 Then
			$LEOF = "T"
			ExitLoop
		Endif

		; MsgBox(0, "PR Line read:", $line)

		; Read QA TIFI
		$line2 = FileReadLine($file2)
		$Len = StringLen( $line2 )
		If @error = -1 OR StringLen( $line2 )= 0 Then
			ExitLoop
		Endif
		; MsgBox(0, "QA Line read:", $line2)

		$recnum = $recnum + 1

		; Check if Fidessa file. This comes in handy when checking for an extra DX in the Trailer codes.
				;  $Local Const $FidessaNames = "ALTA,FIRE,MCQC"   ; Declared above, here for reference.
		if $recnum = 1  Then  						; first record contains the name of the source file.
			$NameSrc = StringMid($line1,191,4)
			Local $NamePos = StringInStr($FidessaNames,$NameSrc)
			$Fidessa = $NamePos > 0

			Local $NameSrcFile = StringMid($line1,191,22)
		Endif

		$LastRec = StringInStr($line1, "REC-CNT=")
		if ($recnum <> 1) AND ($LastRec = 0) Then ; no need to check the header or last record, they will always be different

			; Series of replacements to remove some irritants
			; 1. The left stops before the offending string starts
			; The mid starts after the offending string stops
			; A blank is put into the missing space
			; There has to be an easier way to do this.

			; Remove the Daily transaction Code so it doesn't get compared
			$cline1 = StringLeft($line1, 449) & $SixBlanks & StringMid($line1,456)
			$cline2 = StringLeft($line2, 449) & $SixBlanks & StringMid($line2,456)

			; For all subsequnt comparisons, use cline variables instead of line
			; Remove the T at 438
			; msgbox(0,"437+5",StringMid($cline2,437,5))
			$cline1 = StringLeft($cline1, 437) & $OneBlank & StringMid($cline1,439)
			$cline2 = StringLeft($cline2, 437) & $OneBlank & StringMid($cline2,439)

			; Remove the Date at 56 so it doesn't get compared
			$cline1 = StringLeft($cline1, 55) & $EightBlanks & StringMid($cline1,64)
			$cline2 = StringLeft($cline2, 55) & $EightBlanks & StringMid($cline2,64)

			; Effective Date: If Production has an effective date then check in QA, otherwise ignore. Production is 1, QA is 2
			Local $EffectiveDatePR = StringMid($cline1, 148, 8)
			Local $EffectiveDateQA = StringMid($cline2, 148, 8)
			Local $LenEffectiveDateQA = StringLen($EffectiveDateQA)

			if $EffectiveDatePR = $EightBlanks AND $LenEffectiveDateQA > 0 Then
				$cline2 = StringLeft($cline2, 147) & $EightBlanks & StringMid($cline2,156)
			Endif
			;
			; If Fidessa file and Trailers don't match then check if they match if a DX is appended to the left of Production trailer.
			; If then they match then it's not an error.

			; Check if the error is in the Trailer codes for FIDESSA files. Fidessa Files are MCQC, FIRE and ALTA
			; here is an error to be ignored - if a DX is appended to the left of the trailer code string.

			if  $Fidessa Then
				Local $PRTrailer = StringStripWS(  StringMid($Cline1,350,10) ,8 )
				Local $QATrailer = StringStripWS(  StringMid($Cline2,350,10) ,8 )

				; if QA Trailer is only the DX appended to the Production trailer then no problem.
				Local $QATrailerWDX = "DX" & $QATrailer
				if $PRTrailer = $QATrailerWDX Then
					; Blank out Trailer codes for the comparison
					$cline1 = StringLeft($cline1, 349) & $TenBlanks & StringMid($cline1,360)
					$cline2 = StringLeft($cline2, 349) & $TenBlanks & StringMid($cline2,360)
				Endif ; smatch with DX

			Endif  ; Fidessa

			; Comparison of Production Record with QA record
			Local $CompResult = StringCompare($cline1, $cline2)
			if $CompResult <> 0 Then
				;			;
				Local $i
				For $i = 1 to StringLen($cline1)

					Local $c1 = StringMid($cline1,$i,1)
					Local $c2 = StringMid($cline2,$i,1)
					if $c1 <> $c2 Then
						$errpos = $i
						$numerr = $numerr + 1

						; error position has been found
						; Write to the Diffs file
						Local $RecSamplePR = StringMid($CLine1,$errpos - 5, 10)
						Local $RecSampleQA = StringMid($CLine2,$errpos - 5, 10)

						if $numerr = 1 Then  ; write heading for first error only
							;Local $OutString = $NameSrc & " | " & StripPath($TIFI1) & " | " & StripPath( $TIFI2 ) & " | Rec: " & $recnum & " | Pos:  " & $errpos
							;$OutString = $OutString & " |" & $RecSamplePR & " | " &  $RecSampleQA & "|" &@CR & @LF
							$Outstring1 = stringformat($fmt,$NameSrcFile,StripPath($TIFI1),$recnum,$errpos,$RecSamplePR) & @CR & @LF
							$Outstring2 = stringformat($fmt,$NullStr,StripPath($TIFI2),$recnum,$errpos,$RecSampleQA) & @CR & @LF

						else

							$Outstring1 = stringformat($fmt,$NullStr,StripPath($TIFI1),$recnum,$errpos,$RecSamplePR) &@CR & @LF
							$Outstring2 = stringformat($fmt,$NullStr,StripPath($TIFI2),$recnum,$errpos,$RecSampleQA) &@CR & @LF

						endif
						FileWrite($DiffsFile, $OutString1)
						FileWrite($DiffsFile, $OutString2)
						FileWrite($DiffsFile, $OutStringB)

					endif

					if $numerr >= $MaxErr Then
						FileWrite($DiffsFile,"                        *** Error Limit (" & $Maxerr & ") Reached ***" & @CR & @LF)
						ExitLoop 2
					EndIf

				Next
				; error position has been found
				;
			EndIf
		EndIf

	WEnd

	FileClose($file1)
	FileClose($file2)

	; Compare File Size
	if $FileSizePR <> $FileSizeQA Then
		$numerr = -1  ;
	Else
	;
	EndIf  ; Size Different ?
;
	Local $ReturnResult
	Switch $NumErr
		Case 0
			$ReturnResult = 0
		Case -1
			$ReturnResult = "Size Diff"
		Case Else
			$ReturnResult = $numerr
	EndSwitch

	Return $ReturnResult
;
EndFunc   ;CompTIFI
;-------------------------------------------------
;---------------------------------------------------
; Function Strip path and get filename only
FUNC StripPath($PathAndName)

	Local $NamePos
	Local $FileNameonly

	$NamePos = StringLen($PathAndName) - StringInStr($PathAndName,"\",0,-1)
	$FileNameonly = StringRight($PathAndName,$NamePos)

	Return $FileNameOnly

EndFunc ; StripPath

;~-----------------------------------------------------------------
;;---------------------------------------------------
; Function Strip name and get path only
FUNC StripName($PathAndName)

	Local $NamePos
	Local $Pathonly

	$NamePos = StringInStr($PathAndName,"\",0,-1)
	$PathOnly = StringLeft($PathAndName,$NamePos)

	Return $PathOnly

EndFunc ; StripName

;~-----------------------------------------------------------------
;