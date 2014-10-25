; Compare TIFIs
; From a given Folder,
; - Gets all source files (txt and csv)
; - uses CTL file to get TIFI name
; - Uses source name and test date to get QA source name
; - reads QA TIFI to get QA source name
; - links TIFIs through source names
; - writes to a results file.
;
; Asghar Jafri, started 26 May 2012
;
;--------------------------------------------------------------
#include <Array.au3>
#include <Constants.au3>
#include <File.au3>
#include <Date.au3>
;
;
 Local $GetTIFIFromBatch = "N"   ; Y : Compare QA with TIFI from ETL Processing,   N: Compare QA with Production Output --> Not tested
;
; Error Constants
;                         TIFI.C053.INPPSB0
Global Const $NoTIFI   = "No TIFI          "
Global Const $NoQATIFI = "No QA TIFI       "
Global Const $DuplicateInQA = "Duplicate in QA"
Global Const $None = "None"
Global Const $NoCTRLFile = "No Control File"
Global Const $Overwrite = 1
Local Const $NotCompared = "Not Compared"
Local Const $EndOfReport = " ***************    END OF REPORT    **************************"

; Array declaration setup.
; Use one 2-Dimensional Array to to make it easier to display status

 Local Const $SrcFilesPR          = 1   ; source file to be processed
 Local Const $SrcFilesQA          = 2
 Local Const $CTLFileNamesPR      = 3
 Local Const $CTLFileNamesQA      = 4
 Local Const $TIFIFileNamesPR     = 5
 Local Const $TIFIFileNamesQA     = 6
 Local Const $TIFISame            = 7 ; result of comparison
 Local Const $ErrorDesc			  = 8 ; Reasons for not submiting to QA
 Local Const $NumParms            = 9 ; Array Dimension for $DataArray - 1 added as 0 is not used.
 ;
Global Const $AQATIFIName   = 1
Global Const $AQAGenSrcName = 2
 ;
 ; Predefined Locations
 Local $SourceFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE"
 ; Local $ProductionTIFIFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE\2012-05-02"
 ; Local $SourceFolder = "c:\test"            ; debug
 Local $ProcessingFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\PROCESSING"    ; new EOD dropbox
 ; Local $ProcessingFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\PROCESSING"    ; new EOD dropbox
 Local $QADropbox = "\\Fsmon001\ece\IS\Broadridge\EOD Testing\dropbox"  ;  previous dropbox
 ; Local $QADropbox = "\\montreal\shares\BPSSEND-QA\eod"                 ; new dropbox
 Local $QAOutput = "\\etl01qa-mtl\infa_shared\SrcFiles\EOD_INTACT\OUTPUT"
 ; Local $QAProcessed = "\\montreal\shares\BPSSEND-QA\eod\Processed"   ; new QA
 Local $QAProcessed = "\\Fsmon001\ece\IS\Broadridge\EOD Testing\dropbox\processed"
 Global $TestFolder = "\\Fsmon001\ece\IS\EOD\Testing\Regression Testing"
 ;
 Local $i  ;  counter
 Local $OutString
 Local $TestDate = 20120522

 Global $Progress = 0
 Local $Progresspct = 0
 ;
; Get Folder name where testing / comparison will take place
;
$TestFolder = FileSelectFolder("Choose a Test Work folder (QA TIFI).","",1+2+4, $TestFolder)
If @error = 1 Then
    MsgBox(4096,"","No Test / Work Folder chosen")
Else
;
	; Get Folder name for source
	;
	$SourceFolder = FileSelectFolder("Choose a folder to take source files from.", "",1+2+4,$SourceFolder)
	;
	; Get filenames for testing
	;
	If @error Then
		MsgBox(4096,"","No Folder chosen")
		exit
	Else
		ProgressOn ( "Comparing Production and QA TIFIs", "Working ...","",-1,-1,2+16)
		ProgressSet(0, "Gathering Files ...","Setting up" )

		Local $iFlag = 0
		Local $FileList = _FileListToArrayMultiSelect($SourceFolder, "*.csv", "*.txt", $iFlag = 0)
		; _ArrayDisplay($FileList)
		;
		; Sort by Batch
		ProgressSet(0, "Sorting by Batch ...","Setting up")
		Local $iMax = $FileList[0]

		; Sorting will mix up the array size stored in element 0
		$FileList[0] = 0 ; Make sure it stays the first element
		_ArraySort($FileList)
		$FileList[0] = $iMax   ; restore original value
		;
		;	_ArrayDisplay($FileList, "$FileList")
		; move the flenames into the array SelectedFiles and initialise the Error Desc component

		ProgressSet(0, "Preparing Data Array","Setting up")
		Local $DataArray[$iMax+1][$NumParms]  ; +1 because the array starts at 0 and the 0 is for the array size, storage starts at 1.

		; Get Filenames into DataArrary. Not all csv / txt files are source files. Only the ones with numeric batch numbers on the left.
		Local $j = 0
		Local $BatchTime
		For $i = 1 to $iMax
			$BatchTime = StringLeft($FileList[$i] ,6)
			if $BatchTime > 0 and $BatchTime < 240000 Then
				if StringMid($FileList[$i],8,6) <> "IWT_WT" AND StringMid($FileList[$i],8,6) <> "IBK_BK" Then
					$j = $j + 1
					$DataArray[$j][$SrcFilesPR] = $FileList[$i]
					$DataArray[$j][$ErrorDesc] = $None
					$DataArray[$j][$TIFISame] = $NotCompared

				endif
			endif
		Next
		;
		$iMax = $j  ; $j contains the number of valid filenames
		;
		; Reclaim resources from $FileList
		_ArrayDelete($FileList,$iMax)
		;
		; Search out all the QA TIFIs in the Test Folder and make the array searchable by generic source name (no date)
		;

		; (Array AQA declared in Function GetQATIFI as the Filist array determines its size.
		Local $Dummy
		GetQATIFI($Dummy)
		;
		$ProgressPct = 100 / $iMax
		ProgressSet(0, "Preparing Output Files","Setting up")

		; Prepare file to write results to
		Local $ResultsFileName
		Local $DT = _NowCalcDate() & _NowTime(4)
		$DT = StringReplace($DT,"/","")
		$DT = StringReplace($DT,":","")
		$ResultsFileName = $TestFolder & "\" & "Results" & $DT & ".txt"
		;
		Local $ResultsFile = FileOpen($ResultsFileName,1)
		;
		Local $ResultsSumFileName
		$ResultsSumFileName = $TestFolder & "\" & "ResultsSum" & $DT & ".txt"
		;
		Local $ResultsSumFile = FileOpen($ResultsSumFileName,1)
		;
		;
		; Prepare file to write all errors to
		Local $DiffsFileName
		$DiffsFileName = $TestFolder & "\" & "Diffs" & $DT & ".txt"
		;
		Global $DiffsFile = FileOpen($DiffsFileName,1)
		;
		ProgressSet($Progress, "Starting Loop " & $ResultsFileName)
		;
		; Write headings
		Local $OutString = "compfromsrc " & $ResultsFileName &  ": All Results" & @CR & @LF & @CR & @LF
		FileWrite($ResultsFile, $OutString)
		;
		Local $OutString = $ResultsSumFileName &  ": Results - Files with Differences only" & @CR & @LF & @CR & @LF
		FileWrite($ResultsSumFile, $OutString)

		Local $OutString = $DiffsFileName &  ": List of all Differences" & @CR & @LF & @CR & @LF
		FileWrite($DiffsFile, $OutString)
		;
		Local $OutString = "Comparing Production TIFIs in : " & $SourceFolder & @CR & @LF
		FileWrite($ResultsFile, $OutString)
		FileWrite($ResultsSumFile, $OutString)
		FileWrite($DiffsFile, $OutString)

		Local $OutString = "with QA TIFIs in : " & $TestFolder & @CR & @LF & @CR & @LF
		FileWrite($ResultsFile, $OutString)
		FileWrite($ResultsSumFile, $OutString)
		FileWrite($DiffsFile, $OutString)
		;
		; This Loop
		;	- Fabricates CTL file name for Production,
		;	- Obtains Production TIFI Name from CTL file
		;   - Fills up this information into $DataArray
		;
		For $i = 1 to $iMax   ; This element comtains the number of filenames entered

			$Progress = ($i - 1) * $ProgressPct
			; msgbox(0,"Progress",$Progress)
			ProgressSet($Progress, "Processing " & $DataArray[$i][$SrcFilesPR],$DataArray[$i][$SrcFilesPR])
			; Format the Filesnames of the files to be dropped into QA by putting the test date into their names
			$DataArray[$i][$SrcFilesQA] = StringMid($DataArray[$i][$SrcFilesPR], 8, 7) & $TestDate & StringRight($DataArray[$i][$SrcFilesPR],7)
			;
			; 	Get Production TIFI name
			;   Fabricate Production Control File Name
			$DataArray[$i][$CTLFileNamesPR] = "CTL_" & StringRight($DataArray[$i][$SrcFilesPR],22) & ".XML"
			ProgressSet($Progress, "Processing " & $DataArray[$i][$CTLFileNamesPR],$DataArray[$i][$SrcFilesPR])
			; Use function GetTifiName to get Production TIFI Name
			Local $PathNameCTL
			$PathNameCTL = $SourceFolder & "\" & $DataArray[$i][$CTLFileNamesPR]
			$DataArray[$i][$TIFIFileNamesPR] = GetTifiName($PathNameCTL)
			ProgressSet($Progress, "Processing " & $DataArray[$i][$TIFIFileNamesPR],$DataArray[$i][$SrcFilesPR])
			; 	Proceed only if TIFI exists
			If $DataArray[$i][$TIFIFileNamesPR] <> $NoTIFI Then   ; TIFI Name has been found
				;  - QA - Fabricate QA Control file name
				$DataArray[$i][$CTLFileNamesQA] = "CTL_" & StringRight($DataArray[$i][$SrcFilesQA],22) & ".XML"
			Else  ; No TIFI found
				$DataArray[$i][$ErrorDesc] = $NoTIFI
			EndIf  ; IF No TIFI

			; If no TIFI then don't continue

			If  $DataArray[$i][$ErrorDesc] = $None Then
				; Example of Production source name : 114742.QUEO_BK20120522AAA.csv
				Local $NoBatchSrcName = StringRight($DataArray[$i][$SrcFilesPR],22)
				Local $GenSrcName = GetGenName($NoBatchSrcName)
				;  - QA - Get TIFI Name and location
				Local $LocQATIFI = GetQATifiName($GenSrcName)
				if $LocQATIFI > 0 Then
					$DataArray[$i][$TIFIFileNamesQA] = $AQA[$LocQATIFI][$AQATIFIName]

					Local $TIFI1 = $SourceFolder & "\" & $DataArray[$i][$TIFIFileNamesPR]
					Local $TIFI2 = $TestFolder   & "\" & $DataArray[$i][$TIFIFileNamesQA]
					ProgressSet($Progress, "Comparing " & StripPath($TIFI1) & " with " & StripPath($TIFI2))
					$DataArray[$i][$TIFISame] = CompTIFI($TestFolder, $TIFI1, $TIFI2)

					if $DataArray[$i][$TIFISame] <> 0 Then
						$DataArray[$i][$ErrorDesc] = "TIFIs Different"
						;
						; Write to resultsSum
						$OutString = $DataArray[$i][$SrcFilesPR] & " " & $DataArray[$i][$TIFIFileNamesPR] & " " &$DataArray[$i][$TIFIFileNamesQA]
						$OutString = $OutString & " Differences: " & $DataArray[$i][$TIFISame] &@CR & @LF
						FileWrite($ResultsSumFile, $OutString)
					EndIf
				Else
					$DataArray[$i][$TIFIFileNamesQA] = $NoQATIFI
					$DataArray[$i][$ErrorDesc] = "TIFIs not compared"
				Endif
			Else ; Error finding production TIFI
				$DataArray[$i][$TIFIFileNamesQA] = $NoQATIFI
				$DataArray[$i][$ErrorDesc] = "TIFIs not compared"
			EndIf
			; _ArrayDisplay($DataArray, "Test Status")
			;
			;Write to results and ResultsSum
			;
			$OutString = $DataArray[$i][$SrcFilesPR] & " " & $DataArray[$i][$TIFIFileNamesPR] & " " &$DataArray[$i][$TIFIFileNamesQA]
			$OutString = $OutString & " Differences: " & $DataArray[$i][$TIFISame] &@CR & @LF
			FileWrite($ResultsFile, $OutString)
			;

		Next   ; Loop for obtaining TIFI's from QA and comparing them to production TIFIs
		;
		; Close all open files

		FileWrite($ResultsFile,    $EndOfReport)
		FileWrite($ResultsSumFile, $EndOfReport)
		FileWrite($DiffsFile,      $EndOfReport)


		FileClose($ResultsFile)
		FileClose($ResultsSumFile)
		FileClose($DiffsFile)

		ProgressOff()

		;Display
		_ArrayDisplay($DataArray, "Test Status",-1,0,"","|","||Source File|QA Source|Production CTL|QA CTL|Production TIFI|QA TIFI|Differences|Error")
		;
	Endif    ; No folder chosen
Endif ; No Work / Test folder chosen
;;
;---------- FUNCTIONS ---------------------------------------------------------
; GetTIFIName
; Reads control file to get tifi name
; Returns TIFI name  or "No Tifi"
Func GetTifiName($CTLFilenameWithPath)

	ProgressSet($Progress, "Getting TIFI Name for " & StripPath($CTLFilenameWithPath))
	;	MsgBox(4096,"","PathNameCTL:" & $CTLFilenameWithPath)
	Local $file = FileOpen($CTLFilenameWithPath)
	; Check if file opened for reading OK
	If $file = -1 Then
		; MsgBox(0, "Error", "Unable to open file: ", $CTLFilenameWithPath)
		Return "No Control File"
	Else

		Local $CTL_Rec = FileReadLine($file)
	;	MsgBox(4096,"","CTL Rec" & $CTL_Rec)
		Local $TIFINamePos = StringInStr($CTL_Rec, "<TIFIDataSetName>")
	;    If there's no tifi
		if $TIFINamePos = 0 Then
			$TIFIName = $NoTIFI
		Else
			$TIFIName = StringMid($CTL_Rec, $TIFINamePos + 17, 17)
		EndIf

		FileClose($file)
		Return $TIFIName
	EndIf

EndFunc   ; GetTIFName
;-------------------------------------------------------------
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
;------------------------------
;--------------------------------
;$FontDir = "Directory Path"
;$aArray = _FileListToArrayMultiSelect($FontDir, "*.ttf",  "*.otf", 1)

Func _FileListToArrayMultiSelect($dir, $search1, $search2, $iFlag = 0)
    Local $FileList, $FileList1, $Num, $Err = 0, $Err2 = 0
    $FileList = _FileListToArray($dir, $search1, $iFlag)
    If @error > 0 Then  $Err = 1
    $FileList1 = _FileListToArray($dir, $search2, $iFlag)
    If @error > 0 Then $Err2 = 1

    If ($Err2 = 1) And ($Err = 0) Then Return $FileList
    If ($Err2 = 0) And ($Err = 1) Then Return $FileList1
    If ($Err2 = 0) And ($Err = 0) Then
        $Num = UBound($FileList)
        _ArrayConcatenate($FileList, $FileList1)
        $FileList[0] = $FileList[0] + $FileList[$Num]
        _ArrayDelete($FileList, $Num)
        Return $FileList
    EndIf
    MsgBox(0, "", "No Files\Folders Found.")
EndFunc  ;==>_FileListToArrayMultiSelect

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
;
Func GetQATIFI($Dummy)
;
	;
	; Global Const $AQATIFIName   = 1
	; Global Const $AQAGenSrcName = 2
	;
	Local $FileFilter = "TIFI.Q053.*"
	Local $FileList = _FileListToArray($TestFolder,$FileFilter)
	If @error = 1 Then
		MsgBox(0, "", "No Folders Found.")
		Exit
	EndIf
	If @error = 4 Then
		MsgBox(0, "", "No Files Found.")
		Exit
	EndIf
	;
	Global $AQA[$FileList[0]+1][3]
	;
	For $i = 1 to $FileList[0]

		Local $file1 = FileOpen($TestFolder & "\" & $FileList[$i])

		;Readin first record ito $Instring
		Local $line1 = FileReadLine($file1)

		FileClose($file1)

		; Copy TIFI Name to AQA
		$AQA[$i][$AQATIFIName] = $FileList[$i]

		; now put in generic source name
		;Name of source is in position 191. length 22 characters
		Local $src = StringMid($line1,191,22)

		$AQA[$i][$AQAGenSrcName] = GetGenName($src)
	Next

	Return 0
EndFunc ; GetQATIFI
;~-----------------------------------------------------------------
;
Func GetGenName($src)
;
	; example of a source name : FIRE_AF20120522001.TXT
	Local $GenName = StringLeft($src,7) & " " & StringRight($src,7)
	Return $GenName

EndFunc  ; GetGenName

;-----------------------------------------------------------------------
FUNC GetQATifiName($GenSrcName)
	Local $index
	$index = _ArraySearch($AQA, $GenSrcName, 0, 0, 0, 0,1,2)
	Return $index
EndFunc
;----------------------------------
