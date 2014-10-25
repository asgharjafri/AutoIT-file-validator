; Regression Pick testing
; From a given Folder, copies files by batch to QA then compares the resulting TIFIs with the Production TIFIs
;
; Asghar Jafri, started 2 May 2012
; Select source files to test
; Copy them to QA
; Copy Production TIFI to testing folder
; Copy new QA TIFI to testing folder
; Compare the TIFI files
;
; Possible changes
; Make one multidim ARRAY to make it easier for display and copy
;--------------------------------------------------------------
#include <Array.au3>
#include <Constants.au3>
#include <File.au3>
#include <Date.au3>
#include <_FileFindEx.au3>
;
;
 Local Const $Todaysdate = @YEAR & @MON & @MDAY
 Local Const $SleepInt = 15000 ; Milliseconds to wait between check for presence of file
 ; debug Local Const $SleepInt = 1000 ; Milliseconds to wait between check for presence of file
 Local $MaxWaitTime = 12 * 60 * 1000  ; 12 minutes in milliseconds
 ; debug Local $MaxWaitTime = 1 * 5 * 1000  ; 2 minutes in milliseconds (for testing)
 Local $GetTIFIFromBatch = "N"   ; Y : Compare QA with TIFI from ETL Processing,   N: Compare QA with Production Output --> Not tested
;
; Error Constants
Global Const $NoTIFI = "No TIFI"
Global Const $DuplicateInQA = "Duplicate in QA"
Global Const $None = "None"
Global Const $NoCTRLFile = "No Control File"
Global Const $Overwrite = 1
Local Const $NotCompared = "Not Compared"


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
 Local Const $EndOfReport = " ***************    END OF REPORT    **************************"
 ;
 ; Predefined Locations
 ; Local $SourceFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE"
 ; Local $ProductionTIFIFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE\2012-05-02"
 Local $SourceFolder = "c:\test"            ; debug
 Local $ProcessingFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\PROCESSING"    ; new EOD dropbox
 ; Local $ProcessingFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\PROCESSING"    ; new EOD dropbox
 ; Local $QADropbox = "\\Fsmon001\ece\IS\Broadridge\EOD Testing\dropbox"  ;  previous dropbox
 Local $QADropbox = "\\montreal\shares\BPSSEND-QA\eod"                 ; new dropbox
 ; Local $QADropbox = "C:\Users\ajafri\Desktop\Dummy EOD"    ; test dropbox  -- debug
 Local $QAOutput = "\\etl01qa-mtl\infa_shared\SrcFiles\EOD_INTACT\OUTPUT"
 Local $QAProcessed = "\\montreal\shares\BPSSEND-QA\eod\Processed"   ; new QA
 ; Local $QAProcessed = "\\Fsmon001\ece\IS\Broadridge\EOD Testing\dropbox\processed"   ; old QA processed
 ; Local $QAProcessed = "C:\Users\ajafri\Desktop\Dummy EOD\Processed"            ; debug
 ; Local $TestFolder = "\\Fsmon001\ece\IS\EOD\Testing\Regression Testing"
 Local $TestFolder = "C:\temp"              ; test test folder
 ;
 Local $i  ; increment counter
 Local $FirstInBatch   ; Stores the first file in the batch
 Global $Progress = 0
 Local $Progresspct = 0
 ;
 ; Get Folder name where test/QA TIFIs will be placed
;
$TestFolder = FileSelectFolder("Choose a Test Work folder. (For QA files)","",1+2+4, $TestFolder)
If @error = 1 Then
    MsgBox(4096,"","No Test / Work Folder chosen")
Else
;
	; Get Folder name for source
	$SourceFolder = FileSelectFolder("Choose a folder to take source files from.", "",1+2+4,$SourceFolder)
	; Get filenames for testing
	If @error Then
		MsgBox(4096,"","No Folder chosen")
		exit
	Else

		ProgressOn ( "Regression Testing", "Preparing ...","",-1,-1,2+16)

 ; Files setup.

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
		Local $OutString = "REGTESTF " & $ResultsFileName &  ": All Results" & @CR & @LF & @CR & @LF
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
		ProgressSet(0, "Gathering File Names ..." )

		Local $iFlag = 0
		Local $FileList = _FileListToArrayMultiSelect($SourceFolder, "*.csv", "*.txt", $iFlag = 0)
		; _ArrayDisplay($FileList)
		;
		; Sort by Batch
		Local $iMax = $FileList[0]

		; Sorting will mix up the array size stored in element 0
		$FileList[0] = 0 ; Make sure it stays the first element
		_ArraySort($FileList)
		$FileList[0] = $iMax   ; restore original value
		;
		Local $DataArray[$iMax+1][$NumParms]
		;
		; Copy the files to QA
		;-- Move and rename the files so they will run in QA
		;----- Strip the batch time, put in today's date
		;
		Local $src
		Local $dest

		$ProgressPct = 100 / $iMax

		; This Loop
		;	- Fabricates QA names by replacing the actual dates with today's date in the source filename
		;	- Fabricates CTL file name for Production,
		;	- Obtains Production TIFI Name from CTL file
		;	-  Fills up this information into $DataArray
		;

		; Copy source files to QA in Batches, and then compare them in batches also
		; Reason for running loop twice: Files have to be copied to QA in batches to simulate Production ETL.
		; If the functions of the second loop are carried out in the firstloop then the second loop

		; First loop copies all the source files in the batch to QA
		; second loop gets the TIFIs from QA and compares them.

		Local $FilesSent = 0

		Local $j = 1   ; counter for filtered files --> to be sent for testing ; counter for DataArray
 		Local $BatchTime

		$i = 1
		While $i <= $iMax   	; $iMax comtains the number of filenames gathered in FileList. Not accurate as files are further filtered
								; $i is also incremented in the loops following
			Local $ThisBatch = StringLeft($FileList[$i] ,6)
			; Work by batch, This loop will increment i and cycle through the FileList

			$BatchTime		= StringLeft($FileList[$i] ,6)
			$LastBatchTime 	= $BatchTime  	; Make the same the first time to ensure at least one run of the loop.
											; Minimum one file must be copied
			$FirstInBatch = $j   ; Remember $j value for next loop
			While $BatchTime = $LastBatchTime and $i <= $iMax                             ; While on same batch

				; Don't process IWT,IBK or files without batch times
				Local $BatchTime = StringLeft($FileList[$i] ,6)
				if $BatchTime > 0 and $BatchTime < 240000 Then
					if StringMid($FileList[$i],8,6) <> "IWT_WT" AND StringMid($FileList[$i],8,6) <> "IBK_BK" Then

						$Progress = ($j - 1) * $ProgressPct
						ProgressSet($Progress, "Processing ",$FileList[$i])

						; start filling up DataArray
						$DataArray[$j][$SrcFilesPR] = $FileList[$i]
						$DataArray[$j][$ErrorDesc]  = $None
						$DataArray[$j][$TIFISame]   = $NotCompared

						; Format the Filenames of the files to be dropped into QA by putting today's date into their names
						$DataArray[$j][$SrcFilesQA] = StringMid($DataArray[$j][$SrcFilesPR], 8, 7) & $Todaysdate & StringRight($DataArray[$j][$SrcFilesPR],7)
						; Check if AsOf. Change name if necessary
						; sample : MCQC_AF20120531001.txt
						;          1234567890123456789012
						;                 543210987654321
						if StringMid($DataArray[$j][$SrcFilesQA],6,2) = "AF" Then
							$DataArray[$j][$SrcFilesQA] = StringLeft($DataArray[$j][$SrcFilesQA],5) & "AR" & StringRight($DataArray[$j][$SrcFilesQA],15)
						EndIf
						;
						; Check if duplicate in QA
						If FileExists($QAProcessed & "\" & $DataArray[$j][$SrcFilesQA]) = 1 Then
							$DataArray[$j][$ErrorDesc] = $DuplicateInQA
						Else   ; No Duplicate in QA,
							;							; Fabricate Production Control File Name.
							$DataArray[$j][$CTLFileNamesPR] = "CTL_" & StringRight($DataArray[$j][$SrcFilesPR],22) & ".XML"

							; Use function GetTifiName to get Production TIFI Name
							Local $PathNameCTL
							$PathNameCTL = $SourceFolder & "\" & $DataArray[$j][$CTLFileNamesPR]
							$DataArray[$j][$TIFIFileNamesPR] = GetTifiName($PathNameCTL)
							If $DataArray[$j][$TIFIFileNamesPR] <> $NoTIFI Then
								; 	Copy Source files to QA )

								if $DataArray[$j][$SrcFilesPR] <> "" OR $DataArray[$j][$SrcFilesQA] <> "" Then
									$src  = $SourceFolder & '\' & $DataArray[$j][$SrcFilesPR]
									Filecopy($src, $QADropbox & "\" & $DataArray[$j][$SrcFilesQA] )	; Drop source into QA dropbox
									$FilesSent = $FilesSent + 1
								Else
									MsgBox(0,"Blank src",  $DataArray[$j][$SrcFilesPR] & $DataArray[$j][$SrcFilesQA])
								Endif

								; Prepare variables for checking next batch									;
								$LastBatchTime = $BatchTime													; get next batch number to compare with on the While
								$j = $j + 1		    														; Don't go over the Array size
							Endif    ; NO TIFI ; ok to put into QA
						Endif	; No Duplicate in QA
					endif   ; skip on the IWT and IBK files
				endif    ; Don't process any non-batch files

				; onto the next record
				$i = $i + 1
				if $i <= $iMax Then
					$BatchTime = (StringLeft($FileList[$i] ,6))
				Else
					$BatchTime = 0  ; to signal loop checker to end
				EndIf

			WEnd
			;
			; Loop follows above loop to wait for each CTL file and TIFI
			; MsgBox(0,"Ended Batch", $LastBatchTime)
			if $FilesSent > 0 Then
				Local $k
				For $k = $FirstInBatch to $i
					; Process only if valid to drop into QA
					if $DataArray[$k][$ErrorDesc] = $None Then
						 ; - QA - Get TIFI Name and location
						; ------- QA may not have processed the file yet so we have to wait for
						; the Control File and the TIFI
						; Due to naming convention in QA Processing having a timestamp at the end of the name
						; an exact name can only be determined after the Control File is created. Then the search function with a wildcard is used.
						; Sample name of Control file : CTL_MCQC_BK20120531666.CSV.XML-19172468

						$DataArray[$k][$CTLFileNamesQA] = "CTL_" & StringRight($DataArray[$k][$SrcFilesQA],22) & ".XML"
						$PathNameCTL = $QAProcessed & "\" & $DataArray[$k][$CTLFileNamesQA]

						Local $ok = WaitForFile($PathNameCTL, "</Report>", $MaxWaitTime, $SleepInt)
						if $ok = 0 Then
							MsgBox(0,"Error: No QA Control File", $PathNameCTL)
							$DataArray[$k][$ErrorDesc] = $NoCTRLFile
						 Else
							ProgressSet($Progress, "Getting QA TIFI name " & $DataArray[$k][$SrcFilesPR])
							$DataArray[$k][$TIFIFileNamesQA] = GetTifiName($PathNameCTL)

							; copy the QA TIFI file to testing
							$src = $QAOutput & '\' & $DataArray[$k][$TIFIFileNamesQA]
							$ok = WaitForFile($src, "REC-CNT=",$MaxWaitTime, $SleepInt)
							if $ok = 0 Then
								MsgBox(0,"Error: No QA TIFI", $src)
								$DataArray[$k][$ErrorDesc] = $NoTIFI
							Else
								ProgressSet($Progress, "Copying QA TIFI " & $Src)
								FileCopy($src, $TestFolder,$Overwrite)   ;; Flag set to overwrite existing files
								;
								; Compare the TIFIs
								; MsgBox(0."Check incs,TIFISame: " & $TIFISame & "i: " & $i & "TIFName " & $TIFIFileNamesPR & " TIFI QA: " $TIFINamesQA
								Local $TIFI1 = $SourceFolder & "\" & $DataArray[$i][$TIFIFileNamesPR]
								Local $TIFI2 = $TestFolder   & "\" & $DataArray[$i][$TIFIFileNamesQA]
								ProgressSet($Progress, "Comparing " & StripPath($TIFI1) & " with " & StripPath($TIFI2))
								$DataArray[$i][$TIFISame] = CompTIFI($TestFolder, $TIFI1, $TIFI2)
								;
								if $DataArray[$k][$TIFISame] > 0 Then
									$DataArray[$k][$ErrorDesc] = "TIFIs Different"
								Endif
							EndIf
						EndIf
					EndIf   ; ErrorDesc = $None
				Next   ; k loop ; drops source into QA

				$FilesSent = 0
			Endif ; If FilesSent > 0

		WEnd   ; i <= $iMax

		ProgressOff()

		;
		_FileWriteFromArray($ResultsFileName, $DataArray)
		;
		;Display
		_ArrayDisplay($DataArray, "Test Status",-1,0,"","|","||Source File|QA Source|Production CTL|QA CTL|Production TIFI|QA TIFI|Diff at Rec|Error")
		;

	; Close all open files

		FileWrite($ResultsFile,    $EndOfReport)
		FileWrite($ResultsSumFile, $EndOfReport)
		FileWrite($DiffsFile,      $EndOfReport)

		FileClose($ResultsFile)
		FileClose($ResultsSumFile)
		FileClose($DiffsFile)


	Endif    ; No folder chosen
Endif ; No Work / Test folder chosen


; Close all open files

FileWrite($ResultsFile,    $EndOfReport)
FileWrite($ResultsSumFile, $EndOfReport)
FileWrite($DiffsFile,      $EndOfReport)


FileClose($ResultsFile)
FileClose($ResultsSumFile)
FileClose($DiffsFile)


;;
;---------- FUNCTIONS ---------------------------------------------------------
; GetTIFIName
; Reads control file to get tifi name
; Returns TIFI name  or "No Tifi"
Func GetTifiName($CTLFilenameWithPath)

	ProgressSet($Progress, "Getting TIFI Name from " & StripPath($CTLFilenameWithPath))
	;	MsgBox(4096,"","PathNameCTL:" & $CTLFilenameWithPath)
	Local $FileSearchName = $CTLFilenameWithPath & "*"
	Local $Search = _FileFindExFirstFile($FileSearchName)
	if $Search <> -1 Then
		Local $FullFileName = StripName($CTLFilenameWithPath) & $Search[0]
		$file1 = FileOpen($FullFileName)
		$CTL_Rec = FileReadLine($FullFileName, -1)
		_FileFindExClose($Search)

		Local $TIFINamePos = StringInStr($CTL_Rec, "<TIFIDataSetName>")
	Else
		$TIFINamePos = 0
	Endif

	;    If there's no tifi
	if $TIFINamePos = 0 Then
		$TIFIName = "No TIFI"
	Else
		$TIFIName = StringMid($CTL_Rec, $TIFINamePos + 17, 17)
	EndIf

	Return $TIFIName

EndFunc   ; GetTIFName
;----------------------------------------------------
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
;----------------------------------------------------
; Function WaitForFile
; $File: File to Wait for
; $Confirm: A string in the last record of $File (to confirm it's at EOF)
; $HowLongToWait: How many miliseconds before giving up
; Returns 0 if file did not show up
;
; uses udf _FileFindFirstFile found at website : http://sites.google.com/site/ascend4ntscode/filefindex
;
; SAMPLE CALL : _FileFindExNextFile($aFileFindArray)
;				_FileFindExFirstFile($sFolder & $sSearchWildCard)
; sample of name in processed : CTL_MBTR_BK20120530PG2.csv.XML-15384906
;


FUNC WaitForFile($File, $confirm, $HowLongToWait, $Sleepint)

	Local $file1
	Local $ElapsedTime = 0
	Local $FileOpen = "No"
	Local $line1
	;
	ProgressSet($Progress, "Waiting: " & StripPath($File) & ", " & ($HowLongToWait - $ElapsedTime)/1000 & " s.")
	;
	;MsgBox(0,"","Entering Loop" & $ElapsedTime & " for " & $HowLongToWait & "int: " & $Sleepint)
	; $file1 = FileOpen($File)

	Local $FileSearchName = $File & "*"
	Local $Search = _FileFindExFirstFile($FileSearchName)
	While $ElapsedTime < $HowLongToWait AND $Search = -1
		If $Search = -1 Then
			ProgressSet($Progress, "Waiting: " & StripPath($File) & ", " & ($HowLongToWait - $ElapsedTime)/1000 & " s.")
			Sleep($SleepInt)
			$ElapsedTime = $ElapsedTime + $SleepInt

			 $Search = _FileFindExFirstFile($FileSearchName)
		Endif
	WEnd
	;
	if $Search <> -1 Then
		Local $FullFileName = StripName($File) & $Search[0]
		$file1 = FileOpen($FullFileName)
		$Line1 = FileReadLine($FullFileName, -1)

		_FileFindExClose($Search)
	Else
		$Line1 = ""
	Endif
	; Return a 0 if the last rec does not contain the expected information
	Return StringInStr($Line1, $confirm)

EndFunc  ; WaitForFile
;---------------------------------------------

;$FontDir = "Directory Path"
;$aArray = _FileListToArrayMultiSelect($FontDir, "*.ttf",  "*.otf", 1)
;
; Thanks to http://www.autoitscript.com/forum/topic/91441-wildcards-and-filelisttoarray/
;

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