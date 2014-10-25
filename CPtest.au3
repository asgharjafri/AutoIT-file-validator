#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Cherry Pick testing
; Selects multiplpe files (Maximum 10), Then, one at a time, puts them through QA and displays the results
; Displays the results as difference between Production TIF and QA TIFI
;
; Asghar Jafri, started 25 April 2012
; Select source files to test
; Copy them to QA
; Copy Production TIFI to testing folder
; Copy new QA TIFI to testing folder
; Compare the TIFI files
;
; Possible changes
; Make one multidim ARRAY to make it easier for display and copy
; NOte : This program differs from the later program in that the Array dimensions are reversed. The type of data is first subscript, index is second.
;
; Changes:
; 14 may 2012, To display file name rather than the entire path as the msgbox window is not large enough to comtain the large string.
; 17 May 2012: 	Write results to file
;				Change in comparison to ignore date and col 438 T
; 31 May 2012 : Changes for new EOD QA folder.
; 14 June 2012 : Changed $iMax to flexible
;--------------------------------------------------------------
#include <Array.au3>
#include <Constants.au3>
#include <File.au3>
#include <Date.au3>
#include <_FileFindEx.au3>
;
; Constants
;   maximum size for number of files selectable. Limited to keep a low load on QA

 Local Const $Todaysdate = @YEAR & @MON & @MDAY
 Local Const $SleepInt = 15000 ; Milliseconds to wait bbetween check for presence of file
 Local $MaxWaitTime = 12 * 60 * 1000  ; 15 minutes in milliseconds
 Local $GetTIFIFromBatch = "N"   ; Y : Compare QA with TIFI from ETL Processing,   N: Compare QA with Production Output
 Local $CopyResult = 0
;
; Array declaration setup.
; Use one 2-Dimensional Array to to make it easier to display status

 Local Const $SrcFilesPR          = 1   ; source file to be processed
 Local Const $SrcFilesqa		  = 2
 Local Const $CTLFileNamesPR      = 3
 Local Const $CTLFileNamesQA      = 4
 Local Const $TIFIFileNamesPR     = 5
 Local Const $TIFIFileNamesQA     = 6
 Local Const $TIFISame            = 7; result of comparison
 Local Const $EndOfReport = " ***************    END OF REPORT    **************************"
 ;
 ;
; Predefined Locations
 Local $SourceFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE"
 Local $ProductionTIFIFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\ARCHIVE\2012-05-02"
 Local $ProcessingFolder = "\\etl01-mtl\infa_shared\SrcFiles\EOD_INTACT\PROCESSING"
 ; Local $QADropbox = "\\Fsmon001\ece\IS\Broadridge\EOD Testing\dropbox"   ; old dropbox
 Local $QADropbox = "\\montreal\shares\BPSSEND-QA\eod"                 ; new dropbox
 Local $QAOutput = "\\etl01qa-mtl\infa_shared\SrcFiles\EOD_INTACT\OUTPUT"
 ;  Local $QAProcessed = "\\Fsmon001\ece\IS\Broadridge\EOD Testing\dropbox\processed"
 Local $QAProcessed = "\\montreal\shares\BPSSEND-QA\eod\Processed"   ; new QA
 ; Local $TestFolder = "c:\temp"
 Local $TestFolder = "\\Fsmon001\ece\IS\EOD\Testing"
;
 Local $i  ; increment counter
 Global $Progress = 0
 Local $Progresspct
;
; Get Folder name where testing / comparison will take place
;
$TestFolder = FileSelectFolder("Choose a Test Work folder.","",1+2+4, $TestFolder)
If @error > 0 Then
    Abend("No Test / Work Folder chosen")
Else
;
	; Get filenames for testing
	Local $message = "Choose files to test.Hold down Ctrl or Shift to choose multiple files. "
	Local $var = FileOpenDialog($message, $SourceFolder , "Source Files (*.csv;*.txt)", 2 + 4)

	Local $var1
	; MsgBox(4096,"","You chose " & $var)
	If @error Then
		Abend("No File(s) chosen")
	Else
		; Ask for new identifier.
		Local $identifier = InputBox("Please Enter the New Identifier", "Anythnig other than 2 character will result in no change: ", "", "", _
         - 1, -1, 0, 0)

		ProgressOn ( "Cherry Pick Testing", "Testing ...","",-1,-1,2+16)

		; Prepare file to write results to
		Local $DT = _NowCalcDate() & _NowTime(4)
		$DT = StringReplace($DT,"/","")
		$DT = StringReplace($DT,":","")
		;
		Local $ResultsFileName
		$ResultsFileName = $TestFolder & "\" & "Results" & $DT & ".txt"
		Local $ResultsFile = FileOpen($ResultsFileName,1)

		; Write headings
		Local $OutString = "Program:CPTest.au3, " & $ResultsFileName &  ": All Results" & @CR & @LF & @CR & @LF
		FileWrite($ResultsFile, $OutString)
		;
		; Prepare file to write all errors to
		Local $DiffsFileName
		$DiffsFileName = $TestFolder & "\" & "Diffs" & $DT & ".txt"
		Global $DiffsFile = FileOpen($DiffsFileName,1)
		;
		; Write headings
		Local $OutString = "Program:CPTest.au3, " & $DiffsFileName &  ": List of all Differences" & @CR & @LF & @CR & @LF
		FileWrite($DiffsFile, $OutString)
		;
		; move the flenames into the array SelectedFiles
		Local $VarLen = StringLen($var)
		Local $FirstSep = StringInStr($var,"|")
		if $FirstSep = 0 Then  ; For when only one file is chosen and there are no | seperators
			$FirstSep = StringInStr($var,"\",0,-1)
		EndIf
		$SourceFolder = StringLeft($var,$FirstSep-1)
		Local $JustFilenames = StringTrimLeft($var, $FirstSep)

		; Get Filenames
		Local $FileList = StringSplit($JustFileNames, "|")

 		Local Const $iMax=$FileList[0] + 1
		Local $DataArray[8][$iMax]  ; Array containing the status data;  Dimension 8 means 0 to 7*&*

		For $i = 0 To $FIleList[0]
			$DataArray[$SrcFilesPR][$i] = $FileList[$i]
		Next

		;  _ArrayDisplay($SrcFilesPR, "Source copied file names")
		Local $last = $DataArray[$SrcFilesPR][0]
		;
		; Copy the files to QA
		;-- Move and rename the files so they will run in QA
		;----- Strip the batch time, put in today's date
		;
		Local $src
		Local $dest

		$ProgressPct = 100 / $DataArray[$SrcFilesPR][0]

		For $i = 1 to $DataArray[$SrcFilesPR][0]   ; This element comtains the number of filenames entered

			$Progress = ($i - 1) * $ProgressPct
			; msgbox(0,"Progress",$Progress)
			ProgressSet($Progress, "Processing " & $DataArray[$SrcFilesPR][$i])
			; Format the Filesnames of the files to be dropped into QA by putting today's date into their names
			$DataArray[$SrcFilesQA][$i] = StringMid($DataArray[$SrcFilesPR][$i], 8, 7) & $Todaysdate & StringRight($DataArray[$SrcFilesPR][$i],7)
			; Check if AsOf. Change name if necessary
			; sample : MCQC_AF20120531001.txt
			;          1234567890123456789012
			;                 543210987654321
			if StringMid($DataArray[$SrcFilesQA][$i],6,2) = "AF" Then
				$DataArray[$SrcFilesQA][$i] = StringLeft($DataArray[$SrcFilesQA][$i],5) & "AR" & StringRight($DataArray[$SrcFilesQA][$i],15)
			EndIf
			;
			; If identifier is to be replaced ...
			; sample: MCQC_AF20120531001.txt
			;         1234567890123456789012
		   	;                   1    1 1
			;
			if StringLen($identifier) = 2 Then
				$DataArray[$SrcFilesQA][$i] = StringLeft($DataArray[$SrcFilesQA][$i],15) & $identifier & StringRight($DataArray[$SrcFilesQA][$i],5)
			EndIf

			; Before Submitting a file to QA, make sure it's not a duplicate there.
			If FileExists($QAProcessed & "\" & $DataArray[$SrcFilesqa][$i]) = 1 Then
				$DataArray[$SrcFilesqa][$i] = "Duplicated in QA"
			Else   ; No Duplicate in QA,
				; 	Getting TIFI From Production
				;   Fabricate Control File Name
				$DataArray[$CTLFileNamesPR][$i] = "CTL_" & StringRight($DataArray[$SrcFilesPR][$i],22) & ".XML"

				; Use function GetTifiName to get TIFI Name
				Local $PathNameCTL
				$PathNameCTL = $SourceFolder & "\" & $DataArray[$CTLFileNamesPR][$i]
				$DataArray[$TIFIFileNamesPR][$i] = GetTifiName($PathNameCTL)
				; 	Proceed only if TIFI exists
				If $DataArray[$TIFIFileNamesPR][$i] <> "No TIFI" Then   ; TIFI Name has been found
					; copy the source to testing for documentation under original name
					$src  = $SourceFolder & '\' & $DataArray[$SrcFilesPR][$i]
					$dest = $TestFolder   & "\" & $DataArray[$SrcFilesPR][$i]
					ProgressSet($Progress, "Copying to Test Directory " & $DataArray[$SrcFilesPR][$i])
					$CopyResult = Filecopy($src, $dest,1 ) ; Flag 1 : Overwrite
					If $CopyResult = 0 Then
						Abend("Copy from" & $src & " to " & $dest)
					Endif
					; Copy src to QA for ETL processing with a) the date part of the name changed to today
					;                                       b) and the AF/AR AsOf switch
					;                                       b) and the Identifier
					; src is the Production source
					$dest = $QADropbox & "\" & $DataArray[$SrcFilesqa][$i]
					$CopyResult = Filecopy($src, $dest )  ; QA folder will only accept files with today's date
					If $CopyResult = 0 Then
						Abend("Copy from" & $src & " to " & $dest)
					Endif
					;
					;	Copy the Production TIFI to Test for comparison with QA TIFI
					;  - If checking against TIFIs in the Processing folder then fabricate batches
					If $GetTIFIFromBatch = "Y" Then
						; Processing Folder + Date with Dashes + Batch + TIFIName
						Local $srcdate = StringMid($SrcFilesPR[$i],15,8)
						Local $srcdatedash = StringLeft($srcdate,4) & "-" & StringMid($srcdate,5,2) & "-" & StringRight($srcdate,2)
						Local $SrcBatch = StringLeft($SrcFilesPR[$i],6)
						$src = $ProcessingFolder & "\" & $SrcDateDash & "\" & $SrcBatch & '\' & $DataArray[$TIFIFileNamesPR][$i]
						; MsgBox(4096,"","Processing Path: " & $Src)
					Else
						$src  = $SourceFolder & '\' & $DataArray[$TIFIFileNamesPR][$i]
					EndIf

					; _ArrayDisplay($DataArray, "Test Status")

					; Copy Production TIFI
					ProgressSet($Progress, "Copying Production TIFI " & $src)
					$CopyResult = Filecopy($src, $TestFolder, 1	)
					If $CopyResult = 0 Then
						Abend("Copy from" & $src & " to " & $TestFolder)
					Endif

					;  - QA - Get TIFI Name and location
					; ------- QA may not have processed the file yet so we have to wait for
					; the Control File and the TIFI
						$DataArray[$CTLFileNamesQA][$i] = "CTL_" & StringRight($DataArray[$SrcFilesqa][$i],22) & ".XML"
						$PathNameCTL = $QAProcessed & "\" & $DataArray[$CTLFileNamesQA][$i]

					Local $ok = WaitForFile($PathNameCTL, "</Report>", $MaxWaitTime, $SleepInt)
					if $ok = 0 Then
						MsgBox(0,"Error: No Control File", $PathNameCTL)
					Else
						ProgressSet($Progress, "Getting TIFI name " & $DataArray[$SrcFilesPR][$i])
						$DataArray[$TIFIFileNamesQA][$i] = GetTifiName($PathNameCTL)

						; copy the QA TIFI file to testing
						$src = $QAOutput & '\' & $DataArray[$TIFIFileNamesQA][$i]
						$ok = WaitForFile($src, "REC-CNT=",$MaxWaitTime, $SleepInt)
						if $ok = 0 Then
							MsgBox(0,"Error: No TIFI", $src)
						Else
							ProgressSet($Progress, "Copying QA TIFI " & $Src)
							$CopyResult = FileCopy($src, $TestFolder, 1)
							If $CopyResult = 0 Then
								Abend("Copy from" & $src & " to " & $TestFolder)

							Endif
							;
							; Compare the TIFIs
							; MsgBox(0."Check incs,TIFISame: " & $TIFISame & "i: " & $i & "TIFName " & $TIFIFileNamesPR & " TIFI QA: " $TIFINamesQA

							$DataArray[$TIFISame][$i] = CompTIFI($TestFolder, $DataArray[$TIFIFileNamesPR][$i],$DataArray[$TIFIFileNamesQA][$i])
						EndIf
					EndIf
				EndIf  ; IF No TIFI

				; _ArrayDisplay($DataArray, "Test Status")

			Endif  ; If File does not exist in QA Processed


			$OutString = $DataArray[$SrcFilesPR][$i] & " | " & $DataArray[$TIFIFileNamesPR][$i] & " | " &$DataArray[$TIFIFileNamesQA][$i]
			$OutString = $OutString & " Differences: " & $DataArray[$TIFISame][$i] &@CR & @LF &@CR & @LF
			FileWrite($ResultsFile, $OutString)
			;
		Next

		ProgressOff()

		ProgressSet($Progress, "Writing to FIle: " & $ResultsFileName)
		;
		; _FileWriteFromArray($ResultsFileName, $DataArray, 1)

		; Display results
		; Run("notepad.exe " & $sFile)

		; _ArrayDisplay($DataArray, "Test Status",10,1,"","|","||Source File|QA Source|Production CTL|QA CTL|Production TIFI|QA TIFI|Diff at Rec")

		FileWrite($ResultsFile,    $EndOfReport)
		FileWrite($DiffsFile,      $EndOfReport)

		FileClose($ResultsFile)
		FileClose($DiffsFile)

		MsgBox(0,"Completed."," Results in " & StripPath($ResultsFileName) & " and " & StripPath($DiffsFileName) & " in " & $TestFolder)
		Run("notepad.exe " & $ResultsFileName)
		Run("notepad.exe " & $DiffsFileName)

	Endif ; No Folder Chosen

Endif    ; No files chosen
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
FUNC CompTIFI($TestFolder, $TIFI1, $TIFI2)

	Local $recnum = 0
	Local $numerr = 0
	; Local $Progress = "Comparing..."
	Local $LastRec   ; indicates last record of file (only for regular TIFI files with the string "REC-CNT=" in the last record
	Local $LEOF = "F"
	Local $file1, $file2, $filediff
	Local $line1,$line2  ; records read form file
	Local $cline1, $cline2   ; records processed for comparison

	Local Const $OneBlank    = " "
	;						    123456
	Local Const $SixBlanks   = "      "
	;						    12345678
	Local Const $EightBlanks = "        "
	;						    1234567890
	Local Const $TenBlanks   = "          "

	Local Const $FidessaNames = "ALTA,FIRE,MCQC"

	; Get File Size
	Local $FileSizePR = FileGetSize ($TIFI1)
	Local $FileSizeQA = FileGetSize ($TIFI2)
	if $FileSizePR <> $FileSizeQA Then
		$numerr = " Size Diff"
	Else
		;

		$file1 = FileOpen($TIFI1)
		$file2 = FileOpen($TIFI2)

		Local $File

		; Start loop to check errors in all files
		While $LEOF = "F"

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
			if $recnum = 1  Then  						; first record contais the name of the source file.
				Local $NameCorr = StringMid($line1,191,4)
				Local $NamePos = StringInStr($FidessaNames,$NameCorr)
				$Fidessa = $NamePos > 0
				;
				; Might as well get the name of the source file also
				Local $NameSrc = StringMid($line1,191,22)
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
					$numerr = $numerr + 1
					;
					Local $i
					For $i = 1 to StringLen($cline1)
						Local $c1 = StringMid($cline1,$i,1)
						Local $c2 = StringMid($cline2,$i,1)
						if $c1 <> $c2 Then
							$errpos = $i
						endif

					Next
					; error position has been found
					; Write to the Diffs file
					Local $RecSamplePR = StringMid($CLine1,$errpos - 5, 10)
					Local $RecSampleQA = StringMid($CLine2,$errpos - 5, 10)

					if $numerr = 1 Then  ; write heading for first error only
						Local $OutString = $NameSrc & " | " & StripPath($TIFI1) & " | " & StripPath( $TIFI2 ) & " | Rec: " & $recnum & " | Pos:  " & $errpos
						$OutString = $OutString & " |" & $RecSamplePR & " | " &  $RecSampleQA & "|" &@CR & @LF
					else
						$OutString = "                                                                       " & " | Rec: " & $recnum & " | Pos:  " & $errpos
						$OutString = $OutString & " |" & $RecSamplePR & " | " &  $RecSampleQA & "|" &@CR & @LF
					endif
					FileWrite($DiffsFile, $OutString)

				EndIf
			EndIf

		WEnd

		FileClose($file1)
		FileClose($file2)
		;
	EndIf  ; File Size different
	;
	; Return a 0 if no difference is found
	If $numerr = 0 Then
		Return 0
	Else
		Return $numerr    ;"Rec: " & $recnum & " Pos:  " & $errpos
	EndIf

EndFunc   ;CompTIFI
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
;; Function Abend - Abnormal End. Exit program
FUNC ABend($ErrorStr)

	MsgBox(0,"Fatal Error",$ErrorStr)
	Exit

EndFunc ; StripPath

;~-----------------------------------------------------------------
;;~-----------------------------------------------------------------
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