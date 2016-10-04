Attribute VB_Name = "HeadOfficeStatistics"
'!!! ManualFix Clearing: Problem auto-fix by Project Analyzer 7.1.07 on 15/12/2005
'---------------------------------------------------------------------------------------------------------
' Following procedures could be reinstated in main form as the number of procedures in form
' is reduced and it no longer exceeds the limit for using VBWatch:-
'   subSetRemoteDatabaseOpenedbyField, flNumberOfExistingFranchises, subAddToRemoteEventLog
'---------------------------------------------------------------------------------------------------------
    Option Explicit
    
    'product definition
    Public Const gconUnknownProductName As String = "Data Warehouse"
    
    ' Date     Ver      CR            Who     Description
    '-----------------------------------------------------------------------------------------------------
    ' 01-11-01 B33      SOW-TSG001    PAL     First urgent parts of stage I done now
    '                                           1. reversed dialling sequence (VIC first)
    '                                           2. cut down array from 3 yrs to 1 yr (mem?)
    '                                           3. Auto kick-off Procomm from VB & biff sched
    '                                           4. Ring support if successful upload to BATA
    ' 06-11-01 B34      support       PAL     Use RAS not Procomm to ring support if success
    ' 16-11-01 B34c     support       PAL     Added a loop around the individual dial of one
    '                                         to redress the code-67 problem. Added a new
    '                                         franchiseRetry value into defaults. Also added
    '                                         the error code to the TEST RAS share fail.
    ' 20-11-01 B34nt    support       PAL     Mods for running under NT5:
    '                                         1. added "ts" and "headoffice" to WSNetAddCon.. call
    '                                         2. included GetComputerName & OS ver.
    ' 25-11-01 B341nt  support        PAL     1. New version of pkzipc.exe to avert PKZIP crash
    '                                         2. Modified BATA buttons (upload now works)
    ' 26-11-01 B342nt  support        PAL     1. Timeout 20 minutes after procomm
    '                                         2. Nielsen zip bug (runtime error)
    ' 30-11-01 B40     SOW-TSG001     PAL     1. Multi-file send (includes new procomm script)
    '                                         2. New table BataZipFiles & 'sent' dir
    '                                         3. Cater for both W98 & NT5 here (notepad etc)
    ' 01-12-01 B41     support        PAL     Bug... put upload file into tmpfile
    '                                         Also fixed 'filecount' from zipfile
    ' 03-12-01 B42     SOW-TSG001     TCS     rebuilt front end and added code to hopefully
    '                                         reconcile all bata files sent and do archive table
    ' 07-12-01 B43     SOW-TSG001     PAL     Added guts of upload-to-franchise
    ' 12-12-01 B44                    PAL     All 10 tabs added; guts behind Upload tab;
    '                                         Also modified BATA resend to send minimum
    ' 15-12-01 B46     SOW-TSG001     PAL     Little tidy ups (see list in log) like:
    '                                         - genericize names, etc
    ' 20-12-01 B47h                   PAL     RAS error (no disconnect after 2 coded-67s)
    ' 21-12-01 B50                    PAL     Purge livedata (hangs?..better msgs, etc)
    ' 03-01-02 B50a                   PAL     Stick faster; Sent listview; resend/recreate
    '                                         BATA is better;stick button;execcmd not shell to
    '                                         to avoid Winoldap;
    ' 06-01-02 B54                    PAL     IOEMGA.
    ' 09-01-2  B56                    PAL     Dont allow slave to capture
    '                                         Dial sequence
    '                                         state stuff uses VIC etc not areacode
    '                                         Modem goes in defaults.mdb
    ' 11-01-02 B58                    PAL     rm Trailing spaces on BATA (GlenInnes/Narrabri prob)
    '                                         Move "MasterStatus" to defaults.mdb
    ' 12-01-02 B60                    PAL     More master/slave stuff (bugs, slow, etc)
    ' 13-01-02 B61                    PAL     Improved procomm script & have retry in here
    ' 15-01-02 B62                    PAL     Added Rumpys stock report.
    '                                         Fixed procomm retry bug; utility.exe upload flag
    ' 17-01-02                                Set Rsstarttime; set remote date/time
    ' 18-01-02                                rm end stick rpt; sleep 2s b4 open procommlog
    ' 19-01-02 B64                            no eventlog update for upload; uploadpending listview
    ' 21-01-02 B66                            Added TCS DB merge button at home with Rumpy
    ' 22-01-02 B67                            Polishing MergeDB function at TSG with Andy
    ' 22-01-02 B68                            Don't update querylastrun for manual capture or
    '                                         lastsuccessfulcapture for each franchise, or IOMEGA
    ' 27-01-02 B70                            Debugging timing issues.
    '          B71                            only upload to live franchises
    '          B72                            don't refresh event log to speed things up
    ' 10-02-02 B74                            Slave mods (rec count); no IOMEGA reminder, etc
    ' 11-02-02 B75                            Bug re TSWLP00n.txt filename
    ' 19-02-02 B76                            Summary stick report
    ' 23-02-02 B78                            Revamped Product Report ex TCS; also mods B4
    '                                         Master/Slave commissioning.
    ' 02-03-02 B79                            Lotso little mods: reduce footprint, DB copy button
    '                                         dialup results print; stick print btn enable;
    ' 04-03-02 B80                            IOMEGA is now in defaults.mdb
    ' 06-03-02 B83                            Rejections print.
    ' 14-03-02 B84                            Flash window; rejects to include -ve;
    ' 27-03-02 B85                            Allow uploads to 'Test' machine;
    ' 03-04-02 B86                            Enable ReCreate BATA button; fixed 3265 error on uploads
    '                                         during capture;
    ' 06-04-02 B86FM                          Forked FM branch
    ' 17-04-02 B87                            Finish off faxes for rejects; reports now hook into archive DB
    '                                         if startdate < beginning of live data
    ' 18-04-02 B91                            Reject and flag if livedata > today; show fax on front tab
    '                                         Tidy up PRproductreport...redundant clearing of salesreport stuff
    '                                         redundant functions; Added new 'missing sales' button
    ' 24-04-02 B92                            faxskel const, and NetworkPrinterEnabled, sAction;
    '                                         Completely Reworked updateZipTable function to avoid DOS shell unzipping paradigm.
    ' 25-04-04 B93                            New datasets for procomm file-of-files & 1 days zip.
    '                                         Reworked subBATAUploadWithProcomm
    '                                         gsProcommFileOfFiles created in one spot CreatedFileOfFilesForProcomm
    ' 28-04-04                                Make subCreateZipFiles a function
    '                                         Added new 'missing frans' button
    ' 01-05-02 B94                            Reject all but one from pre-live table if date > today
    ' 06-05-02 B95                            Allow Stick report print
    ' 24-05-02 B96                            Bug in Stick Rpt whereby 'out of range' if no tobacco products
    ' 25-05-02                                Hardcode sharename = 'statistics'
    ' 26-05-02 B97                            Multi-select on product report; FUTURE date warning at end of cycle;
    ' 19-06-02 B98                            New Promotions tab
    ' 26-06-02 B99                            Started faxes for non-compliants
    ' 09-07-02 C00                            More work on Promo-Management and non-compliance
    ' 21-07-02 C02                            Promo manager, etc
    ' 29-07-02 C03                            Upload Promos; also tighten upload to rogues;
    '                                         Fix bug whereby 'No promos' msgbox stops capture
    ' 07-08-02 C05                            Finetune Promo Mgr, rpts, faxes etc;
    '                                         Versions screen; dateformatbad warning.
    '                                         Fixed 'unknown supplier' crash in Stock tab.
    ' 11-08-02 C06                            Promo - tidy up.
    ' 14-08-02 C08                            Promo precision (4-dec places) bug
    ' 17-08-02 C09                            Promo report
    ' 28-08-02 C11                            Promo fax bugs; Slave DB on Network
    ' 04-09-02 C12                            RMver; compact/zip/copy DBs
    ' 11-09-02 C14                            promo bugs (need to merge Laptop mods too)
    ' 18-09-02 C15                            5c promo
    ' 19-09-02 C16                            5c naildown
    ' 22-09-02 C18                            zip/unzip; DBs on Network;
    ' 27-09-02 C19                            delete yesterday-1 when load promo tab; doevents
    ' 28-09-02 C20                            Upload promo bug; stop timer when zip DB
    ' 29-09-02 C21                            3078 err.clear upload.promo
    ' 02-10-02 C22                            dialup-results display bug; auto upload upgrades.
    ' 09-10-02 C23                            nielsen weekly change;
    ' 09-10-02 C24                            undo Nielsen change; slaves dont need procomm
    ' 23-10-02 C25                            OSVER bug; slave OK w/o archiveDB; new-promo rpt tidy;
    '                                         delete *.mdb from IOMEGA b4 copy; force portrait on rej-fax
    ' 30-10-02 C27                            quote path (SFG WORK);
    ' 07-11-02 C28                            quote path (SFG WORK) properly.
    ' 13-11-02 C29                            unzip locally & then copy to server
    ' 20-11-02 C30                            Trap error 3045 (DB in use)
    ' 26-11-02 C31                            master starts DCCserver at end-of-cycle then Slave checks master & sends email;
    '                                         BATA sent bugs (only add entry if filecnt > 0;
    '                                         BATA sent bugs (add sent files correctly (find prevVersion))
    ' 30-11-02 C34  Biff DCC idea. Use DUN instead.
    ' 12-01-03 C36  WRS. (dont forget to add email constants from slave to avoid server error, also rm ping)
    ' 23-01-03 C37  WRS Don't use spare1 as RM ver use 'version'.
    ' 28-01-03 C38  Add trailer to BATA files as per spec from Dexter. temporarydata table: add TSGId & change date to date/time
    '               Use spare1 as DB version and 'version' as RS version
    ' 29-01-03 C39  Start adding WHS to reports
    ' 03-02-03 C42 continue WHS reports; fix Dexter bug whereby whs file had total sell not whs sell.
    ' 05-02-03 C43 Dialin to fix 'headoffice'flag in remote defaults; also send faxes auto
    ' 05-02-03 C44 slave
    ' 12-02-03 C45 Rejected fax for future dates. Allow capture of selected stores
    ' 1-03-03  C46 Polish up WRS (promo, report print etc)
    ' 05-03-03 C47 Don't msgbox if eventlogerror; network path to rpt files; selective delete uploads pending;
    ' 12-04-03 C50 Datestamp rejects; Fix bug when rejecting futuredate txns (do all not just first);
    '              Add cancel to sales rpt; show txnstartdate on settings tab; dont store non-tobacco rejects (fax complaint)
    '              Fix bug picking up OSVer & RMVer from capture cycle.
    ' 13-04-03 C51 Don't msg box if other users logged into DB.
    ' 16-04-03 C52 Show current live version on Versions tab; Also flag 'y' for utility.exe upload.
    ' 23-04-03 C54 Kill connection and continue if 'RAS Connection already exists' during capture cycle.
    ' 07-05-03 C55 version tab - dont show RMver on Towers etc; don't delete old promotions (flag as deleted)
    '              Also allow upload of "current session" and ignore the pending ones
    ' 14-05-03 C56 Two bug fixes: loadNonCompliants ex Andy & PRORlandscape crashes promo rptprint ex Estee
    ' 17-05-03 C57 Implement dynamic RAS password for new stores from now on based on fran name and ID (TechRentals supplier); I'm backin away so don't ring me;
    ' 28-05-03 C59 Fixed bug whereby new promos not showing in top window
    ' 14-06-03 C60 fix bug Be tolerant of NULL in whsQty in loadNonCompl
    ' 25-06-03 C61 Fixed bug causing promo deletion to crash
    ' 02-07-03 C62 Warning if stores offline for > 4 days; percent success; less cluttered 'eventlog'
    '              Dont do 'uploadspending' for a test RAS if we are a slave
    ' 10-07-03 C63 Upload 'promo message' checkbox
    ' 27-07-03 C65 Export stock button (from either main DB or nominated DB)
    ' 29-10-03 C65a Added on error resume next to avoid promo crash (null whsqty whstotal etc)
    ' 04-12-03 c65b Special fix to delete promo1246 (from VIC stores)
    ' 28-12-03 c71  import promos from csv; Email Nielsen report from SLAVE
    ' 06-01-04 c72 Added reference to RAS_AUTO.dll (VB6 version)
    ' 12-01-04 c74 nielsen email from one slave; Introduce idea of PLEB
    ' 28-01-04 c74b rm exit sub from Nielsen
    ' 28-01-04 c75 Biff Nielsen email temp.
    ' 25-02-04 c76 Fix product report bug (error 450) ex Neil/Andy
    ' 18Mar04 c77  Remove redundant code in error handlers in subCaptureData
    '              Close recordsets in error handlers in subCaptureData with a wrapper function
    '              Wrapper function ignores errors and therefore avoids errors occuring in an active error handler
    '              Calling code has no error handlers hence when errors are promoted form here they become runtime errors
    ' 18Mar04 c78  Log emailing from Server (even if overwritten when Master data is transferred)
    '              Replace generic object declaration of oRASConnection with 'Dim oRASConnection As RAS.Connection'
    ' 18Mar04 c79  Completed reinstating emailing of overnight log.
    '              Required realigning where NotifySlave() wrote data and SlaveCheckingStatusOfMaster expected it.
    '              Required configuration including enabling emailing and creating and sharing a folder on slave
    ' 05May04 c80  Add Err number, desc & line to logging in error handling code of subCaptureData (improve diagnostics)
    '              Change icon of main form to that used for program shortcut
    '              Added index DateTimeDescending to DateTime field of Event Log in TSGDataWarehouse.mdb
    '              Added reference to Scripting.FileSystemObject to cater for customisation of VB Watch Error handling
    '              Added extra code module to assist in gradual movement toward modular code
    ' 03Jun04 c81a Started cleaning/re-writing frmTSGDataWarehouse.cmdPRPrint_Click() (ie Report button on 'Product Report' Tab)
    '              SQL for refreshing event log display optimised [25% speed increase] - see gsubRefreshEventLogDisplay()
    '              fsErrDetail re-written because VB Watch error handling reinitialised error parameter (On Error ...)
    ' 10Jun04 c82 RAS object declaration readied code inclusion rather than dll inclusion (is commented)
    '             VB Watch NoErrorHandler Tag added to Function gfsRASErrorMessageFromErrorCode() so that
    '             this procedure does not have to be manually excluded from Error Handling when using VB Watch
    '             EXACT VB WATCH TAGS CANNOT BE IN MODULE COMMENTS BECAUSE THEY WOULD APPLY TO WHOLE MODULE
    '             Remove redundant commented out procedures
    '             Comment out emailing of Nielsen report (Phil was working on it but it wasn't completed)
    '             Add temporary code to exclude dialling BATA between "12 Jun 2004 00:01" and "12 Jun 2004 09:00"
    ' 17Jun04 c83 VB Watch NoTraceProc Tag added to NoErrorHandler Tag in gfsRASErrorMessageFromErrorCode()
    '             becuase procedures called by tracing have their own Err Handling
    '             EXACT VB WATCH TAGS CANNOT BE IN MODULE COMMENTS BECAUSE THEY WOULD APPLY TO WHOLE MODULE
    ' 01Jul04 c84 Extend commenting out emailing of Nielsen report to NOT create Nielsen report on SLAVE
    '             machine. (Phil was working on it but it wasn't completed) On occsions program crashed
    '             trying to acces network drive and consequently the overnight email could not be sent
    '             Comment out all gconProductionVersion conditional code as part of removing clutter
    '             Alter Error Handlers that call subIgnoreErr_CloseRst() so they remember error
    '             details for reporting before subIgnoreErr_CloseRst() clears the details
    ' 15Jul04 c85 Changed date format in overnight email subject line for better sorting
    '             Made program start on Product Report tab and NOT unconditionally refresh Event Log which would prevent
    '             SLAVE starting during high network activity (eg backups) Required numerous linked changes.
    '             Commented out fsNielsenZipFullPathAndFilename() procedure which became redundant when I commented
    '             out incomplete work by Phil on emailing of Nielsen report on SLAVE machine
    '             Comment out redundant error handling code and flag with VC85 (Err.Clear preceding On Err statement, etc)
    '             General code cleanup (OpenSession &  SendEmail procedures)
    ' 22Jul04 c86 Removed 'INTELligeNT Technology' from Company Name of project file and therefore executable
    '             Added module wide "Enum enmTab" to frmDataWarehouse to replace duplicated local constants
    '             Helper routine created for subTabMainClick and routine called in frmLoad to ensure the tab
    '             the program starts on is correctly initialised. (Startup tab was chnaged in vers c85)
    '             Removed two week delay for Nielsen Report(gconNielsenDateDifferential = -14)
    '             Maintained utility code [Added SetMousePointer(), Removed MSSQLDateTime()]
    '             Reorganised commented out redundant code
    '             Wrote a series of NEW Nielsen report functions with parametised date periods to replace old
    '             functions that relied on control values and would only produce weekly reports - and added code to test
    '             {subRefreshNielsenReportListBox(), fsNielsenFileSpecificationNEWPROC(), CreateNielsenRptNEWPROC() }
    '             {subPurgeNielsenReportsNEWPROC(), fsNielsenRptFullnameNEWPROC() }
    ' 26Aug04 C87 Removed commented out code in CreateNielsenRptNEWPROC, and changed its
    '             where clause to use equality test (cf BETWEEN) when StartDate = EndDate
    '             Cleaned up error handling in subCaptureData including the fsErrDetail function
    '             Removed VB Watch tags from module comments because
    '             EXACT VB WATCH TAGS CANNOT BE IN MODULE COMMENTS BECAUSE THEY WOULD APPLY TO WHOLE MODULE
    '             Renamed gsubAddMessageToStatusBar to StatusBar and made its last 2 parameters optional
    '             (Start drive toward more concise naming)
    '             Added DoubleQuote and SingleQuote functions for readability
    '             Replaced code in subCaptureData that created Nielesn report (2 weeks in arrears) with code calling
    '             my new routines to created 3 separate Nielsen reports for each of the 3 previous weeks.
    '             Created new ZipNielsenReports function that takes a date parameter and zips the corresponding report
    ' 29Aug04 C88 Made some fixes to new code fragment in subCaptureData used to create Nielesn Reports
    ' 02Sep04 C89 Update version no. (overlooked in last version)
    '             Copy Nielsen Zipped report files to sub-folder rather than text files
    '             & log  "Nielsen report cycle completed"
    ' 09Sep04 C90 Changed gsubRefreshEventLogDisplay to retrieve data using new lngDate field in EventLog table
    '             Changed TestButton to be conditional on TestMode command argument
    '             Cleaned up loads of code that I have commented out since arriving. Basically a cleanup release
    ' 16Sep04 C91 Added daily reports to Nielesen report cycle. (produces daily reports for each day in 3 reported weeks)
    '             Commented out some unused constants
    ' 23Sep04 C92 Added function ZipFiles() and call it from CreateNielsenReports
    '             Added functions:  GetFilePath(), GetFileName(), GetFileExtension()
    ' 30Sep04 C93 Fixed runtime error when sending Product Report (Summarised) to file
    ' 07Oct04 C94 Compiler switch set to optimise for small code b/c master has memory problems when Explorer is running
    '             Replace subPurgeNielsenReportsNEWPROC with more specific subPurgeNielsenReport()
    '             Fixed a bug where a Nielsen report being created as part of a batch of reports was accidentally purged
    ' 28Oct04 C95 Renamed subRefreshNielsenReportListBoxNEWPROC to subRefreshNielsenReportListBox and replaced all calls
    '             to subRefreshNielsenReportListBoxOLD with calls to subRefreshNielsenReportListBox which was removed
    '             Changes required new simple procedure fdatNielsenWeekEndingDate()
    '             Numerous code cleanups including slowly reducing code conditional on different OSs
    '             Reduced number of controls on form by adding frames not referenced in code to Frame control array
    '             Added controls for displaying Quatro details. Have not enabled editing yet, but have discovered
    '             that other controls on the form are enabled for editing by double clicking them.
    ' 04Nov04 C96 Completed removing code conditional on OS verion. Included removing SysInfo control and reference
    '             Removed code from subCaptureData and put it in a new helper routine subCaptureDataReconcileTaskLog
    '             in attempt to begin the divide and conquer of subCaptureData
    '             Changed FranchiseIncludedInStatistics flag in franchises table from a number to a boolean and
    '             made associated code changes. Various other code clean ups.
    '             Added some stubs in code for moving toward Non Rm downloading
    ' 11Nov04 C97 Minor changes: Changed some if statements for readability only
    ' 09Dec04 C98 Remove code commented out in VC96
    '             Minor changes: Changed some if statements for readability only, removed unneccessary optional params
    ' 16Dec04 C99 New module (ImportBataReport) starte for commencement of accepting BATA reports
    '             Added editing of Quatro fields, Franchise Type field as a combo box
    ' 23Dec04 301 Removed gconBetaVersionNumber constant and moved to pure VB versioning (useful with VB Watch)
    '             Catered for missing entry in FranTypes table for FranchiseType value in Franchises table
    '             Minor code cleanup of commented code
    ' 06Jan05 302 Removed "RASAUTO VB6 Sep03" (C:\WINDOWS\SYSTEM\ras_auto.dll) from references and replaced the dll
    '             with source code directly in the project - this had to be done sometime so that it could
    '             be debugged and also altered for other operating systems etc.
    '             All new RAS modules have VB Watch tags in module headers to prevent VB Watch instrumenting the
    '             code with error handling.
    ' 13Jan05 303 Fix run-time error when adding a new franchise (bug introduced with addition of FranchiseType combo)
    ' 23Jan05 304 subOpenFileUsingExternalViewer altered to open mdb but NOT exe files
    ' (Sunday)    Function used in places like double clicking on files in upload files list
    '             (Files opened are txt, zip, and mdb. Program lets Windows decide what to open programs with)
    '             Commented out whole module: ImportBataReport. Data realignment project was put on hold 13Jan2005
    '             Comment out RAS code not used by program (except for RAS_GLB.BAS - may get to later)
    '             Looking at following problem I am going to do all I can to reduce memory usage.
    '
    '             Nielsen reports not produced Sunday or Monday night (16/1/05, 17/1/05)
    '             Problem appears to be that Zip didn't work. It didn't work for backing up the database 16/1/05
    '             Worked fine when I tested things after a reboot.
    '             I suspect a memory leak and suspect that master PC was not rebooted Monday (16/1/05)
    '
    '             Increase number of retries for BATA upload from 3 to 5. Change wait between attempts
    '             from 1 minute to start at 30 seconds and increase by 30 seconds with each attempt
    '             (Numerous minor clean ups - program is still in need of clean up) Service Packs applied to SLAVE
    ' 03Feb05 305 Clean up various commented out code. Remove NON-RM-DOWNLOADS flagged code which was being introduced
    '             as part of some development which has now been canned. Remove ImportBataReport.bas from project (part of same canned development)
    '             UploadFilesToOneFranchise(): declare variables intended as strings but inadverntenlty variants to be stirngs
    '             Add numerous EventLog entries in CreateNielsenReports() and immediately prior to and
    '             subsequent to calling it in attempt to find the problem in the said procedure
    '             Add EventLog entries to ZipFiles procedure. Add a 30 second sleep/wait at the end of ZipFiles procedure
    ' 10Feb05 306 SafeDivide function created as a (quick) fix for runtime error in Sales Report [cmdAllItems_Click()]
    ' 14Feb05 307 Rewrite ExecCmd: Remove NoWait flag and include optional timeout parameter (default of INFINITE)
    '             so that any timeout period results from a conscious decision (cf defaulting to 20 mins when NoWait=False)
    '             Modify ExecCmd calls while retaining same functionality.
    '             Use ExecCmd to call external Pkzip program in ZipFiles procedure (cf asynchronous 'Shell' to program)
    '             Change manual creation of new Nielsen reports to use new Nielsen reporting (eg daily rpts etc)
    '             Remove code that applied to old system of Nielsen reporting.
    '             Add more temporary EventLog entries to diagnose why zipping of Nielsen DailyReports is failing
    ' 24Feb05 308 Add DatePicker control wrapped in a user control to replace label and spin control for EventLog
    '             date on data capture form. Remove label, spin control and assoicated code
    '             Renamed CreateNielsenRptNEWPROC, fsNielsenRptFilenameNEWPROC, fsNielsenRptFullnameNEWPROC
    '             Add CtlMove, CtlMoveNext and CtlMovePrevious utility functions (use them user control wrapping DatePicker)
    '             Add optional param (pBriefDelay) to StatusBar() so 800 millisec delay not performed by default
    '             Increase retries on BATA upload from 5 to 8 with wait intervals ranging from 15 secs to 8 mins
    '             giving a maximum wait time of 31.75 mins if all attempts fail.
    ' 27Feb05 309 Fix problem introduced with DatePicker on DataCapture tab. MaxDate property of that DatePicker
    '             is now set in tmrCaptureData_Timer() so it doesn't prevent roll over of dates
' 03Mar05 3.0.9002
'   Modified fdtmLastNielsenRptStart() so the Last(Lateest) Nielsen reporting week (Monday to Sunday)
'   is at least a week behind current date.
'   Modified code to cater whatever field size is set to for Comment field in tasklog table
'   (and in database reduced the field size from 200 chars to 100 chars
' 10Mar05 3.0.9003
'   Reinstate a delay when logging to EventLog table, BUT move it from calling procedure StatusBar
'   to gsubAddRecordToLocalEventLog which sometimes called directly, AND make the delay only when
'   required (multiple events on the same second) rather than non-conditionally waiting 0.8 seconds
'   on every EventLogging (~ 3000 Log entries per night)
'   Comnent out diagnostic event logging for Nielsen Reporting and in ZipFiles() function
' 14Apr05 3.0.9004
'   Add event logging for commencement of Nielsen report cycle
'   Remove commented out code
'   Rename user control based on DatePicker FROM ctlDatePicker TO TDatePicker
'   Improve "Wait for " & n & " minutes before retry." message when Uploading to BATA
'   Removed dead procedures from frmTSGDataWarehouse:
'       (subIgnoreErrorsAddMsgToStatusBar, GetMasterStatus, fsNielsenFirstReportDayDate, fdTaxRateFromTaxCode,
'        fsBATAFileSpecification, subCreateNewUploadScript, sendMessageToSupport, killProgram)
'   Removed Global Const CANNOT_FIND_MASTER = "Cannot find Master drive"   (dead constant)
' 18Apr05 3.0.9005
'   Removed Const gconAlsoAddToLocalEventLog and all references to it in ongoing war against bloatware
'   Miscelleanous tidy ups while in code
'   Renamed fsNielsenFileSpecificationNEWPROC to fsNielsenFileSpecification
' 26Apr05 3.0.9006
'   Add diagnostic code for when program bombs out during Creating/Zipping Mielsen files and Windows reports "Not
'   enough memory to satisfy the conventional memory requirements for this program" (Access97/DA0/Win 98 memory leaks?)
'   Revert Nielsen reporting to report until last Sunday (Gary Courtney request) & more code cleanups of this process
' 28Apr05 3.0.9007
'   Changed TransferSalesRecord() to accept recordset parameters ByRef instead of ByVal
'   (Done for speed and because it is the obvious intention)
'   Renamed fdtmLastNielsenRptDay to fdtmLastSunday
'   Commenced replacing date controls on Promotions Tab (labels with buddy spin controls) with DatePicker
'   controls as there is a bug in the validation of these controls which causes are runtime error
' 05May05 3.0.9008
'   Completed replacing date controls on Promotions Tab with DatePicker controls
'   Added more debug code in the hunt for problem with program running out of memory (Windows reports the problem)
'   when attempting to Zip Nielsen files on Mondays during an automatically invoked Nielsen Report cycle.
'   Increased timeout on ZipFiles function from 20 minutes to 40 minutes
'   Shifted order of event logging in subCaptureData error handlers (RemoteLiveDataTableInitialisationFailed &
'   DataTransferWasInterrupted) so error information in Err object is preserved for reporting event log
' 12May05 3.0.9009
'   Added OnErrorResumeNext error handler in subCaptureData to control code under On Error Resume Next
'   so it could provide ErrorLog info about what errors were handled by this sledge hammer approach in the
'   most critical routine and that it could slowly be replaced by error handling which handled specific
'   errors without the risk of shadowing others.
'   Added 'AUrbanUpToHere flags to continue work on extremely nasty Port Open problem investigated this week
'   Very numerous tidy ups while performing analysis on code
' 19May05 3.0.9010
'   Added WNetError() to report errors when disconnecting from a network resource using WNetCancelConnection2
'   Errors previously reported as "Auto network share connection terminated normally" which hampers problem diagnosis
'   particularly on erratic "Port is already open problem". Errors are reported in the Capture cycle but are not yet
'   reported for Testing Ras. When there are errors disconnecting the program will now make three attempts
'   at disconnecting the network resource using WNetCancelConnection2
'   Disconnecting RAS using fDisconnectRAS also misreported errors with "RAS connection terminated normally"
'   fDisconnectRAS has been changed to at least report the error code when there is an error disconnecting RAS
'   It will require some thorough investigation of the RAS code to accurately and robustly report the RAS errors
' 26May05 3.0.9011
'   Added Public Type udtError as a data structure for storing error details
'   Added fGetErrorUdt() to return a udtError when an error handler is activated (or within On Error Resume Next)
'   Replace gconEventTableEventFieldWidth, gconEventTableFranchiseFieldWidth with dynamically inspecting field size
'   Added subIgnoreErr_CloseDbAndSetToNothing() in attempt to solve problem of always disconnecting network drives with an error condition
'   Passed fForce parameter as True to WNetCancelConnection2 when disconnecting Network share drives
' 02Jun2005 3.0.9012
'   Added Public Enum WNetErrEnum to isolate constants used only for WNetErr() function
'   Remove retries on disconnecting/terminating network share because it was having no effect
'   Fix code in subCaptureData which raised an error when no error condition existed.
'   Did not cause problems but filled EventLog with a non-existent error for every franchise
' 16Jun2005 3.0.9013
'   Use fGetErrorUdt() function in subCaptureData (1st active use of this procedure)
'   Replaced spin controls on BATA Report tab with DatePicker control (& associated code changes)
'   Begun replacing spin controls on Nielsen Report tab
' 30Jun2005 3.0.914
'   Added some CnvNulls calls in subCaptureData to ensure that WholesaleQty & WholesaleActualSell are not
'   downloaded as Null - is only a problem at "Beenleigh Market Place"
'   Added global constants gconStkDescUpgradePrefix = "TSStkDesc" & gconRemoteDefaultsTableStkDescUpdateFlag = "StkDescUpdateFlag"
'   in preparation for code for downloading stk description only updates for all products for updating via Price Module
'   Added Sub AddDAOFieldToTable() & Function IsDAOFieldExists() in preparation for same update
' 07Jul05 3.0.9015
'   SlavePhoneNum field added to TSGDataWarehouse.mdb default table to replace hard coding.
'   Neil can now move the PCs to the new phone lines at his own convenience.
'   Some minor clean ups
' 14Jul05 3.0.9016
'   Add code for downloading Stk Description Update files (Prefix of gconStkDescUpgradePrefix)
'   File uses same flag and directories as TSWLP (Wholesale List Price) updates
' 11Aug05 3.0.9017
'   Add memory logging to see if there is a memory leak (use CreateObject and GetSystemInfo function from VB Watch dll)
'   rather than add another dependency
' 18Aug05 3.0.9018
'   Replaced spin32/label combination of controls on Nielsen Report tab with UpDown/DatePicker control combination
'   (=> Great reduction in code and further toward removing all instances of Spin32 control)
'   Removed some dead procedures and dead constants
' 02Sep05 3.0.9019
'   Replaced reference to DAO350.dll with reference to DAO360.dll
'   Was a problem on Windows XP (Andrew's PC - AW) which has a copy of later dll but not a redundant copy of earlier dll.
'   Added function IncrementStkFileNum() to increment StkFileNum in StkFileNums table as it is only incremented
'   in addToUpdateFile() if it is Read-only. This will not be the case if people have been manually editing the
'   files in which case the file might get stuck on a number. Also re-wrote incrementing of StkFileNum in StkFileNums
'   table in addToUpdateFile() to make it more robust and avoid a bug of attempting to open a read-only file for Append.
' 15Sep05 3.0.9020
'   Renamed subOpenFileUsingExternalViewer() to subOpenFile() and rewrote to work across all Windows versions.
'   [NB 'Shell Start [app] [file]' does not work for Windows XP.]
'   Changed tmrCaptureData_Timer() to skip BATA Uploads on Saturday night (ultimately plan to skip Fridays as well)
'   Highlight event log for Franchise selected in list box on Data Capture form [HighlightSelectedFranchiseInEventLog()]
' 22Sep05 3.0.9021
'   HighlightSelectedFranchiseInEventLog() altered to UN-select any highlighted event log entry when newly selected
'   franchise has no items in event log, also changed to leave focus unaltered (rather than set to lvwEventLog)
'   so that user can select items in franchise list box via keyboard.
' 29Sep05 3.0.9022
'   Changed tmrCaptureData_Timer() to skip BATA Uploads on Friday and Saturday night
' 11Oct05 3.0.9023
'   Changes for uploading Stock field update files (TSSFU*.txt) (Prefix of gconUpdateStkFldsUpdatePrefix)
'   Generic code replaces code for downloading Stk Description Update files (Prefix of gconStkDescUpgradePrefix)
'   which only applied to the one Stk Field (ie description)
'   Stock field update files are comma separated value files with the first row containing the field names
'   The files must include the barcode so they can be matched with stock in a retail manager database
' 05Dec05 3.0.9024
'   Minor enhancements and cleanup of fsErrDetail()
'   Shortened return string of fsVersion from "Version Blah" to "V-Blah"
'   Reduced detail returned by MemoryString() to fit in with reduced size of event field
'   Fixed bug in StripBracketedSubStrings() not catering for right bracket being the end of the string
'   Commented out logging of memory status (i.e. '  LogMemoryUse)
'   Reduced size of messages logged to event log in error handlers in subCaptureData
' 12Dec05 3.0.9025
'   Removed all triple commented out code (i.e. ''' Blah, blah)
'   Commented out dead code from Non RAS Modules with triple commenting
'   (Dead code = Dead procedures, dead variables, variables only written, dead procedure parameters, unused controls)
'   Rewrote memory logging so it can be commented in and out of code at one line -> easy to turn on for diagnosis
' 04Jan05 3.0.9026
'   Reinstated fsVersion to return "Version Blah" instead of "V-Blah" but changed fsErrDetail to ensure it used
'   an abbreviated version
'   Changed scope of Sub subCaptureData to Private
'   Changed msg in MsgBox for no entry in event log
'   Removed from code and interface: Quatro fields, FranchisePostalAddress
'   Changed labels from Physical Address to Street, and Postal Suburb to Suburb
'   Rearranged controls on Data Capture and Sales Reports tabs

'   Marked 'Clearing' and 'Fixed File Number' problems in frmTSGDataWarehouse form and HeadOfficeStatistics module.
'   Marked all problems in subCaptureData(). Addressed each problem and some incidental rewriting while addressing the problems.
'   Conditional compilation of new code in flNumberOfExistingFranchises() which creates new values for franchises.FranchiseIDTSG.
'   The code is conditional as it will rely on changing the field to AutoNumber in the live databases.
'   Checked all uses of flNumberOfExistingFranchises() and ensured there will be no side effects.
'   Removed automatic creation of RAS password for new franchises as it used flNumberOfExistingFranchises() and is an unnecessary complication.
'   Simplified setting of lFranchisesNotIncludedThisDataCaptureCycle in subCaptureData remaining true to the logic which
'   was revealed as questionable (now that it is legible) and will be revisited.
'   Removed newly dead code.
'   Renamed Franchise.FranchisePostalSuburbAndPostcode to FranchiseSuburb and changed name of associated controls.
'   Combined LockDCTabFranchiseCtls and subUnLockTextBoxes in to single LockDCTabFranchiseCtls routine.
' 10Jan06 3.0.9027
'   Changed fsVersion() to cater for intermittent "Run-time error 326 - resource with identifier version not found."
'   Removed triple commented variable declarations.
'   Addressed 'Fixed File Number' in all code except cmdPRPrint_Click()which requires a rewrite (bug prone/bad practice)
'   Removed some commented out code and added comments explaining code and highlighting code requiring attention.
'   Added combo box for State field of Franchise (=> only valid data) and made the field mandatory.
'   Numerous tidy ups of controls on Data Capture tab.
'   Removed conditional compilation code and constant
'   Remove all uses of gsProductName and replace with App.Title (or remove where MsgBox would default to same value) + associated code cleanups
'   Added various comments (eg Flagged where different order b/w If and Then clauses would improve readability
'   Fixed "bug!" in subCaptureDataReconcileTaskLog() that occured when flNumberOfExistingFranchises() was fixed
'   to return correct value and code that previously worked by chance and by not deleting franchise records now
'   no longer worked for Franchises with FranchiseIDTSG values greater than the total number of franchises.
' 18Jan06 3.0.9028
'   Extensive rewriting of subCaptureData (& minor of subCaptureDataReconcileTaskLog) to simplify, remove
'   redundant code, fix errors and prepare for coexisting dialup and VPN access to Franchises.
'   qryTaskLog added to TsgDatawarehouse.mdb Forms basis for migrating to sensible date data types
'   rather than strings & provides basis for purging this table
'   Rewrite of subCaptureDataReconcileTaskLog() to use this query
'   Replaced gconTaskTableTaskSuccessfullyCompletedField with "TaskSuccessfullyCompleted" (started then thought I may as well finish)
'   Numerous utility functions added and used [eg GetDateFrom_ddmmmyy(), IsDateFmtOk() used in Form_Load of main form, etc]
'   [For readability] Replaced DBEngine.Workspaces(0).OpenDatabase with OpenDatabase
'   [For readability] Renamed gdbsTSGDataWarehouse to dbDW and placed in a global udt (ie g)
' 19Jan06 3.0.9029
'   Fixed bug in cmdAllItems_Click() introduced when removing fixed file numbers
'   (File number variable was not prefixed with # when it should have been)
'   Rewriting/streamlining of subCaptureData (massive shortening of variable names in the procedure)
'   Remove all references to FranchiseNumberOfTradingDays in franchise table as it was only half
'   implemented and was adding complexity for no added functionality.
' 23Jan06 3.0.9030
'   Rewrote subEditCurrentTaskLogRecord() to accept TaskLog recordset as a parameter & => eliminated
'   need for a global recordset for the TaskLog. Renamed subEditCurrentTaskLogRecord() to subEditCurrentTaskRecord()
'   Above changes forced associated changes to fConnectFranchiseMapShareDisk which had a parameter change and internal
'   changes yet remains functionally equivalent.
'   Removed grstTaskLog, gconTaskWasCompletedSuccessfully, gconTaskWasNotCompletedSuccessfully,
'   Minor changes to UI (Data Capture tab)
' 30 Jan06 3.0.9031
'   Replaced 'Global Const' declarations with 'Public Const'
'   Wrote generic subPurgeTable() to replace subPurgeTemporaryDataTable().
'   Replaced original procedure and replaced some inline code with new procedure. (New procedure is quicker and simpler)
'   Made Public procedures fGetLastWord, GetFilePath, gfsSplitDate & fsYesterdaysDate private and moved them from main
'   module to main form. Renamed procedures, cleared object variables etc.
' 09Feb06 3.0.9032
'   Removed triple commented code
'   Numerous changes to cmdPRPrint_Click() [Product Report] to fix runtime error when perofming a Summarised report
'   for selected products and one of the product has no sales
'   Bug fixed by changing code for sizing the Results array
'   (nb number of products selected does not always equal number of products with sales)
'   Changed logic of when and how routine asked whether you wanted to overwrite the file
'   Some information being printed was not making it to the file since the file was closed and then re-opened
'   Removed redundant time consuming loop that had no effect on processing
'   Rearranged and reindented code to improve readability
' 15Feb06 3.0.9033
'   Simplified and removed redundant code from subCaptureData
'    Removed OnErrorResumeNext error handler in subCaptureData. Restricted scope of error handler with "On Error GoTo 0"
'    Left only the inline error handling following the 'On Error Resume Next'
'    No events in error log relating to the 'OnErrorResumeNext' error handler (log spans July 2004 to Feb 2006)
'    Removed check for existing RAS connections at start of subCaptureData. (again no entries in event log)
'    Cleared rstSummary in subCaptureData (last uncleared variable in the procedure)
'   Reinstated non-conditional refresh of event log on program startup and set starting tab to Data Capture
'   Wrapped gsubRefreshEventLogDisplay() with Hourglass cursor
' 08Mar06 3.0.9034
'   Created new Network module to hold new and reliable network functions
'   Rewrote WNetError as NetError and greatly improved the function. Moved it to Network module
'   Wrote NetConnectShare, NetDisconnectShare, NetError and NetGetLastExtendedError used in NetError to get the last Network specific (Extended) error
'   Rewrote cmdTestConnection to use new network functions to give them a test run.
'   Moved global ArchiveDataWarehouse.mdb database variable to global udt and opened db variable exclusively
'   Moved TsgDataWarehouse.mdb defaults table variable to global udt
'   Locked and disabled the 'Keep this many days of LiveData' field on the Settings tab because it doesn't work properly.
'   You could change it's value but not save it to Db which is what the program was using
'   Rewrote subTransferAllPreLiveDataToTheLiveDataTable so that status bar messages were not misleading (had a problem)
'   so that variables were cleared, so my own table purge routine was used, etc. etc.
'   Moved subPurgeTable to main module and made it public so it could be called from subTransferAllPreLiveDataToTheLiveDataTable
'   Renamed various variables to less verbose names (eg rstDynDataWarehousePreLiveData TO rstDWPreLiveData)
'   Moved manual Capture Data button to decrease chances of accidentally clicking it, as well as adding an
'   "are you sure?" confirmation to the button for cases where it was accidentally clicked
'   Assorted rewrites to simplify and clarify code for readability/maintainability/extensibility etc
' 09Mar06 3.0.9035
'   Remove triple commented code
'   Add facility for logging long event log entries across multiple records rather than truncating them
'   Log long network extended error messages across multiple records rather than truncating them
'   Rename Network module to Network_WNet to indicate it uses the WNet windows functions
' 10Mar06 3.0.9036
'   Fixed NetGetLastExtendedError to correctly trim Null terminated strings returned by WNetGetLastError
'   Changed calls to StatusBar using pLogWholeMsg:=True because it slowed system down
'   System probably not as slow now that the return string is properly truncated and therefore splitting
'   over so many records is not required, but changes will be implemented incrementally over a few days.
'   Noted required changes to fix error reported when disconnecting shared drive
'   (ie change parameters in call to UnmapShareDiskDisconnectFranchise in subCaptureData)
'   Added new error condition to NetError
' 13Mar06 3.0.9037
'   Reinstate logging long error messages returned from calls to NetError in UnmapShareDiskDisconnectFranchise
'   and fConnectToFranchiseover over a number of event log records if required.
'   Minor Changes: Triple commented some dead code, type declaration for some procedures, etc
' 14Mar06 3.0.9038 (Version created 13Mar06 in advance of 13Mar06 deployment
'   Removed trailing backslash from sRemotePathName parameter in call to  UnmapShareDiskDisconnectFranchise
'   from subCaptureData. (caused an un-reported network error that was reported as succesful disconnection by hardcoding)
'   Switched order of 'if then else' statements in subCaptureData to improve readability
'   Removed public variable grstDynRemoteDefaults and replaced with a private variable rstRemoteDefaults in subCaptureData
'  (Only procedure relying on grstDynRemoteDefaults being global was subSetRemoteDatabaseOpenedbyField which is now passed the recordset.
'   Replaced use of a number of large field name constants with hardcoded field names in subCaptureData
'   Renamed gsubStatusBarMsg to StatusBar and contracted the calls to this procedure where possible
'   (eg. removed call statement and removed optional arguments where appropriate)
' 16Mar06 3.0.9039
'   Replaced a few FieldName constants with hard coded field names.
'   Simplifide code and triple commented out dead code.
'   Removed some event logging from CreateNielsenReports.
'   Explicitly cleared VBA.Collection variables in CreateNielsenReports
'   Comment out NotifySlave() and call to it from subCaptureData. (creates email attachment showing results of capture cycle)
'   Comment out SlaveCheckingStatusOfMaster (Sends emails showing results of capture cycle)
'   Comment out mapdrive function used by NotifySlave
'   (Slave not kept on O/N since Neil went full time - when system moved to new Master can email directly from Master)
'   Simplify sub Form_Load and fix problem with subOpenDatabases on SLAVE not finding network copy of Archive database
'   Use new NetDisconnectShare procedure in UnmapShareDiskDisconnectFranchise() (-> used in O/N capture cycle)
'   Use new NetDisconnectShare procedure in fConnectFranchiseMapShareDisk()  (-> used in O/N capture cycle)
' 17Mar06 3.0.9040
'   Change to order of lines in SetGlobalVariables so
'   gsProcommFileOfFiles is set properly and files are then sent to BATA.
' 18Mar06 3.0.9041
'   Remove triple commented code accumulated over past versions
'   Numerous code tidy ups (variable renaming, simplify creation of SQL strings, etc)
'   Changed large number of calls from UploadFilesToOneFranchise and the subs being called.
'   Rsts and Dbs were passed as variants ByRef when they should have been passed ByVal as Rsts and Dbs- they now are
'   Numerous rsts were passed for a sing value already known in the calling code (FranchiseName) so in these cases
'   now pass the FranchiseName instead of rsts. UploadFilesToOneFranchise is a mess and needs further attention.
'   Temporarily changed sub tmrCaptureData so that downloading would occur on Saturday nights (TONIGHT) so I can
'   test the numerous changes tonight and if needed rectify any problems tomorrow
' 19Mar06 3.0.9042
'   Add command line options of "TestMode" (Displays the Test Button)
'   and "LogMemoryUse" (logs Memory use at particular places in program to event log to diagnose memory leaks)
'   Commented out gLocalTest code
'   Wrote CloseDatabases(closes main & archive db) to replace closeDatabase which only closes main database
'   Changed a call to StatusBar in CompactDatabase to use 'pLog:=False' because both databases are closed
' 20Mar06 3.0.9043
'   Reverted sub tmrCaptureData to skipping overnight capture cycle on Friday and Saturday nights
'   Add UseVpn and VpnIpAddress fields to main database and Data Capture tab along with some crude maintenance code.
'   Change fConnectFranchiseMapShareDisk to use VPN rather than making a RAS connection for franchises which have the
'   appropriate VPN fields completed appropriately
' 21Mar06 3.0.9044
'   Changed cmdTestConnection to use common Network Connection and Share code (fConnectFranchiseMapShareDisk).
'   Restored position of FranchiseNode text box back to where it should be on DataCapture tab.
' 23Mar06 3.0.9045
'   Change sub tmrCaptureData to skip Saturday nights ONLY. Request from Neil to give more nights to catch up Win98 machines
'   Review and fix code for maintenance of chkUseVPN and txtVpnIpAddress via DataCapture tab
'   Changed BackStyle of labels from Opaque to Transparent. Noticed all were Opaque by some overlapping controls
'   Replaced txtSupp1ID & txtSupp1Rebate with labels (control array) to be lighter weight and simpler. Captions
'   were only ever BATA-ID & BATA rebate and didn't need complexity of storing values in db. Now also transparent
'   background and do not stick out when there are diff colour schemes like on the SLAVE
'   Added test on startup for availability of VPN: IsVpnAvailable(). Indicate via user interface whether VPN is available and
'   whether it will be used.
'   Minor change to fsPlural and the calls made to it.
' 24Mar06 3.0.9046
'   Disconnect share in IsVpnAvailable() if one is attached during testing availabitly of VPN.
' 28Apr06 3.0.9047
'   Commented out dead constants.
'   Removed triple commented code and large sections of commented out code in cmdTestConnection
'   Give a type to pSplitLongMsgs parameter in gsubAddRecordToLocalEventLog (Boolean)
'   Typed some variables in BackupDatabase
'   Investigated all procedures called from subTransferAllPreLiveDataToTheLiveDataTable in
'   attempt to eliminate memory leaking. Cleared all uncleared objects and made some optimisations.
'   Made check for IsVpnAvailable() conditional on computer being MASTER.
'   Attempt to find and remove the final memory leaks in O/N processing. Rearranged memory logging to
'   more specifically identify remaining problem areas. Investigated subCreateZipFiles() &
'   subBATAUploadWithProcommfunctions() and their call trees. Made numerous fixes, simplifications and
'   rewrites along the way.
' 13Jun06 3.0.9048
'   Remove triple commented code from previous versions
'   Rejig position and size of controls on Uploads tab
'   Rename cmdTestRAS button to cmdTestConnection (it has for a while been using either RAS or VPN connnections)
'   Remove some unnecessary logging of succesful operations, but retain
'   logging of exceptions which can then be attended to.
'       Removed "Terminating RAS connection", "Auto network share connection terminated normally"
'               "Manual network share connection terminated normally"
'   Conditionally call fDisconnectRAS() from UnmapShareDiskDisconnectFranchise() when RAS connection to franchise
'   Remove MoveLast record operations where not strictly necessary. eg Now just report what record we are downloading
'   in O/N downloads rather than the current record out of how many.
'   New generic routines ListBoxClearSelections(), PopulateCombo()
'   Numerous Cleanups (eg remove unused vars/consts made redundant by changes, Change function type of
'   function fdtmYesterday() from string to date as was originally intended, etc)
'   Added and rearrange Memory logging for further diganosis of leaking memory in O/N capture cycle
'   Simplifications of subSaveFranchiseDetails() and renaming table StateSpecific to tlkpState
'   Changed subCaptureData to check for existence of Wholesale data fields to determine whether to download
'   data rather than to check for a RemoteDatabase.mdb version field (Spare1 !) as a flag which may or
'   may not be set correctly. The field was not set correctly in Cliffy. I thnk Spare1 is now completely unused

'   Promotions Changes: Driven by request for applying promotions to particular regions (eg Metro),
'   ------------------  but including a number of improvements and other changes.
'   Toggle "Show Expired" button on promotions tab to "Show Current" so you can return to what you were viewing/just added.
'   Add region combo to Franchise Details on Uploads tab
'   Add printing of regions to printing of Promotions
'   Changed IsPromoApplicableToFranchise() to cater for Promo Regions (previously a different named sub)
'   Set pkt discount on lost focus of ctn discount rather than got focus of pkt
'   Comment out code for deleting promotions b/c button calling it was never visible
'   Comment out importPromos() b/c was not being used, is untested and users have no faith in it
'   Replaced calls to gsubAddSubItemToListview in LoadPromoListview() with inline code b/c more concise and doesn't
'   rely on setting a global ListView item variable before call. (gsubAddSubItemToListview is a very useless procedure)
'   Changed UploadPromotion(ByVal rFranch As Variant, to Sub UploadPromotion(ByVal pFranchiseName As String,
' 14Jun06 3.0.9049
'   Fix bugs introduced in last version
'       - typo/incorrect cut and paste -> CHANGE IsDAOFieldExists(rstDWTempData, TO IsDAOFieldExists(rstRemoteLiveData, ...
'       - reference to wrong column: lvwPromo_DblClick (recall/delete promotions) gets PromoId from list view but
'         addition of column for region changed the index for the PromoId
'       - changes AddPromoToUploads() to cater for recalling old style Promos without a value for RegionID
'       - change to printPromoList() to cater for printing old style Promos without a value for RegionID
'   Rename fraUploadStates to Frame (now one of a large control array) and change backcolour so it doesn't stand out
' 19Jun06 3.0.9050
'   Fix bug in IsPromoApplicableToFranchise() by splitting a boolean test across two
'   if statements to avoid Invalid use of Null error.
'   Changed NetDisconnectShare() and NetDisconnectShare() from Subs to Functions and gave them boolean success return
'   values. Changed calls to these procs to make use of the new return variables.
'   Replace Summary table constants with literals
'   In subCaptureData replaced "Do Until iTaskProcessingRetrys = iMaxCycleRetries" with
'   "For iProcessingCycle = 1 To iMaxProcessingCycles" and simplified use of loop counter
'   with associated changes to fConnectFranchiseMapShareDisk() and calls to fConnectFranchiseMapShareDisk()
'   Substantial reworking of UnmapShareDiskDisconnectFranchise() including: removing 2 second
'   wait/sleep after disconnecting network share drive,
'   Substantial reworking of fConnectFranchiseMapShareDisk() prompted by request from Neil not to
'   dial up VPN franchises when the VPN is not available. (VPN franchises do not retain
'   a redundant modem because the save the money of an extra phone line and put it toward
'   their VPN costs)
' 15Oct06 3.0.9052
'   Removed triple commented code.
'   Simplified code in subCaptureData:- Changed order of If ... Then  ... Stmt,
'   used CnvNulls to reduce number of lines, replaced some verbose field constants.
'   Renamed module and its file from Network_WNet to WNetFunctions.
'   Renamed functions for converting between boolean values and checkboxes to CheckboxToBoolean() and BooleanToCheckbox().
'   Rearranged locations of utility functions.
'   Renamed NetConnectShare() to NetConnectShareDisk()
'   Added then removed IsDestinationReachable() from NetConnectShareDisk() as can't be guaranteed to function correctly in
'   all circumstances. Research shows it conflicts with some VPN software (eg Cisco VPN software Version 3.5.4) and I have
'   shown it to bomb out on Tsg VPN when the destination is unreachable.
'   Add new module Ping.bas with IsPingSuccessful() function
'   Changed fConnectFranchiseMapShareDisk() to verify VPN connection to Franchise with IsPingSuccessful() before attempmting
'   to connect share folder with NetConnectShareDisk which takes long time to time out when connection to franchise not available.
'   Made related changes to event logging in fConnectFranchiseMapShareDisk()
' 30Nov06 3.0.9053
'   Triple comment out some dead variables
'   Remove use of FranchiseShareName field in Franchises table as it is hard-coded in Remote Statistics and if anything
'   else was put in the table the system would fall over (REMEMBER TO REMOVE THE FIELD FROM TSGDATAWAREHOUSE.MDB)
'   Remove un-necessary arguments (FranchiseRASUsername & FranchiseRASPassword) in call to NetConnectShareDisk()
' 10Jan06 3.0.9054
'   Restore what appeared to be un-necessary arguments (FranchiseRASUsername & FranchiseRASPassword) in call to NetConnectShareDisk()
'    It turns out that with the security implemented on around 30 or 40 stores these arguments must be supplied
'   Default timeout in IsPingSuccessful() changed from 10,000 to 1000 (10 to 1 second)]
'   Call to IsPingSuccessful() in IsVpnAvailable() given explicit timeout parameter increasing on successive retries from 1 to 3 seconds
'   Call to IsPingSuccessful() in fConnectFranchiseMapShareDisk() given explicit timeout parameter of 3000 (3 seconds).
'   Changed unlocking of cboRegion from Label_Click() [control array event proc] to lblRegion_DblClick()
'   Remove some triple commneting
' 17Jan06 3.0.9055
'   IsPingSuccessful() requires dlls present in NT, 2000 and XP but not in Win95, Win98 and WinME
'   Old Master (Win98) is still periodically run to collect data from Franchises requiring dialup from a Win98 machine
'   Code changed to use alternative to IsPingSuccessful in IsVPNAvailable() when OS is not NT, 2000 or XP.
'   IsPingSuccessful only then used when VPN is available which is via the New Master (Win XP Machine)
'   Rearranged controls on the Uploads Tab to more clearly reflect their relationship with one another
' 07Feb07 3.0.9056
'   Program had been changed so that TsgDataWarehouse.mdb and Archive.mdb are both opened on startup so any
'   problems with a process having a database open occurs on startup rather than in the middle of the night.
'   Report functions were however closing the Archive db if they had needed it and used it. This caused subsequent reports
'   needing the Archive to crash; or the program to crash on Shutdown when attempting to close an already closed Archive db
'   These problems have now been addressed and the code simplified. Both the databases remain open for the duration of the program
' 19Feb07 3.0.9057
'   Review subCaptureData:
'     -Replace "On Error Resume Next" catering for missing field (OSVersion) in remote Statistics.mdb with IsDAOFieldExists()
'     -Replace nonsensical test "If rstFranchise!FranchisePriceModuleVersion = rstRemoteDefaults!PriceModuleVersion <> "" Then"
'      With                     "If LenB(Trim$(CnvNulls(rstRemoteDefaults!PriceModuleVersion, vbNullString))) Then "
'     -Comment out rededundant variables, comment areas of code which might be improved, some code formatting
'   Preparatory work for importing Batscan Files
'   Modified UploadFilesToOneFranchise() so it won't crash when uploading a file that is open/being used and
'   therefore can't be deleted or overwritten. In this case file to upload will remain in list of files to upload.
'   Removed error handling in UploadFilesToOneFranchise() which doesn't stop the program from crashing but
'   can obscure the reasons why the program crashed
' 21Feb07 3.0.9058
'   BUG FIX: Fixed bug with uploading Msgs that was introduced in last version.
'            Bug fixed by modifying Function DeleteFile() to return True when file to delete doesn't exist.
'   UploadFilesToOneFranchise: Modified so it doesn't crash when gconRemoteDefaultsTableRunUtility ("spare2")
'   field is missing from Defaults table of remote database. (I didn't name or start the use of spare2!)
'   More work toward importing Batscan Files
' 15Mar07 3.0.9059
'   Completed ImportBatScanFile() and integrated it into program
'   Replaced fsVersion() in HeadOfficeStatistics.bas with a new version in shared TsgUtilities.bas module
'   Moved some global variables in to global udt (ease of programming and neatness) and added strBatscanFolder to global udt
'   Rewrote GetDateFrom_yyyymmdd() to validate that passed strng represents a valid date. Included adding new proc IsValidDate_yyyymmdd()
'   Added new proc GetDateFrom_dd_mm_yyyy()
'   Modified LogMemoryUse() so it doesn't automatically refresh the EventLog via its call to StatusBer()
'   Rearranged frmTSGDataWarehouse.Form_Load so that it reads more clearly (done when added code to only show cmdImportBataFiles for Master)
'   Removed a suspect sleep [Call Sleep(2000)] from CompactDatabase()
' 20Mar07 3.0.9060
'   Bug Fix: Fixed ImportBatScan() so that it included TransactionDate in its test for duplicate records in PreLive table
'   Bug Fix: Fixed ImportBatScan so that it converts cents in Batscan files into dollars for currency fields in Prelive table
'   Enhanced Msgs supplied to StatusBar() when Importing Batscan files.
' 10MayO7 3.0.9061
'   Prevent possibility of conflicting franchise selections in Uploads tab
'   [ie Previously could select particular States when 'ALL Franchises' option button was selected  ]
'   [Included coincidental code improvements to CreateUploadsPending() and fGetStateFranchiseList() ]
'   Remove unnecessary call to DisplayUploadsPending from UploadFilesToOneFranchise() - should speed up uploading
'   Promos unconditionally added for Test Franchise whether it is included in Capture Cycle or not
'   Remove backwards compatability for old style Promos (superseded in June 2006) which were not sent according to Region
'   (Old style promos sent according to State, New style sent according to State & Region [Metro, MAJOR Regional, ... ])
'   Added UseLocalFranFolder compiler switch for using local folder (gkLocalDriveFranchiseFolder) as Franchise folder for all franchises
' 25Jun07 3.0.9062
'   Code cleanup: Remove triple commented out code, remove use of gconAllFields
'   Fix non-fatal problem in UnmapShareDiskDisconnectFranchise() when using a local Franchise folder (often used in development)
'   Simplify verbose code for determining time elapsed between connecting to and disconnecting from a franchise
'   Simplify fcreateBATAReports() prior to changes for transmitting data to BATA by FTP
'   Remove fbTheSystemDateFormatIsCorrect() and replaced calls to it with calls to more complete IsDateFmtOk().
'   (Aim is to remove dependence on system date format and hence all occurrences of IsDateFmtOK())
'   Remove project reference to MSSTDFMT.DLL (Microsoft Data Formatting Object Library 6.0)
'   Added numerous comments
' 04Jul07 3.0.9063
'   Remove triple commented code from previous version
'   Remove unnecessary Sleeps from subCaptureData (one 5 secs and another 1 sec for each franchise)
'   Rewrite some code to obviate need for Public variable gdbsRemoteDatabase
'   (variable now local to subCaptureData and passed to subAddRecordToRemoteEventLog)
' 09Jul07 3.0.9064
'   Removed un-necesary Error Handler from subCaptureData - 'On Error GoTo RemoteDatabaseWasNotLocated'
'   Removed tmrEventLogRefresh control by renaming it to ZtmrEventLogRefresh (un-necessary complication & poorly implemented)
'   Replaced overcomplicated fbRemoteModuleIsNotCurrentlyInTheProcessOfUpdatingTheRemoteDatabase with in-line If Then stmnt
'   (function only called once)
'   Replaced 'If rstDWTempData.RecordCount' with 'If rstDWTempData.RecordCount = 0' in subCaptureData and flipped If Then Stmt
'   Set pLogWholeMsg:=True in StatusBar calls in subCaptureData error handlers. Err handling in this proc is illogical and often
'   cloaks errors -> analyse and replace with assistance of more detailed msgs. (see log of 06Jul2007 for poorly handled/cascading err)
'   Added fsErrDetail() to err handling in UploadFilesToOneFranchise (and reworded err msg) for better diagnoses of
'   problems in the overnight capture cycle as per the event log 06Jul2007
' 16Jul07 3.0.9065
'   Added validation code for Capture Cycle - Start Time (Setup Tab) and matching changes to TsgDW.mdb defaults.CaptureStartTime
'   Removed gconWarehouseDefaultsTableCaptureStartField constant and gCaptureTime variable
'   Stopped skipping Saturday night capture cycle (no longer have memory problems requiring a reboot each night -> simplify)
'   Replaced gbDataCaptureCycleIsNotAlreadyRunning with g.bCaptureCycleRunning (was difficult to read in complex expressions)
'   Rewrote parts of subCaptureData (don't attempt UploadFilesToOneFranchise if couldn't connect to remote mdb - could conceivably
'   have caused any number of unidentifiable crashes, bugs etc.
'   Replaced un-necessary error handler RemoteDatabaseInitialisationFailed with in-line err handling, removed redundant form refreshes)
'   Removed un-neccessary test in subCreateZipFiles - (Can't invoke proc mid-capture cycle b/c cycle disables button 1st thing)
'   Create but don't call subCreateZipFiles_NEW from subCreateZipFiles as part of development toward BATA-SFTP
' 25Jul07 3.0.9066
'   Removed ZtmrEventLogRefresh, Removed triple commented code from previous version
'   SubCaptureData: Replaced un-necessary error handler sRemoteDatabaseDefaultsTableInitialisationFailed with in-line err handling
'   Removed label used to confirm VPN is available (retained label used for WARNING when VPN is not available)
'   Non-conditionally call subCreateZipFiles in subCaptureData (had mistakenly changed to call only if called from Timer)
'   Wrote a series of public functions while rewriting Bata Rpting for SFTP
'       GetBATARpt(), WriteUnsentBataRpts(), GetBataRptsFolder(), GetBataRptName()
'       Functions rely on new enum (BataRptTypeEnum) and new queries in TsgDW.mdb (qrptBataTotalSales & qrptBataWSSales)
' 31Jul07 3.0.9067
'   Reinstated skipping Saturday night data capture at the request of Neil (see header notes in tmrCaptureData()
'   Removed another Goto statemenet from subCaptureData
'   Shortened a number of names for code and interface elements to improve readability in subCaptureData etc
'    subUpdateLastDialupResultForFranchise -> subUpdateFranDialupResult, subEditCurrentTaskRecord -> subEditTaskRecord,
'    subAddRecordToRemoteEventLog -> subAddToRemoteEventLog, gsubAddRecordToLocalEventLog -> gsubAddToLocalEventLog
'    stbDataWarehouse -> stb ... and also shortened variable names in subTransferAllPreLiveDataToTheLiveDataTable()
'   Renamed fbRecordIsReasonable() to IsValidData() and made it a more complete test of validity of data. Previously invalid
'   data was getting through and at least in the case of Bata Rpts was being handled by clumsy and non-specific err handling.
'   Fn now tests all Qty and Currency fields and checks for coexistence of Zero Qty (TotQty) and a positive Wholesale Qty
'   Corresponding changes were made to the Master TsgDW.mdb
' 02Aug07 3.0.9068
'   Modified IsValidData() to handle Null fields in PreLiveData. Null fields not necessarily rejected (maintaining
'   previous functionality) but would nice to revisit and tighten up what is rejected and what is not, and where
'   and how it is rejected. Should all be processed through a single function reponsible for validation
' 06Aug07 3.0.9069
'   Removed last Goto statemenet from subCaptureData (GoTo CloseRemoteLiveDataTable) - two activated error handlers remaining
'   Created BataRpt class and moved appropriate Bata Rpt code (functions, enums, etc) in to it
'   Modified current Bata rpt test code called from subCaptureData to use new Bata Report object
' 13Aug07 3.0.9070
'   Remove triple commented code from previous version
'   Reworded message/eventlog entry for rejecting records in subTransferAllPreLiveDataToTheLiveDataTable
'   Added BataRpts collection class (clsBataRpts)
'   Added SFTP class (clsSFTP - wrapper for CuteFtp) which is currently a skeletion with stub procedures that are called
'   Created UpdateUploadTables() to facilitate tracking Bata Rpts Uploads by updating appropriate tables (accompanying db changes )
'   Db changes for making non Bata recognized franchises have a Null FranchiseIdBATA and making the field a uniquue index
'   [included creating CnvZerosToNull()]
'   Fixed problem that crept into code where transfer mismatches not reported (ie mismatch = expected & downloaded record count differs)
' 30Aug07 3.0.9071
'   Numerous changes for new SFTP of Bata Rpts
'    Added BataUploads Grid to Bata Tab (=> VSFlex Grid to project)
'    Added viewing, psuedo uploading of selected, and pseudo uploading of unsent BataRpts
'    Added printing of Bata Uploads grid (and have some commented out code for saving grid as a file)
'    Numerous changes to clsBataRpt and clsBataRpts including rasing events which are handled in the main form.
'   Fixed display of Dialup results so you can return to displaying even log and can switch between 'ALL' & 'Failed' results
'   Removed all calls to LogMemoryUse except in subCaptureData() & subTransferAllPreLiveDataToTheLiveDataTable()
'   Moved 16 procedures from main form to this module as the number of procedures were preventing VB Watch completing processing
'   Rearranged some code in subCaptureData
' 05Sep07 3.0.9072
'   Added export of BataUpload grid to Excel 97 spreadsheet
'   Replaced new TxDate TDatePicker custom control with DatePicker control (Bata Tab)
'   Rearranged order of columns in BataUploads grid so columns could merge on common values.
'   Improvements to subCaptureData using new GetDAORst function from main module
'   Added fdlgCommon for getting filename when exporting BataUploads grid to an Excel Spreadsheet
'   Added functionality to delete Bata Rpts folder when terminating a BataRpts collection (only if no files or subfolders in folder)
'   Added ".whs" extension to file extensions opened by Notepad in subOpenFile. (whs files are Bata Wholesale report files.)
'   Added new parameter to GetAbsoluteFileName() to determine whether to use ShowOpen or ShowSave method to get filename.
' 12Sep07 3.0.9073
'   Removed triple commented code from subCaptureData and continued cleanup of subCaptureData
'    [Use new GetDAOMdb() to open dbRemote & obviate need for in-line error handling, and Logging for when wrapper procedures ]
'    [for closing rst and db encounter and handle an error. Aim is that these wrapper functions should no longer be needed    ]
'   Changed IsValidData() [called from subTransferAllPreLiveDataToTheLiveDataTable] to validate absolute amounts against
'   MaxCurrencyValue and MaxQtyValue stored DW Defaults table. (Request from AWhite via NBarron b/c of dodgy data in stick rpts)
' 25Sep07 3.0.9074
'   Replaced another error handler in subCaptureData with GetDAORst() and other simplifications of subCaptreData
'   Fixed runtime error in Product report [cmdPRPrint_Click] where program crashed if  no data for a selected product.
'   (fixed numerous other bugs and reformatted indentation of procedure while I was at it)
' 03Oct07 3.0.9075
'   Removed triple commented code, renamed/shortened a few variable names
'   Removed unnecessary procedure SetMasterStatus() and replaced with calls to StatusBar where necessary
'   Modified subEditTaskRecord() so it could handle being passed Nothing in prstTaskLog. This enabled removing
'   conditional calls to proc and proc now uses prstTaskLog as a flag for whether to update the TaskLog
'   (TaskLog update not required when testing a Fran)
'   Changes to subCaptureData including new FranID parameter for calling proceudre for single a Franchise data capture.
'   Some code is conditional on new parameter but is not yet excercised as parameter is not in any calls yet.
'   Gave some buttons a fluoro green background to highlight to Neil whether they are used and whether they can be removed
' 04Oct07 3.0.9076
'   Change calls to StatusBar in CompactDatabase & zipDataBase to add to event log (assist tracking recent error in these procs)
'   Increased timeout in zipDataBase() shell to pkzip
'   Removed triple commented code, renamed/shortened a few variable names
'   Removed unused buttons for copying MasterMdb from and to IOMEGA zip drive (and removed associated code)
' 11Oct07 3.0.9077
'   Reverted calls to StatusBar in CompactDatabase() & zipDataBase() [called sequentially in backupDatabase()] to pass pLog:=False
'   (pLog:=False had been inadvertently commented out and caused the program to crash as no mdb was open to Log to)
' 27Nov07 3.0.9078
'   Removed fsPlural() and replaced calls to it with calls to the more generic Plural()
'   Moved subOpenFile() into this module
'   To reduce number of ctls so VB Watch Error handling can work properly
'    - Removed btnPurge (functional but never used) and txtPurgingStatus
'    - Removed txtPurgeStartDay NON functional control (Value read from Db and written to Db and thats it!)
'   Wrote PurgeLiveData() & PurgeEventLog() and call each unconditionally at end of subCaptureData()
'    (TsgDw.Defaults!MonthsOfEventLog field added)
'   Changed fdbGetCurrentLiveDB() to use date data types rather than strings.
'   Change g.rstDWDefaults!LiveDataStartDate data type from text to a date field. (& associated changes)
' 31Jan08 3.0.9079
'   Code cleanup: removed commented out code, changed eventlog entry for adding a new franchise
'   Changed subCaptureData so New Bata Rpts test code is only called during O/N cycle (not during a manual daytime run)
' 07Feb08 3.0.9080
'   Changes made to speed up EventLog logging and retrieval. Utilises EventLog.Sequence field added to TsgDW.mdb 31Jan2008
'   - gsubAddToLocalEventLog(), gsubRefreshEventLogDisplay(), & replaced local EventLog rst with global rst for performance/caching
' 13Feb08 3.0.9081
'   Change gsubAddToLocalEventLog() to reinstate and improve spliting long LogMsgs across multiple records and make it default behaviour.
'   Also made the rst for the EventLog a static variable to hopefully speed up the procedure.
'   Various code cleanups including removal of triple commented code
' 20Feb08 3.0.9082
'   Hard code wholesale field values to zero in ImportBatscanFile() rather than relying on table default values.
'   Fine tune gsubAddToLocalEventLog() - was losing first char on split lines when lines were split on a space
'   Renamed IsUseLocalFranFolder() to IsUseLocalDriveFranFolder(). [Now shared across multiple Tsg applications]
' 29Feb08 3.0.9083
'   Modified RemotePromotionRecall() so if a promotion being recalled had been notified/displayed via RemoteStatistics,
'   the remote database (PromoEnd) would be edited so that it would expire immediately
'       Overahauled RemotePromotionRecall() & UploadPromotion() in the process.
'   Fixed problem sorting Bata Rpt grid by Report Type (Report Type column in grid somehow mistakenly was a date type - changed to any)
'   Removed rogue debug.print statement from ConfitureBataTabButtons()
'   StatusBar() & gsubAddToLocalEventLog() - Simplified/removed option for tuncating EventLog entry (non-conditionally log whole msg)
'   gsubAddToLocalEventLog() - Optimised for case where EventLog entry doesn't require splitting
'                            - Reinstated MsgBox when program unable to log to EventLog table - probably indicates major problem
'                              warranting immediate investigation
'                            - Added more extensive reporting of error by interrogating DAO.Errors collection and returning Err Descs & Nos)
' 05Mar08 3.0.9084
'   Fixed bug where EventLog rst wasn't reopened when database were closed for backup
'     Made EventLog rst a global var rather than static var in gsubAddToLocalEventLog() as when Dbs were shut then reopened
'     (part of bakcup process) the test of the static var was showing it as NOT Nothing and static rst var wasn't being reopened
'     Global rst for EventLog is now reopened when opening databases
'   Improvements to backupDatabase() which could really do with a rewrite. [Error reporting improved, 10 second delay removed]
'   Fix bug (runtime error) "Product Report"/cmdPRPrint_Click() - when selecting multiple products for all franchises the
'   description sorting loop tried to compare the last row with the last row plus one. Loop now traverses from first row to last but one
'   Minor code cleanups
' 13Mar08 3.0.9085
'   New version completely driven by desire to add whole of procedure error handling for GetDaoMdb().
'   Durnig O/N processing after testing for existence of mdb file in GetDaoMdb() the same proc could
'   not then retrieve a file object from the file and the program crashed. (network dropped out?)
'   Whole of procedure error handling for added to GetDaoMdb(), then GetDaoMdb() merged with best of GetDaoDb()
'   DAOUtilities and TDAOUtilities modules merged into TDAOUtilities then GetDAOMdb() & GetDaoRst() moved from this
'   TDAOUtilities added to the project
'   After some live testing GetDaoMdb() will be renamed to replace GetDaoDb() in TDAOUtilities
' 29Apr08 3.0.9086
'   Finalize merging of GetDaoMdb & GetDaoDb into GetDaoDb() and replace calls to GetDaoMdb with calls to GetDaoDb
'   Add automatic purging of FranchiseUploads at the end of subCaptureData
'   Disable ctls for creating new promotions for all installations except MASTER [see EnablePromotionCreationCtls()]
'   [V369 Disable ctls for editing promotions for all installations except MASTER   ]
'   [V369 EnablePromotionCreationCtls() renamed to EnablePromoEditCtls()            ]
'   Minor code cleanups
' 20May08 3.0.9087
'   O/N cycle includes BATA SFTP running in parallel with Procom uploading of Bata Reports (rpts now actually sent)
'   strBataFtpHostAddress, strBataFtpUser, strBataFtpPwd and bMaster members added to global udt (g)
'   Additional parameter (pRefreshLog) added to StatusBar() to control refreshing of EventLog ListView control
'   Other minor code cleanups/simplifications (eg No longer display 'Refreshing Event Log..." in status bar when refreshing
'  '**********************************************************************'
'   pbCalledFromTimer NOW USED FOR SETTING RETRY TO TRUE FOR ALL          '
'   UNSUCCESSFULLY COMPLETED TASKS WHEN MANUALLY PERFORMING CAPTURE CYCLE '
'  '**********************************************************************'
' 21May08 3.0.9088
'   Rejig StatusBar() so it can't refresh EventLog unless an item has been logged as well as displayed in Status Bar
'   Fix Sql in clsBataRpts.AddRpts_Unsent to select FranchiseIDBATA which was referenced in code in last version but not added to SQL
'   Replace special case calls to StatusBar (no Msg but refreshing log) with calls to gsubRefreshEventLogDisplay
' 22May08 3.0.9089
'   Add DoEvents to StatusBar() to ensure interface changes are displayed (problem has always existed).
'   Also removed a redundant line from StatusBar()
' 29May08 3.0.90
' Fix problem where Nielsen Rpts not being automatically created on Monday.
' Added DoEvents to top and tail of subTabMainClick to improve screen refreshing.
' 12Jun08 3.0.91
'   Remove fsNodeType(), optMaster, optSlave, optPleb and containing frame -> With fewer ctls VB Watch is now working completely again
'   Fix problem where EventLog text wasn't being trimmed before being added to EventLog table
'   Prepend Error Number to string returned from gfsRASErrorMessageFromErrorCode()
'   Change FTP BATA Rpts. Populate 9th field in detail records with "0' rather than leaving them blank.
'    - (9th field is Total sales value at actual sell price Ex GST in cents.)
'   Some optimisations to Bata Rpts
'   Code cleanups (eg. Remove triple commented code before adding new triple commented code, remove dead code)
' 19Jun08 3.0.92
'   Move labels not referenced by code into the generic label() array.
'   - label1, label2, label3, label8, label10 label15 label16 label18, label19
'   - label20, label19 label20 label22 label23 label24 label25 label26 label28,
'   - Label5() and Label6() two item arrays not referenced in code
'   - Label11() one item array
'   - lblFranID() on item array
'   lblState(7) - label array of one item changed to a single control
'   Line1(0) - line array changed to a single control
'  Remove commented out code from previous version. [NodeType code, openSession(), SendEmail(), SendFaxes()]
' 26Jun08 3.0.93
'   Cut a profiling version of TsgDW for manually running data caputre cycle
' 07Jul08 3.0.94
'   Added carriage return to trailer records of BATA upload files. (requested by BATA post live implementation)
'    (added vbNewline which is actually a carriage return followed by a line feed)
'   Removed grdBataRpts_DblClick b/c can accidentally start time consuming process of loading many reports
'   fsGetRmVersion() now accepts RemoteDflt rst as a DAO.Recordset rather than a variant
'   Include uploading of Bata Rpt files in manual data capture (while Neil is on holiday keep all reporting in synch)
'   and swap order of FTP and Procomm transfer of files. FTP transfer now occurs first
'   Add sorting tooltip for Bata Rpt Grid when hovering over header row
'   Optimised clsSFTP.Upload (somewhat un-necessarily since speed problem was with BATA FTP server although they initially denied it)
'   Changed clsBataRpts.Upload so that if an upload fails b/c remote file exists then the UploadTables will be
'   edited to reflect the file has been uploaded so that the upload is continaully retried. (EventLog will indicate file pre-existed)
'   Replaced Global Const gconFranchiseRMVer As String = "FranchiseRMVersion" with in-line coding
'   Fixed bug in gsubAddToLocalEventLog(). Case where string segment being created exactely equalled desired segment lenght
'   was not being handled properly.
'   Removed global const gconFranchiseRMVer
' 17Jul08 3.0.95
'   Changed clsBataRpt to save temporary report files to Windows Temporary directory rather than C:\TS\Data\BataRpts
'   Removed deletion of BataRpts subfolder (if empty) from BataRpts class terminate. Rpts now saved to Windows Tempporary folder
'   clsSFTP: big tidy up and added numerous comments, simplified Upload method and removed unnecessary optimisation to simplify code
'   Continued removing unneceesary Master/Slave status code.
'    -  Removed all references to following fields in App.Path\Defaults.mdb!default: ThereIsAnotherMachine, MasterNodeName
'   Miscellaneous code cleanups
' 31Jul08 3.0.96
'   Replace inaccurate fsGetRMVersion() with accurate GetRMgrVersion()
'   Stop automatically uploading BataRpts when peforming MANUAL/DAY-TIME capture cycle
'   Simplify cmdBataTabUploadUnSent_Click()
'   Simplify CreateNielsenReports()
'   Remove triple commented code
' ??Aug08 3.0.97
'   GetRMgrVersion moved from TsgDW to Remote Statistics which has access to franchise folders where Retail Mgr mdb resides
'   '' Version comments need completion ''
' 21Aug08 3.0.98
'   Remove redundant Procomm/Modem uploading of Bata sales data.
'   Clean up some other redundant code
' 04Sep08 3.0.0099
'   Modify ImportBatscanFile to conditionally set wholesale data fields in TsgDW.mdb to zero only
'   when they were previously Null as we may soon import BatScan EOD wholesale sales summary files.
'   The changes though small are important and are given copious in-line comments.
'   Clean up all triple commentd code that was commented out when Procomm was decomissioned.
' 25Sep08 3.0.0100
'   Main reason for version is to change type of field PromotionID in remote RStats.mdb from integer to long (was causing run-time error)
'   UploPromotion altered to (A) chnage remote PromotionID field type to Long (from integer) and purge expired promotions with an
'   PrommoEnd older than a month at the same time, (B)more arccurately return success or failure and to more fully record this
'   result in the EventLog. Added UpgradeRemotePromotionsTable()
'   UploadFilesToOneFranchise() modified so that when UploadPromotion fails it is recorded as failed (not successfull) and will retry later
' 09Oct08 3.0.0100
'   New fn UpgradeDefaultsOSVersionFld() for adding OSVersion field to remote [RStats.mdb].defaults table - particularly Cliffy
'   Code to flag stores without RMgrVers & OSVersion in [RStats.mdb].defaults table with ultimate aim of removing code catering for exceptions.
' 06Nov08 3.1.1
'   Minor changes to utility code. Get version numbering to more sensible format
' 08Jan09 3.1.2
'   Clean up subCaptureData and cmdTestConnection:-
'   - remove code catering for missing remote OSVersion field as it is no longer required
'   - add code to create remote RMgrVer field as field still does not exist at all Franchises.
'   - removed UpgradeRemotePromotionsTable(), UpgradeDefaultsOSVersionFld() and all calls to them
'   - added UpgradeRMgrVerFld()
' 09Jan09 3.1.3
'   Fix to subCaptureData() and cmdTestConnection() - ON Capture cycle crashed last night
'   Wrapped collecting remote RMgrVerFld with Cn [ie ConvertNulls to cater for old RStats fields which Allowed Null values]
'   eg strRMgrVer = Cn(rstRemoteDefaults(gkRStatsMdbRMgrVerFld).Value, "")
'   Fixed code which creates RMgrVer field so that the field is populated with vbNullString rather than Null
' 25Feb09 3.1.6
'   Add FranName to BataRpt object for reporting in events to the UI
'   Accommodate FranName in BataRpts Add method
'   Add UploadSummary property to BataRpts object (and use in capture cycle)
'   Remove Rpt objects as successfully uploaded in preparation for retry code
'   Use new BataRpt & BataRpts properties for updating the UI during and after uploading rpts
'   Partail rewrite/improve UploadFilesToOneFranchise() so that DB is passed ByRef rather than opening two separate variables on the same mdb.
'   Changes to calling of UploadFilesToOneFranchise()
'   Partail rewrite/improve fConnectFranchiseMapShareDisk()
'   Incremental improvements to subCaptureData in preparation for cmdTestConnection and subCaptureData calling common code
'   Remove logging of successfull events ("Verifying VPN connection to ...", & "Network share connected") in agreement with policy of moving toward logging exceptions and problems rather than the overkill of loggin every minor success
'   Modify subIgnoreErr_CloseRst(), subIgnoreErr_CloseDbAndSetToNothing(), LogBugFix() to facilitate tracking and improving problem code
' 03Mar09 3.1.7
'   Remove code for logging memory use
'   Remove code for adding RMgrVer field to RStats.mdb during data capture and no longer cater for missing field.
'   Add code for adding Wholesale data fields to RStats.mdb during data capture if they are missing.
'   Minor code cleanups
' 03Mar09 3.1.8
'   Fixed bug in UploadFilesToOneFranchise(). Proc was not closing db and disconnecting franchise when it made its own connection
' 19Mar09 3.1.9
'   Rename TsgDw.defaults!QueryLastRun to LastAutoCaptureDate
'   Temporarily hard code CaptureStartTime (gkCaptureStartTime=11:45pm) and use it to 1. trigger O/N capture time and 2. reconcile task log.
'   Remove triple commented code (frmTsgDataWarehouse, ...) and various cleanups
' 319Mar09 3.2.0
'   Rename TsgDw.defaults!QueryLastRun to LastAutoCaptureDate
'   Temporarily hard code CaptureStartTime (gkCaptureStartTime=11:45pm) and use it to 1. trigger O/N capture time and 2. reconcile task log.
'   Change Upload method of clsBataRptst to: have 3 attempts at uploading each rpt, raise AfterUpload event on last attempt,
'   use different wording when setting class UploadSummary property
'   Minor cleanup (rename parameter and remove/update comments)
' 02Apr09 3.2.1
'   Temporarily convert null data in Wholesal fields of RemoteStatistics to 0 while Cliffy RStats dbs are standardised at
'   Northgate and Upper Eastlands
' 08Apr09 3.2.2
'   Further untangle the mess that is subCaptureData in preparation for DataCapture of selected franchises
'   - remove last Error Handler Routine from subCaptureData(): [DataTransferWasInterrupted], ...
'   Add EventLog entry for verifying VPN connection to franchise so there is a clear indication log as to whether
'   the capture cycle is doing a 2nd or 3rd pass over the franchises
' 08Apr09 3.2.3
'   Remove code for creating and uploading a 'PromoMsg' - was not used and is not catered for with new auto promotion code
'   Remove RAS/Fax option box on 'Upload Tab'. Investigation showed that it had no effect
'   Remove commented out code from subCaptureData() and comment out error handlling around setting remote DatabaseOpenedBy field
' 30Apr09 3.2.4
'   Changes to subCaptureData()
'   1. Comment out pCaptureOne_FranID code in
'   2. Remove code creating Wholesale fields (WholesaleQty & WholesaleActualSell) in RStats.mdb - all franchises catered for
'   3. Add code to cater for TransactionDate field in RStats.mdb being of Date type in preparation for migrating all fields
'      in RStats.mdb that hold date data to now be of date type instead of text => remoove reliance of particular Windows date settings
' 26May09 3.2.5
'   Simplify subAddToRemoteEventLog() to attempt no retries if it fails.
'   Commented out 'Set Remote System Time'. Hid ctls and commented out code (set flags in RStats.mdb for use by RStats program)
'    (didn't work for upload immediately as per a lot of other upload types which were fixed)
'   Fixed upload immediately for a range of upload types.
'   Replaced code referencing TransactionDate fields that used ambiguous string date fmt (mostly ddmmmyy) and am closer
'   to removing code reliant on a specific date format. [RStats upgrade 5.3.7 is removing the reliance on specific date fmt)
'   [replaced uses of gconMicrosoftAccessDateFormat "mm/dd/yyyy" in Where Clauses with MSSqlDate(date) ->#dd mmm yyyy#]
'   Once the remaining ctls that store dates as strings are replaced there may only be a few date string manipulation
'   routines that need investigating before we can use sensible date manipulations
'   Remove code in subCaptureData which determined whether RStats Wholesale flds exist. (Used to TsgMsgCentre to confirm they exist everywhere)
'   Moved backupDatabase in subCapture to after purging as the routine contains database compacting code
'   Remove commented out code from subCaptureData
'   Modifications to GetDate_FromTsgDAODateFld() & SetDate_ToTsgDAODateFld() using conditional compilation
'   constants so they can be shared with TsgDw and remote s/w where they log problems/status to the TsgMsg.mdb
' 27May09 3.2.6
'   Fix problem of event logging franchises with missing sales twice
'   Comment out a few dead variables.
' 02Jun09 3.2.7
'   Remove triple commented code from previous versions and assoicated Z_Prefixed ctls (ZchkSetSystemTime, ZchkNewPromoMessage, ZFrame, ZoptRASorFAX)
'   Fix some sloppy code from version 3.2.5 where MSSqlDate() had not been applied or had been incorrectly applied to reporting code
' 26Jun09 3.2.8
'   Remove Cn (convert nulls) when collecting Wholesale fields in subCaptureData()
'   Bug fixes to breaktext() - new version now resides in TsgShared.bas
'   btnSaveMessage_Click(): removed extra line feed before last line of dashes in Popup messages
'   Disable tabs for Business Manager PC as appropriate
' 26Jun09 3.2.9
'   Change subCaptureData() to iterate through a collection of FranIDs (picking up single record Fran rsts) rather than
'   moving through a rst of all Franchises. This is another step toward offering DataCapture for selected/nominated franchises
'   Remove last use of gconMicrosoftAccessDateFormat ("mm/dd/yyyy")
'   Replace gconFranchiseTableLastSuccessfulCaptureDateField with literal and change DataType of field it refers to from string to date
'   Disabled Uploads tab for all but Master PC
'   Add Ctl-Shift-A key press combination to show version (Data Capture tab with version info not enabled on Rpt Server)
'   Call SplitText in gsubAddToLocalEventLog() to replace inline code which had a rarely occuring infinite loop bug
'   Cleanup some triple commented code
' 16Jul09 3.3.0
'   Replace 'Test Connection' with 'Capture Selected' on Data Capture tab
'   Move StatusBar updates for calling Neilsen report cycle into CreateNielsenReports()
'   Fine tune event logging of failed BataUplaods
'   TsgDw.Defulats.LastAutoCaptureDate now set only in tmrCaptureData() [manipulation removed from subCaptureData()]
'   Remove MAPIMessages and MAPISession ctl from main form (not using and won't be for forseeable future)
'   Remove TsgDw.Defaults!CaptureCycleRetries: now hard coded - no need for this to be configurable
'   Remove some un-used global constants
' 22Jul09 3.3.1
'   Replaced single promotion recall (via double click on lvwPromo) with multiple promotion recall via multiselect
'    on lvwPromo and cmdPromotionRecall button and sub procedure.
'   Removed updating of Summary table in subCaptureData (may reinstate functionality after procedure rewrite is complete)
'   Use subIgnoreErr_CloseRst() to close rstTaskLog toward end of subCaptureData() loop
'   Remove Tooltip from ('Double-click to delete a promotion') from promotions list
'   Remove lblNoAttempts & txtNumAttempts from Settings tab
' 23Jul09 3.3.2
'   Bug fix for 'Caputure All' modes of subCaptureData() [ie bContinue = False ' Initialisation is VITAL ...]
'   Enable/Disable cmdPromotionsRecall according to whether lvwPromo has list items selected
'   Set default button for MsgBox confirmation of cmdPromotionRecall to No
' 04Aug09 3.3.3
'   Remove uses of TaskLog table. Required massive simplification of subCaptureData(), rewrite of fConnectFranchiseMapShareDisk()
'   and removal of subEditTaskRecord(). Simplifications of fConnectFranchiseMapShareDisk() and standardisation of
'   subCaptureData() algorithm remains to be done
' 06Aug09 3.3.4
'   Change code to reject records where ((Quantity = 0) AND (WholesaleQty = 0))
' 11Aug09 3.3.5
'   Comment out subCaptureDataReconcileTaskLog() and fix up problems with autostarting of the Capture Cycle
'   Refine autocapture to include any included franchises with pending uploads even if 'data capture' isn't pending
'   (For example data may have been collected manually through the day, but uploads have since been added)
' 20Aug09 3.3.6
'   Update Franchise & Stock table in Archive db at end of automatic data capture cycle as per request from Neil.
'   Remove CaptureStartTime from Dw.mdb and remove maintenance code for this field. May reinstate on later code reworkings.
' 03Sep09 3.3.7
'   TEMPORARY change for uploading specific Forbes data (1st - 23 May inclusive) that has been placed in TsgDw
'   Minor changes to some EventLog entries
' 08Sep09 3.3.8
'   Remove TEMPORARY change for uploading specific Forbes data (1st - 23 May inclusive) that was placed in TsgDw transferred then purged
'   Modify IsValidData() to classify sales data where Barcode has embedded Single Quote(s) (ie ') as INVALID
'    (change an in-line comment and an event log item to reflect this)
'   Modify fbBarcodeIsATobaccoProduct() to use SqlQuote() when constructing a Where Clause with a Barcode value
'   Minor cleanup [removed a commented out line and modified procedure header notes in IsValidData()]
' 25Sep09 3.3.9
'   Added option for adding Cigars ('Cigar') and Tobacco ('TOBAC') to export file created from Export button on stock tab.
' 01Nov09 3.4.0
'   Add error trapping to UploadFilesToOneFranchise() and alter calls to it to log the returned error msgs
'   Alter some of the existing EventLog entries from UploadFilesToOneFranchise()
'   subCaptureData(): Remove some code through judicous use of Cn() procedure
'                     Use GetDaoRst() for remaining few rsts in subCaptureData() not created with this procedure
'                     Extend scope of error handling & simplify error handling Gosub at FranProcessingInterrupted: label
'   Removed dead variable and empty procedure
'   Standardise and add more detail to EventLog entries for recalling promotions (originally to diagnose reported bug, but will be useful)
'   Add more detail to MsgBox confirming promotion recall. Include count of promotions since promotion list can be truncated by MsgBox
'   Log requested promotions recall (confirmed by MsgBox) to EventLog (doesn't look pretty in event log but is useful diagnostically)
'   Change woeful naming of grstStock parameter in "Sub addToUpdateFile(grstStock As DAO.Recordset, ...)
' 08Nov09 3.4.1
'   Add facility for uploading BataRpts for a seletecd franchise and date range (added frmUploadBataRptsForSelFranAndDate)
'   Added AddRpts_FranIdDateRange() to clsBataRpts & some minor changes
'   Rationalised code for setting system date dependent settings
'   Simplified some event log entries
'   Removed use of following fields from TsgDwMdb.Defaults
'       (DefaultRemoteModuleVersion,DefaultPriceModuleVersion, LiveDataStartDate, NonTobaccoPrompt)
'   Removed default version textboxs from Version tab and no longer populate Upgrade fields in ListView ctl of Versions tab
' 10Nov09 3.4.2
'   C:\TS\Temp folder replaced with C:\TS\Logs folder. Folder was/is only used to store remoteupgradelog.txt files
'   (existence of replacement folder no longer test for but rather the folder is created as neeeded)
'   PkZipC.exe is now stored in C:\TS\Programs directory. Testing for the file now no longer ends the program startup
'   but simply gives a warning that zipping files will not work until PkZipC.exe is copied to the appropriate folder
' 21Nov09 3.4.3
'   Extend recreating, viewing and uploading of Bata reports beyond the range of the live data base to the range of the Archive Database
'   (achieved through numerous changes including new linked LiveDataArchive table in TsgDw.mdb and a number of new queries)
'   Added PopulateListBox() and reloacated PopulatedCombo into TDAO.bas
'   Add following members to global udt (g): strLogFolder, strRptsFolder, strNielsenRptsFolder and adjust affected code
'   Delete gsubFindProgramsAndFolders() and replace its functionality with code predominantly in SetGlobalVariables() - renamed from subSetTSGDataWarehouseDatabaseAndGlobalVariables()
' * Numerous code fixes for code which misguidely attempted to disable error handling within error handlers by means
'   other than 'resume' or exit sub/function/property including a big cleanup of such code in subCaptureData()
'   General cleanup of comments in subCapturedData()
'   Added temporary kludge to prevent uploading/recalling promotions() during Selected Franchise data capture as it     (''' revisit kludge ''')
'   magnified an already existing problem with code which must later be addressd but for the moment remove magnifiction
'   of problem before being given authority/commission to work on the problem
'   Added 'ALL REGIONS' selection for creating promotions. (Included new constant frmTSGDataWarehouse.mkAllRegionsID)
' 22Nov09 3.4.4
'   Fixed previous version change to loadNonCompliants() that inadverantly logged event that it shouldn't have
' 25Nov09 3.4.5
'   Added facility for export stock files of selected stock items.
'   Removed the use of the following two field in TsgDwMdb.Defaults for setting fld values when creating a new franchise:
'   (DefaultRemoteModuleVersion, DefaultPriceModuleVersion - was an oversight in previous changes b/c flds had already been removed/prefixed)
' 08Dec09 3.4.6
'   Fixed oversight in subCaptureData where active error handling was not completed with a Resume, Exit Sub or Exit Function
'   (caused subsequent errors to be handled by calling procedure which in most cases was handled by vbWatch error handling)
'   Fixed oversight in UploadFilesToOneFranchise where strLogFolder was replaced with g.strLogFolder
' 16Dec09 3.4.7
'   Major version change is to provide a button for flagging franchises as closed and to only display such franchises when appropriate
'   Version relies on Mdb changes
'   - Renamed qryBataFranchises to qryFranchiseBata
'   - Add qryFranchiseLive
'   Fixed problem with writing Neilsen reports to a new location. Restored writing of rpts to the original location
' 22Dec09 3.4.8
'   Very many changes for flagging Stock as deleted
'   (relies on following changes to TsgDw.mdb: new deleted field in Stock table, new query qryStock)
'   Removed following global variables (& dependencies on them): gdbsStockDatabase, grstStock
'   Numerous tidy ups while making changes
' 20Jan10 3.4.9
'   Fix bug where non-compliant 'ALL REGIONS' promotions were not reported as non-compliant
'   [added 'Or (pPromoRegionId = mkAllRegionsID)' to test in IsPromoApplicableToFranchise()]
' 19Feb10 3.5.0
'   Added facility for printing NonCompliantPromotion details for selected franchises (while at it cleaned up
'   code and optimised loading of NonCompliantPromotion list view be reducig redundantly populating the control)
'   Modified btnExportStock_Click() to allow creating stock price files from external database
'   (& modified existing code to use CommonDialog for selecting external datbase instead of input box)
'   Removed last 'Exit Sub' (excluding the one preceding Error Handlers) from subCaptureData()
'   Added strUploadFolder as a member to global udt
'   Create Uploads folder on startup if it doesn't exist (Required immediately after installation
'   because apart from uploading the program uses the folder for things like CREATING stock export files)
'   Cleaned up comments and commented out code
' 25Feb10 3.5.1
'   Added zipping of Uploads folder at the end of Capture Cycle
'   (New proc zipUploads() added and called from BackupDatabase
'   Added a call to Cn (ConvertNulls) in TransferSalesRecord to cater for Null value in WholesaleActualSell coming
'   from data collated by RStats Version <= 3.4.5
' 26Feb10 3.5.2
'   Fixed problem with changes for zipping Uploads folder and copying it to USB/IMOEGA.
'   (wrong path given for copy destination and caused BackupDatabase to bail out)
' 12Mar10 3.5.3
'   Added automatic FTP of 'Nielsen Reports' to AZTEC
'   Removed unused & unnecessary code/controls/constants etc (Mainly from Settings Tab)
'   Cleaned up various sections of code
' 29Mar10 3.5.4
'   Completed changes whereby program excludes frans flagged with 'Live = False' from reporting
'   (cf excluding frans excluded from the capture cycle). This better caters for franchises
'   which email in their data (eg BATSCAN franchises). Changes most notable in missing days report
'   on 'Sales Reports' tab, and in the EventLog when listing franchises with Sales Data missing for the last 5 days.
'   Numerous code cleanups and simplifications
' 07May10 3.5.5
'   Add facility for creating promotions of different grades and applying the promotions accroding to grades assigned to franchises.
'   (required addition of PromoGrade field to Franchises and Promotions tables. tblFranchisePromotions also added.)
'   Exclude ALL-FRANCHISES item from Region combos in DataCapture/Franchise tab and Promotions tab
'   FranchiseRebate - remove code for maintaining FranchiseRebate (and accompanying removal of TsgDw.mdb.Franchises!FranchiseRebate)
'                   - remove rebate calculations and displays from Stick Reports
'   Bug fix in UploadFilesToOneFranchise() - FileSystemObject variable was being used without object being created
'   Add Promos To FranchiseUploads  - when franchise is excluded from capture cycle as the exclusion may be temporary (previously excluded)
'   Create FranchiseUploads         - when franchise is excluded from capture cycle as the exclusion may be temporary (previously excluded)
'   Minor cleanups
' 10May10 3.5.6
'   UploadPromotion() - Bug Fixes
'   1. Remove rstRemotePromo.Close where rst hadn't been opened
'   2. Change last minute check that PromoGrade of Promo & Fran match [ie. hasn't changed since Promo was created]
'      so that check isn't applied to old legacy promotions without a Promotion Grade.
'   Rename rst variables in same procedure to assist clarity
' 19May10 3.5.7
'   Fix typo in field name when creating SQL string in IsPromoApplicableToFranchise()
'   [Was causing crash at end of O/N cycle when loadNonCompliants() called IsPromoApplicableToFranchise()]
' 15Jun10 3.5.8
'   Added facility in Stk creation for creating CigPkts linked to CigCtns so that when 'linked' pkts are uploaded
'   they can be programatically linked to Ctns in RMgr by Price Module. Uses changes to TsgDw.mdb such as a
'   additional Package table same as that in RMgr
'   Numerous cleanups
' 27Jun10 3.5.9
'   Overhaul of NonCompliance testing
'   Remove rounding up to nearest 5 cents for 3,4,8,9 and chaned rounding to
'   use custom predicatable TRound() instead of  unpredicatble VBA.Round
'   Added saving of NonCompliant reports to files
'       2 Options: 1 Text file report, 2. A CSV report
' 21Oct10 3.6.0
'   Add 'ALL [States]' selection for promotions (creating, recalling, reporting, ...)
'   Tidy up printout of Promotions list while accommodating the new 'ALL [States]' selection
'   Modified msg in confirmation of recalling promotions msgbox to give more information & to be easier to read
'   Create local NewMessages Folder on Startup
'   Numerous cleanups of triple commented code
'   Changed order of a few 'If Then' stmts with ridiculously large 'If' clauses coupled with micro 'Else' clauses
'   Comment out conditional code in IsVpnAvailable() that catered for TsgDw.exe being run on Win95,Win98 or WinMe (it no longer is
'   Comment out now redundant code in fOKToUploadItem(), IsPromoApplicableToFranchise()
' 23Aug11 3.6.1
'   Added code to run O/N data capture cycle according to boolean flag in TsgDw.Defaults!AutoDataCaptureCycle
'   Modified code in BackupDatabase to re-enable tmrCaptureData only if is a Master installation (g.bMaster = True)
'   Moved a line of code for readability
' 11Sep11 3.6.2
'   Collect StockCurrent & StockDiary tables from RStats.mdb at Unicenta oPOS (Japanese) franchises
'   - requires two new tables in TsgDw.mdb: FranStockCurrent & StockDiary
'   Wrote new procedurea GetRemoteStkData() and called it from subCaputreData()
'   Removed commented out procedure: WriteStockToOpenTextFile() and fixed a few typos in comments
' 21Nov11 3.6.3
'   Fix bug where 'O/N Data Capture' & 'Capture ALL' crash at end of cycle when updating Franchises or Stock table in Archive.mdb
'   Renamed tmrCaptureData to tmrAutoDataCapture in the process for readability for those that follow
' 30Aug12 3.6.4
'   Optimise creation of promotions (apparently has become so slow that saving a new promotion can take 3 minutes)
'   Chief optimisation was limiting size of a Dynaset built from tblFranchisePromotions that a Find operation is performed on.
'   Accompanying code changes was addition of relationship with Cascade Delete between Promotions & tblFranchisePromotions tables in TsgDw.mdb
'   Numerous code cleanups and comments to guide the way for improvements - particularly for tracking promotions on an
'   individual franchise basis so promotions can be recalled outside of O/N processing (e.g. via 'Capture Selected' button)
' 22Apr13 3.6.5
' *** EXTENSIVE CHANGES ***
'   Migrate existing DAO code to ADO
'   - remove all but one use of RecordCount property (need to investigate remaining use in next version)
'   - begin standardising SQL: using SqlQuote() to create SQL, reomving some occurrences of
'     query delimiter (ie ;), etc
'   Remove RAS code (all connections now over the VPN)
' 22Apr13 3.6.6
'   Fix 'Out of Memory' bug when trasnferring records to Archive database by changing
'   cursor location on g.cnnArchive to )
'   Fix a number of bugs by changing SQL wild cards in WHERE CLAUSES from '*' to '%' (ADO & MySQL compatible)
'   Rename fsFranchiseNameFrom() to GetFranName() and simplify procedure.
'   Add RstTypeEnum items for emulating DAO rst types.
'   Add SetValStringValue(). Used in setting MySQL connection string.
' 08May13 3.6.7
'   BugFix in AddRpts_Unsent(). Missing/commented out FindFirst in loop needed reinstated.
'   Program repeatedly didn't find BataRpts as uploaded and therefore attempted to create
'   and upload the last three months (as determined by settings of live data to keep).
' 09May13 3.6.8
'   Fix runtime bug in AddRpts_Unsent() that crashed BataUploads. Cater for Empty rst when
'   searching tblBataUploads. Optimisation: limit searching tblBataUploads to to Live franchises.
' 22May13 3.6.9
'   Tsg Data Warehouse.bas
'       Limit writing to EventLog table to Master PC program instance.
'   Tsg Data Warehouse.frm
'       Complete rejig of how non compliant promos are reported. Slaves will no longer write to the database.
'       Keep NonCompliantPromo records for a week and purge anything older as part of capture cycle.
'       (records are only added to table for the previous day on the current day)
'       Review all uses of Find method to ensure it is not used on an empty rst and that the cursor
'       is always positioned on the appropriate record before it is called.
'       Better handle failure to open cnn to Defaults.mdb (when in most part be when another process/program has it open)
'       Rearrange Form_Load in preparation for using Sub Main() as startup object
'       Fix bug in Market Share report (was closing rst that may not be open)
'   TsgTADO.bas
'       Fix bug in GetRst() which preventing function correctly handling an error. When an error happens the
'       function value should be returned as Nothing and the pErrMsg paramaeter should be appropriately populated.
'       GetCnn() modified so it doesn't open ALL connections in share mode when run in the IDE.
'       (commented out to assist testing in the migration of TsgDw to MySQL)
'   clsBataRpt.Cls
'       Fix bug in AddRpts_Unsent() where thorough search for second Bata rpt type not working
'       because was no SkipRecords parameter on the second call to Find method
' 23May13 3.7.0
'   Fix runtime bug in FlagExpiredPromos() which is called from tmrAutoCaptureData_Timer.
'   Was missing a trailing bracket in an SQL string.
' 27 May13 2:10pm 3.7.1
'   Review of all SQL
'   - inline concatenation of Where Conditions replaced with helper procs: MsSqlDate & SqlQuote
'   - removal of SELECT TOP ... (Jet/Ms SQL extension)
'   - use new ADOFindMethodDate() proc to create Criteria for ADO Find method
'   - replace Jet/Ms specific wild-cards in LIKE conditions and quote the conditions with SqlQutoe
'   Add new dialog form preceding Capture Selected Frans so user can elect to update
'    NonCompliantFrnachises table or not. (last version introduced non-condition updating of table)
'   Use GetReordCount() in flNumberOfExistingFranchises()
' 27 May13 5:45pm 3.7.2
'   Fix runtime bug in DialupResults(): trailing bracket missing in Like Cluase
' 30May13 5:45pm 3.7.3
'
' 10Jun13 5:45pm 3.7.4
'   Change Startup object from frmTSGDataWarehouse to Sub Main
'  clsBataRpts
'   Change AddRpts_Unsent() to search for unsent rpts for dates after the most recent txns uploaded.
'   Removed pFromDate & pThorough parameters. Change Upload() to accommodate changes in AddRpts_Unsent()
'   Take GetBataRptsMinDate() from TsgDw.bas and make it private
'  frmUploadBataRpts
'   Cater for changes in clsBataRpts and tidy up
'  TsgDw.bas
'   Create Sub Main() and use as startup object
'   Remove GetBataRptsMinDate() which is now in clsBataRpts.
'  TsgDw.frm
'   Move initial code segment out of Form_Load into newly created sub Main() in TsgDw.bas
'  TsgTADO.bas:
'   Add RstTypeEnum.eReadOnlyStatic and cater for it in GetRst()
'   Add GetRstFldVal()
'
'   In general, numerous tidy ups and greate use of GetRstFldVal() & GetRecordCount()
'
' 13Jun13 3.7.5
'   Predominantly code tidy up.
' 04Jul13 3.7.6
' Accommodate running against MySQL via new field: Defaults.mdb!MySqlCnnString [fdlgGetCnnMySQL, etc]
' Change background colour of main form depending on whether connected to Mdb or MySQL
' Add table and routines for logging changes data changes to tables and start using it.
'   (limit data access to only when required)
' Add UpdateDwDbStructure() procedure for adding new tables and fields for MySQL version (haven't added Defaults.mdb!MySqlCnnString yet)
' Add DwSqlDate() (& helper functions) to create appropriate SQL date according to Database cnn.
' Reduce number of pinging attempts in IsVpnAvailable() when running in IDE.
' Remove FlagExpiredPromos() because no longer required for program function (may be required for users using the table!)
' PopulateNonCompliantLview() & LoadPromoListview() now just called as required, and conditionally
' load ctls as required according to log of last data update for tables.
' Remove all use of rstFran!FranchisePriceUpdate = True Field (Should be ZPrefixing)
' Clean up of subPopulateVersions() and associated ctl
' Removed last remaining assoication of a file variable with a hard coded file number rather than a
' number retrieved with FreeFile. [cmdPRPrint_Click()] A lot of coupled code associated with change
' Bug fix for GetRecordCount() to accommodate MySQL behaviour where every subQuery must have an alias
' Disable fdlgGetCaptureOptions.chkUpdateNonCompliants b/c no-longer relevant with new just in time
' approach to loading NonCompliant records.
'
' 09Jul13 3.7.7
' Add UpdateAppDefaultsDbStructure(pCnn() to create MySQLCnnString and call from sub Main().
' Change g.rstEventLog to pRstType:=eEditableFwdOnly and use SQL that selects from EventLog table
'   but returns no records as the record source.
'  (Changes remote MySQL access opening rst from 11 mins to 1 second)
' Added some commented out code added to gsubAddToLocalEventLog() while I experimenting.
' Prevent auto-capture running when connected to a MySQL database.
' Some tidy up of comments
'
' 09Jul13 3.7.8
'   Numerous changes but time constraints (ridiculous) prevent adding any version notes
'  Changes to: TSG DataWarehouse.frm, Application.bas, TSG DataWarehouse.bas, TsgTADO.bas,
'              bas, TsgTUtilities.bas, fdlgGetCaptureOptions.frm
'
' 09Jul13 3.7.9
' Exclude oPOS franchises from all TsgDw.exe uploads (includes gconUtilityExe)
' Add pTimeStamp  parameter to SetTableUpdateTime() for cases where call to proc is NOT immediate
' Remove commented out code in SetGlobalDwRsts()
'
' 05Aug13 3.8.0
' Fixes for 3 Email-ErrRpts since going live with MySQL databasea
' (Emails: 1. 03/08/13 1:35 PM, 2. 03/08/13 1:42 PM & 3. 05/08/13 9:17 AM)
' 1. Stop code attempting to use an archive db when running in MySQL mode. (all data now in one db)
' 2. Problem importing Batscan fixed by changing recordset type
' 3. Stop code trying to backup db to zip file when running in MySQL mode. [TSG/7th Beam responsibility]
' Enable Auto Data Capture when running against MySQL database
' Removed automatic purging of LiveData at end of capture cycle when running against MySQL database
' Change background of main form to bright green when using an Access mdb file (now the exception)
' Change SetTableUpdateTime() to use new pTimeStamp parameter rather than using machine time as at call
' Numerouse tidy ups
'
' 13Aug13 3.8.2
' Make back ground battleship grey when connecting to a MySQL db
' (background made bright green when connecting to Access db in last version)
' Clean out cmdTest_Click() and leave code for creating runtime error for testing error handling.
' Some code tidy ups
'
' 21Aug13 3.8.3
' clsBataRpts.cls:
'   Replace GetRst() with GetRstAddOnly() for tblBataREUploads rst with aim of speeding up re-uploads
' TsgTADO.bas:
'   Removed Enums from RstTypeEum had DAO type names and were used to emulate DAO Recordset Types in ADO.
' TSGDataWarehouse.frm:
'   Replaced use of qryStock with Stock table in fbProductIsIncludedInThisProductReport() which is called from reports code.
'   Replaced DAO emulating RstTypeEnums in GetRst() calls in  subTransferAllPreLiveDataToTheLiveDataTable()
'   Get rid of silly MsgBox regarding how many pending uploads by altering LoadUploadTab() AND displayUploadsPending()
'   Now done only in a label and label is correctly updated as changes are made
'
' 22Oct13 3.8.4
'   Enable editing of stock description field on the stock tab
'   Simplified program to run only for a MySQL Data warehouse database
'    - removed conditional code for running program against either MySQL or Access database
'      (removed g.bMySQL and all conditional code running for 'Not g.bMySQL')
'    - removed different back color displays according to what type of db program was running against
'    - Fixes to gsubAddToLocalEventLog (put into the desired folder and check properly for VBA error
'    - Multitudes of cleanups and rearrangement
'
' 28Oct13 3.8.5 - unreleased version
' Revision Ver not updated in vbp file but projecgt labelled in VSS
' Only changes are:
'   - removal of commented out code
'   - changes to comments
'   - renaming of variables from rstDyn... to rst...
'
' 28Jan14 3.8.6
'   Added selected franchise(s) promotions (EXTENSIVE CHANGES) while maintaing
'   compatability with old style of distributing promotions.
'   PurgeEventLog() rewritten to work more efficiently and hopefully eliminate intermittent bug reported when
'    using old algorithm that used a recordset
'   Optimised AddRpts_Unsent()/BataRpts for speed. Hard coded earliest date to report from if no
'   previous reports have been uploaded as being 12O days from today.
'   GetRstFldVal rewritten as GetRstVal() which accommodates problems uing Max() and Min() functions on
'   selection criteria selecting no rows AND enforces that SQL provided returns a single field. (-> efficient & concise)
'   frmUploadBataRpts: Minimum date of From date control set to earliest data in LiveData table rather than 01Jan200 as
'     it was with Access mdb Version that had a LiveDataArchive table. (MySQL db conversion left the Archive table behind)
'     Renamed filename (cf object name) from frmUploadBataRpts.frm to UploadBataRpts.frm.
'   Add optional parameter pDataFld to LoadCombo_Rst()
'   Renamed CheckboxToBoolean() to ChkBoxToBool()
'   Renamed BooleanToCheckbox() to BoolToChkBox()
'   Add Tx() (& helper fn) for implementing nested transaction calls with data providers that don't them (eg MySQL)
'   Add GetCollectionFromRst()
'   Add optional parameter pDataFld to LoadCombo_Rst()
'   GetRstFldVal rewritten as GetRstVal() which accommodates problems uing Max() and Min() functions on selection
'   criteria applying to no rows, AND enforces that SQL provided returns a single field. (-> efficient & concise)
'   GetCnnMySql()
'       1. Hard code CnnMode as it has no effect on {MySQL ODBC 5.1 Driver}
'       2. Change cnn string options to so multiple SQL statements separated by a semi-colon can be used.
'       3. Add significant comments to module header for GetCnnMySql()
'   Moved UploadFilesToOneFranchise() higher up in SubCaptureData() to where it always should have been
'   and wasn't, and have the feeling this may fix some rarely occuring difficult to explain bugs.
'   Modify fOKToUploadItem() to better accommodate oPOS stores and optimise the procedure at the same time
'   Remove some unused controls on settings tab such as txtIOMEGADrive
'
' 30Jan14 3.8.7
' Fixed bugs:
'   1. promotions for nominated fran(s) would not upload
'   2. promotions for nominated fran(s) were  writing FranProm records for most Frans but not for Unicenta Frans.
' Neatened up layout of fdlgGetCnnMySQL.
' Made fdlgGetCnnMySQL.cmdOK the default button.
' Fixed bug when passing pValString:=vbNullString to SetValStringValue(). Fixing it also fixed bug adding (cf editing) a Name=Value pair to a value string. Showed up when using a new defaults.mdb and connecting to MySQL db with fdlgGetCnnMySQL.
' Added Czls() - for use in a bug fix for uploading nomintated Franchise promotions.'----------------------------------------------------------------------------------------------------------------------------------
'
' 30Jan14 3.8.8
' Fixed bugs:
'   1. recall promotions for nominated fran(s) not working (FP!TfrStatus not being set on upload)
'   2. when uploading a new promotion exclude promos for nominated franchise from being checked for change in PromoGrade (not previously done)
'   3. runtime bug in PrintPromoList (missing trailing bracket in a Where Clause)
' Add comments flagging areas for review
'
' 17Feb14 3.8.9
' Fix bug with uploading Bata reports. 'rstFran.MoveNext' was inadvertantly removed from
'  Sub AddRpts_Unsent() when procedure was rewritten in V386 -> reinstated 'rstFran.MoveNext'
' Added TfrPosLiveToPreLive() - called from a button on Data Capture tab
' Added CnnDwExecute() and replaced all calls to g.cnnDW.Execute calls with this function.
'   (CnnDwExecute() creates a runtime error a SLAVE instance tries to Create, Update or Delete records)
'    Rsts created with GetRst() have for a while been read-only for SLAVE instances to ensure the same behaviour preventing
'    SLAVE instances from modifying data.
' fdlgGetCnnMySQL: Change 'password char' (ie mask char) to standard password char symbol.
'
' 21Feb14 3.9.0
'   Command line call to MySqlCheck at end of CaptureCycle (AutoCaptureCycle) to optimise TsgDw db.
'   Call via a batch file created on the fly. Procedures added: ExecuteAndWait() and OptimiseDB()
'
' 27Feb14 3.9.1
'   Fixed bug in OptimiseDb() where program went to open hardcode Fullfilename for OptimiseDwDatabaseOutfile.txt
'  (which worked on my machine b/c of my folder structure) instead of using Fullfilename variable
'   Significant changes to AddPromo() &  AddPromoToUploads() and minor changes to SaveNewPromotion()
'   to better cater for the cases where a Promo combination of State, Region and PromoGrade does not
'   apply to any current franchises. Changes make code easier to understand and different behaviour
'   can be coded more easily once the behaviour is determined in consultation with TSG.
'   rstFranIDs seln in AddPromo() further refined so only appropriate records selected to avoid unnecessary
'   calls to AddPromoToUploads() which previously eliminated these recrords with a call to fOKToUploadItem()
'
' 19Aug2014 3.9.2
'   Change Aztec uploads to use SFTP (as per BATA). Required TSG upgrading CuteFTP from version 8 to 9
'   Numerous and signiificant code tidy ups since last version required b/c of chnages to library code
'   when developing program for upgrading RStats sites to Unicenta oPOS
'   May have been a number of bug fixes. Did fix a bug in clsSFTP when using FTPS (no live code has yet used FTPS)
'
' 09Oct2014 3.9.3 ~~~~~~~~~~~~~
'   Changed GetCnnMySql() from adUseServer to adUseClient.
'
' 23Oct2014 3.9.4 ~12:03pm
'   Reverted GetCnnMySql() to adUseServer (Txns can only be used with a server side connection)
'   subTransferAllPreLiveDataToTheLiveDataTable() renamaed to TfrAllPreLiveDataToLiveData() and rewritten
'   TfrAllPreLiveDataToLiveData optimised to (1) batch inserts for livedata, rejects and duplicates table
'   into statements for multiple records rather than individual statements for each record (2) cache
'   values from a row in the PreLive table in an array rather than interrogating rst for each fld value.
'   (3) a number of procedures called from TfrAllPreLiveDataToLiveData() changed to assist optimisation
'   Numerous code clean ups.
'
' 23Oct2014 3.9.5 ~5:20 pm
'   BUG FIX: Alter TfrAllPreLiveDataToLiveData() to cater for REMOTE field in PreLiveData table
'            Field was not present in my development db
'
' 29Oct2014 3.9.6
'   LiveDataArchive table added to db. LiveDataArchive contains a complete copy of all data and LiveData
'   is purged by PurgeLiveData() in subCaptureData() according to setting in Defaults!DaysOfLiveData.
'   (I consider this as a retrograde change, but it was requested and I am complying)
'
'   BugFix: GetValList([ByVal|ByRef] pSrcFldCollection As VBA.Collection, ...)
'           Regardless of how pSrcFldCollectioneven is passed any new item added the collection will
'           be available as data to the object reference in the calling procedure. Under particular
'           circumstances this caused a runtime error so the bug was fixed.
'
' 10May2015 3.9.7
'- In response to request to fix bug
'   Removed g.rstEventLog and rewrote gsubAddToLocalEventLog() to execute SQL rather than use a rst.
'   gsubAddToLocalEventLog() rewritten to uncover bugs being masked and trapped in a loop.
'   CnnDwExecute() rewritten to log failed attempts at executing SQL against the database cnn to the
'   error log text file that until now was only being written to by gsubAddToLocalEventLog().
'   Now any problem gsubAddToLocalEventLog() has writing records will be logged by this new common code.
'   Faulty logic in uncovered in both gsubAddToLocalEventLog() and CnnDwExecute() rewritten to assist
'   finding and fixing bugs.
'   Fixed incorrect trimming of returned string from dll call to GetComputerName that became apparent
'   when using SQL in gsubAddToLocalEventLog(). When a NulChar was embedded in SQL it caused a problem
'   which I think occurred because when crossing processing boundaries when submitted to MySQL the char
'   was seen as the termination of the string.
'- Enhancement requests
'   Removed VPN availabilty test at startup. Removed label warning when VPN was not available.
'    (Not needed now most frans have their own network arrangements and push their data to H/O)
'   Filtered out PreliveData with barcodes > 13 chars and sent it to RejectData table
'
'   Numerous tidy ups while in the code
'
' 12May2015 3.9.8 (Unfortunately version not labelled in VSS so can't be retrieved from VSS
'   but instrumented code version can be retrieved from a zipped file of instrumented code)
' Increase barcode length of rejected records in TfrAllPreLiveDataToLiveData() from 13 to 15
' Rename error log file in CnnDwExecute from EventLog_ERRORS.txt to SqlExecution_ERRORS.txt".
' Add an extra vbNewLine into its companion MsgBox
'
' 23May2015 3.9.9
'- Changed summarising of OPosLive for transfer to PreLive. Now grouped by FranID, Date & Barcode only
'   whereas it previously also included NormalSellInc and CostInc. The previous grouping
'   created multiple sumamries for Fran/Date/Barcode combinations because these price fields
'   would vary slightly in OPos according to the qty sold. Program would then only transfer the first
'   summary for the Fran/Date/Barcode combination from PreLive to Live.
'- Changed filtering out PreliveData with barcodes > 13 chars to filtering data with barcodes > 15
'
' 03Jul2015 4.0.0
'   Fixed bug in PromotionRecall() where in some circumstances a test would incorrectly determine a
'   promotion as an OldStyle/legacy promotion and call PromoRecallOldStyle(). {DefaultValue in a call
'   to GetRstVal() should have been something other than False as records could have a valid value of 0}
'   We are now well beyond any OldStyle promotions persisting so PromoRecallOldStyle() and procedures
'   it called {AddPromoToUploadsPendingOldStyle(), DeletePromoFromUploadsPending()} were removed.
'   Renamed AddPromo to AddNewPromo to more accurately reflect actions of the function
'   Renamed AddPromoToUploads() to AddPromoToFUandFP() to more accurately reflect actions of the function
'   Remove kludge retricting uplodaing promo data (New Promos or Recalls) to complete DataCapture cycles
'   (cf selected franchise data capture). In any case, I confirmed you could and can still upload promo
'   data via the uploads tab. The need for this kludge was obviated with the rewrite of promos for oPOS
'   The original flawed design inherited could still do with some more rewriting but the need for thet
'   particular kludge has gone. Uploading promo data to frans from this program is no longer controlled
'   by Promotions!PromoStatus but the field is maintained for informational purposes and communication
'   with the oPOS suite of software
'
'  13Jan16 4.0.1 - (Test version)
'  14Jan16 4.0.2 - Releas version
'
'  20Jan16 4.0.3
'   Fixed bug where 'Upload This NOW' uploaded successfully, but failed to correctly update FranchiseUploads
'   B/c FranchiseUploads not updated correctly, the program would attempt same upload in O/N capture cycle.
'   Code fix made in UploadFilesToOneFranchise()
'
'  DdMmmYy 4.n.n -
'
'---------------------------------------------------------------------------------------------------------
' DdMmmYy 3.6.?
'' Next version:
' - Consider Purge routines for tblBataUploads & tblBataReUploads
' - RE: Public Sub ZipFiles() - When creating Version 360 Noticed that zipped files are now uploaded
'   automatically by the programm to AZTEC. To be safe this procedure should probably be converted to making
'   a synchronous call to ensure that zipping of files is always complete before upload is attempted.
'   Since introduction of automatic uploading of files to AZTEC we have relied on the zipping of daily rpts
'   being completed while weekly rpts were being created, and the upload of daily rpts while weekly rpts
'   are being zipped for everything to work Ok.
' - perhaps improve deleting stock by using a check box rather than a button (Stock Tab)
' - progressively move code confiugring Stock tab ctls into ConfigureStkTabCtls()
' - search for tblFranchisePromotions and fully utilise all the wonderful vistas it opens
'   (like uploading appropriate 'new promos' and 'promo recalls' each time)
'   (a connection is made to a Fran rather than only on Capture All events)
'---------------------------------------------------------------------------------------------------------
'COULD SAVE LOTS OF PROCEDURES BY REMOVING ALL THE SPIN CONTROLS AND ALL THE RIDICULOUS SINGLE LINE PROCEDURES THAT ENABLE
'CONTROLS ON THE FRANCHISE TAB WHEN THERE IS A DOUBLE CLICK POSSIBLY MAKE A BATA FUNCTION MODULE OR EVEN MOVE INTO CLASSES!
'
' MASTER/SLAVE SETTING IS A SETTING THAT COULD BE LEGITIMATELY STORED IN THE REGISTRY RATHER THAN DEFAULTS MDB
'---------------------------------------------------------------------------------------------------------
'   Could add automatic purging of promotions which are not currently automatically purged
'   Also the text field sizes could be greatly reduced for the promotions table.
'---------------------------------------------------------------------------------------------------------

Public Type udtDisplayUpdates
'   VB-Watch requires User Defined Type to be declared Public when used in a Public variable
    dtmPromoListView As Date
    dtmNonCompliantLView As Date
End Type

Public Type udtGlobal
'   VB-Watch requires User Defined Type to be declared Public when used in a Public variable
    bMaster As Boolean
''' bVpnAvailable As Boolean    ''' V397
    bBusinessMgrPC As Boolean
    bAutoDataCapture As Boolean
    bCaptureCycleRunning As Boolean
    dtmLiveDataStart As Date        ' Set to oldest TxDate in LiveData (Previously Earliest live data in TsgDw.mdb (earlier data kept in archive mdb)
                                    ' V386: USED IN A FEW PLACES AND WITH CAREFULL REWRITING COULD POSSIBLY BE REMOVED
    lngEventLogEventFldSize As Long ''' V397
    lngEventLogFranFldSize As Long  ''' V397
    strAppDrive As String           ' like C:\ or \\server\somewhere\other
    strAppRoot As String            ' g.strAppDrive plus gCompanyIdentifier, eg: C:\TS, eg: \\server\etc\sub\TS
    strLogFolder As String
    strRptsFolder As String
    strBatscanFolder As String
    strUploadsFolder As String
    strNielsenRptsFolder As String
    strBataRptsFolder As String     ''' V401
    strLocalMessageFolder As String
    strTsTemp As String ' Holds the InHQButNotInAndy.txt" & "AndyDiffHQ.txt" files
                        ' that created when merging databases via cmdMerger_Click()
    strNodeType As String
    strNodeName As String
    strPkZipCExe As String
    udtDtmCtlUpdated As udtDisplayUpdates
    cnnDW As ADODB.Connection           ' DataWarehouse database (main/application database)
    rstDWDefaults As ADODB.Recordset    '!!! ManualFix Clearing: Object WILL REQUIRE clearing
''' rstEventLog As ADODB.Recordset ''' V397 ' ONLY USED IN gsubAddToLocalEventLog(); global kept for efficiency-Could try leaving Jet Engine to look after caching?
    rstAppDefaults As ADODB.Recordset   ''' Review Chould be persisted as an XML file or defaults
                                        '''        stored as an ini file or in the registry and should possibly
                                        '''        be renamed to rstPcDefaults/rstMachineDefaults/InstallDefaults/...
                                        '''        Also, if browse to database, [NodeType, ...] was removed
                                        '''        there would be no editing and the values could simply
                                        '''        be read into a UDT- the least resource hungry approach
End Type
Public g As udtGlobal ' VB-Watch requires User Defined Type to be declared Public when used in a Public variable

''' Event logged with this string for Franhcise value can later be filtered out for display
''' Public Const gkEventLogWarning As String = "WARNING"
''' EVEN BETTER WOULD BE AN EXTRA FLAG FIELD IN EVENT LOG TO DETERMINE STATUS
''' (EG Normal, Warning, Error, Failure, ....)
                                                        
Public Type udtError
    Description As Variant
    HelpContext As Variant
    HelpFile As Variant
    LastDllError As Variant
    Number As Variant
    Source As Variant
    ErrLine As Variant
End Type

'-------------------------------------------------------------------------------------------------------------------------------------------------
'general variables
    Global gCompanyIdentifier As String ' like TS
    Global gbClickEventIsSuppressed As Boolean
    Global gbEventLogRefreshIsEnabled As Boolean
    Global gbEventLogRefreshIsNotAlreadyInProgress As Boolean   ''' Should be useless but code may allow buttons to be clicked while other code is executing
    Global gvListItem As Variant
    Global glMaximumFranchises As Long
    Global gsProductReportPathAndFilename As String
    Global gsReportPeriodWording As String
    Global gsStickReportPathAndFilename As String
    Global giTopSellers As Integer
    Global gfFutureDate As Boolean
    Global gfDateFormatBad As Boolean ''' Review Fix up this crap when time permits (Date fmt reliance has been removed from Tsg RStats)

    'data access
    
'   CONSTANTS
'   ---------
    'general constants
    Public Const gcon22DigitDecimalFormat As String = "#####################0.#0"
    Public Const gcon5DigitDollarFormat As String = "$####0.#0"
    Public Const gcon6DigitDollarFormat As String = "$#####0.#0"
    Public Const gconAllFranchises As String = "(All franchises)"
    Public Const gconContactSystemAdministrator As String = "Contact the System Administrator"
    Public Const gconDoNotDisplayAnyItems = -1
    Public Const gconNielsenFilePrefix As String = "TSG"
    Public Const gconDisplayFirstItem = 0
    Public Const gconOtherSuppliers = 1
    Public Const gconFmtDateInNielsenFilename As String = "ddmmyy"
    Public Const gconReportManager As String = "Report Manager"
    
    Public Const gconReservedFranchiseID As Long = 99999

    Public Const gconSpace As String = " "
    Public Const gconStandardDateFormat As String = "ddMmmyy"
    Public Const gconStandard4x3Format As String = "###0.##0"
    Public Const gkFmtDateTime As String = "ddMmmyy hh:mm:ss"
    Public Const gconStandardQuantityFormat As String = "#####"
    Public Const gconTopSellersDefault As Long = 30
    Public Const gconTruncateCharacter As String = "..."
    Public Const gconTruncateDescriptionBriefAt As Long = 8
    Public Const gconTruncateExtensionWidth As Long = 4
    Public Const gconZeroValue As Long = 0
    Public Const gconTextFileSuffix As String = ".txt"
    Public Const gconNewMessageFilePrefix As String = "TSMsg"
    Public Const gconNewStockFilePrefix As String = "TSStk"
    Public Const gconWLPUpgradePrefix As String = "TSWLP"
    Public Const gconUpdateStkFldsUpdatePrefix As String = "TSSfu" 'TSSFU = TS Stock Field(s) Update

    Public Const PROMO_NOT_SENT As String = "not uploaded yet"
    Public Const PROMO_SENT As String = "uploaded"
    Public Const PROMO_RECALLED As String = "recalled"
    
    Public Const DATEFORMATBAD As String = "DATEFORMATBAD"

    ' TSG Specific directories & filenames
    Public Const gconProductReportFilename As String = "ProductReport.txt"
    Public Const gkRemoteDbFilename As String = "RemoteStatistics.mdb"
    Public Const gconStickReportFilename As String = "StickReport.txt"

    ' genericize
''' Public Const gconEventlogError As String = "EventLog_ERRORS.txt"    ''' V398

    Public Const gsNewRSexe As String = "NewRemoteStatistics.exe"
    Public Const gconPriceModule As String = "NewPriceModule.exe"

    Public Const gconUpgradeRS As String = "upgradeRemoteStatistics.exe"
    Public Const gconUtilityExe As String = "utility.exe"
    Public Const IMMEDIATELY = 0
    Public Const LATER = 1
    Public Const ALL_UPLOADS = 0
    Public Const CURRENT_UPLOADS = 1

    'tayble naymz
    Public Const gconLiveDataTable As String = "LiveData"

    'stok tayble
    Public Const gconStockTableStockIDField As String = "stock_id"
    Public Const gconStockTableBarcodeField As String = "Barcode"
    Public Const gconStockTableBarcodeFieldWidth As Long = 15
    Public Const gconStockTableCustom1Field As String = "custom1"
    Public Const gconStockTableCustom2Field As String = "custom2"
    Public Const gconStockTableSalesPromptField As String = "sales_prompt"
    Public Const gconStockTableAllowFractionsField As String = "allow_fractions"
    Public Const gconStockTablePackageField As String = "package"
    Public Const gconStockTableTaxComponentsField As String = "tax_components"
    Public Const gconStockTableDescriptionField As String = "description"
    Public Const gconStockTableLongDescriptionField As String = "longdesc"
    Public Const gconStockTableCategoryField As String = "cat1"
    Public Const gconStockTableSubCategoryField As String = "cat2"
    Public Const gconStockTableGoodsTaxCodeField As String = "goods_tax"
    Public Const gconStockTableCostField As String = "cost"
    Public Const gconStockTableSalesTaxCodeField As String = "sales_tax"
    Public Const gconStockTableSellField As String = "sell"
    Public Const gconStockTableSupplierIDField As String = "supplier_id"
    Public Const gconStockTableSticksField As String = "sticks"
    
    'remote defoltz taybl
    '   NB: This definition also exists in remotestatistics, so modify hand-in-hand
    Public Const gconRemoteDefaultsTableDatabaseOpenedByField As String = "DatabaseOpenedBy"
    Public Const gconRemoteDefaultsTableUpgradeField As String = "Upgrade"
    Public Const gconRemoteDefaultsTableRunUtility As String = "spare2"
    Public Const gconRemoteDefaultsTableMessageFlag As String = "MessageFlag"           ' PAL
    Public Const gconRemoteDefaultsTableNewStockFlag As String = "NewStockFlag"         ' PAL
    Public Const gconRemoteDefaultsTableWLPUpdateFlag As String = "WLPUpdateFlag"       ' PAL
    
    'warehouse application defoltz taybl
    
    ' defaults.mdb fields
    'CompanyIdentifier
    'StatisticsDatabase
    Public Const gconDocketPrinterEnabled As String = "DocketPrinterEnabled"
    'NodeName
    'ModemName
    'MasterStatus
    'NetworkPrinterEnabled
    
    'warehouse franchise details tayble
    Public Const gconFranchiseTableTSGFranchiseIDField As String = "FranchiseIDTSG"
    Public Const gconFranchiseTableBusinessNameField As String = "FranchiseBusinessName"
    Public Const gconFranchiseTablePhysicalAddressSuburbAndPostcodeField As String = "FranchisePhysicalAddressSuburbAndPostcode"
    Public Const gconFranchiseTableContactNameField As String = "FranchiseContactName"
    Public Const gconFranchiseTableAreaCodeField As String = "FranchiseAreaCode"
    Public Const gconFranchiseTableStateOfOz As String = "FranchiseStateOfOz"
    Public Const gconFranchiseTablePhoneField As String = "FranchisePhone"
    Public Const gconFranchiseTableModemField As String = "FranchiseModem"
    Public Const gconFranchiseTableFaxField As String = "FranchiseFax"
    Public Const gconFranchiseTableNodenameField As String = "FranchiseNodename"
    Public Const gconFranchiseTableRASPasswordField As String = "FranchiseRASPassword"
    Public Const gconFranchiseTableRemoteModuleVersionField As String = "FranchiseRemoteVersion"

    'warehouse lyve data Table
    Public Const gconLiveDataTableTSGFranchiseIDField As String = "FranchiseIDTSG"
    Public Const gconLiveDataTableBarcodeField As String = "Barcode"
    Public Const gconLiveDataTableQuantityField As String = "Quantity"
    Public Const gconLiveDataTableTotalIncTaxField As String = "TotalInc"
    Public Const gconLiveDataTableNormalSellIncTaxField As String = "NormalSellInc"
    Public Const gconLiveDataTableCostIncTaxField As String = "CostInc"
    Public Const gconLiveDataTableWholesaleQty As String = "WholesaleQty"
    Public Const gconLiveDataTableWholesaleActualSell As String = "WholesaleActualSell"
    
    ' FranchiseUploads Table
    Public Const gconUploadFranchiseIDField As String = "FranchiseID"
    Public Const gconUploadFileField As String = "UploadFile"
    Public Const gconUploadDateField As String = "UploadedDate"
    
    'supplyar taybl
    Public Const gconSupplierTableSupplierIDField As String = "supplier_id"
    Public Const gconSupplierTableSupplierNameField As String = "Supplier"
    
    'ascii
    Public Const gconSpaceAscii As Integer = 32
    Public Const gconSingleQuoteAscii As Integer = 39
    Public Const gconNonDestructiveSingleQuoteAscii As Integer = 146

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Type STARTUPINFO
              cb As Long
              lpReserved As String
              lpDesktop As String
              lpTitle As String
              dwX As Long
              dwY As Long
              dwXSize As Long
              dwYSize As Long
              dwXCountChars As Long
              dwYCountChars As Long
              dwFillAttribute As Long
              dwFlags As Long
              wShowWindow As Integer
              cbReserved2 As Integer
              lpReserved2 As Long
              hStdInput As Long
              hStdOutput As Long
              hStdError As Long
     End Type

     Type PROCESS_INFORMATION
              hProcess As Long
              hThread As Long
              dwProcessId As Long
              dwThreadID As Long
     End Type

    Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
    Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
    
Public Const gkGstRare As Currency = 0.1

Public Sub CloseCnnAndSetToNothing_IgnoreErrors(ByRef pCnn As ADODB.Connection, _
                                                ByVal pLogIfAvoidsBug As Boolean, _
                                                ByVal pCalledFromTag As String)
' If you try to close a Connection while it has open Recordset objects, the
' Recordset objects will be closed and any pending updates or edits will be canceled
On Error Resume Next

    pCnn.Close
    
    If pLogIfAvoidsBug Then
        If Err.Number Then
            LogBugFix fGetErrorUdt(), pCalledFromTag:=pCalledFromTag
        End If
    End If

    Set pCnn = Nothing

End Sub

Public Sub CloseDatabaseCnns()
'   Close any open recordsets then close connection
'   ADO Close Method: Closes an open object and any dependent objects.
'   Closing a Connection object while there are open Recordset objects on the connection rolls
'   back any pending changes in all of the Recordset objects. Explicitly closing a Connection
'   object (calling the Close method) while a transaction is in progress generates an error.
'   If a Connection object falls out of scope while a transaction is in progress,
'   ADO automatically rolls back the transaction.

    g.cnnDW.Close
    Set g.cnnDW = Nothing

End Sub

Public Function CnvNulls(ByVal pValue As Variant, ByVal pReplaceWith As Variant) As Variant
    If IsNull(pValue) Then
        CnvNulls = pReplaceWith
    Else
        CnvNulls = pValue
    End If
End Function

Public Function CnvZerosToNull(ByVal pValue As Variant) As Variant
    If pValue = 0 Then
        CnvZerosToNull = Null
    Else
        CnvZerosToNull = pValue
    End If
End Function

Public Function DeleteFile(ByVal pFileToDelete As String) As Boolean
'    Sub DeleteFile(ByVal pFileToDelete As String)
'   Returns True if file is deleted or didn't exist and False otherwise
'   (or put another way Returns false if there is a file to delete but it can't delete it and True otherwise)
'   (Returns True if file didn't exist to maintain compatability )
'   If the file doesn't exist procedure will return False
Dim bResult As Boolean

On Error GoTo Procedure_Error

    If Dir(pFileToDelete) = "" Then
        bResult = True
    Else
        SetAttr pFileToDelete, vbNormal
        Kill pFileToDelete
        bResult = True
    End If
    
Procedure_Exit:
    DeleteFile = bResult
    Exit Function
    
Procedure_Error:
    bResult = False
    Resume Procedure_Exit
    
End Function

Public Function DoubleQuote(ByVal pExp As Variant) As String
'   Return passed expression enclosed in Double Quotes
    DoubleQuote = """" & pExp & """"
End Function

Public Function ExecCmd(cmdline$, Optional ByVal pTimeoutInMinutes As Long = -1)
''Public Function ExecCmd(cmdline$, ByVal fNoWait As Boolean, Optional ByVal pTimeoutInMinutes As Long = -1)
'   AUrban V307 pTimeoutInMinutes parameter added to make timeout choice a considered decision
'   fNoWait superceded by the Timeout in minutes which can be set to zero if not wanting to wait
'   Default value of pTimeoutInMinutes -1 gives an infinite timeout
'   (Help for this function defines a constant of INFINITE = -1
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ret As Long
Dim lngTimeout As Long
    
    If pTimeoutInMinutes <> -1 Then
        lngTimeout = CLng(1000) * 60 * pTimeoutInMinutes
    End If
    
'   Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    
'   Start the shelled application:
    
    ret = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)
    
    ' Wait 20 minutes for the shelled application to finish:
    Call WaitForSingleObject(proc.hProcess, lngTimeout)   ' INFINITE)
    Call GetExitCodeProcess(proc.hProcess, ret&)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    ExecCmd = ret
    
End Function

Public Function ExecuteAndWait(cmdline As String) As Long
Const INFINITE As Long = -1&
Dim lngReturn As Long
Dim NameOfProc As PROCESS_INFORMATION
Dim NameStart As STARTUPINFO
   
'   Initialize the STARTUPINFO structure:
    NameStart.cb = Len(NameStart)

'   NameStart the shelled application:
    lngReturn = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, _
                               0&, vbNullString, NameStart, NameOfProc)

'   Wait for the shelled application to finish:
    lngReturn = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
    Call GetExitCodeProcess(NameOfProc.hProcess, lngReturn)
    Call CloseHandle(NameOfProc.hThread)
    Call CloseHandle(NameOfProc.hProcess)
    ExecuteAndWait = lngReturn
      
End Function

Public Function fdtmLastSunday() As Date
'   Returns: Last Sunday (not today if today is a Sunday)
'    Nielsen reports run from Monday to Sunday
'    They are generated each Monday up to fdtmLastSunday
Dim dtmTemp As Date

    dtmTemp = Date
    
    Do
        dtmTemp = DateAdd(Interval:="d", Number:=-1, Date:=dtmTemp)
    Loop Until Format$(dtmTemp, "dddd") = "Sunday"

    fdtmLastSunday = dtmTemp

End Function

Public Function fdtmYesterday() As Date
    fdtmYesterday = DateAdd(Interval:="d", Number:=-1, Date:=Date)
End Function

Public Function fGetErrorUdt() As udtError
'!  vbwNoErrorHandler vbwNoTraceProc !
'!  DO NOT ALTER ABOVE LINE - INCLUDES TAG FOR VB WATCH ERROR HANDLING AND MUST BE FIRST LINE IN THE PROCEDURE !

'   This function needs to be shielded from VB Watch Error Handling so that error details are preserved
'   VB Watch tags stop VB Watch inserting error handling that would reset the
'   Err object and clear the information held by it
'   TraceProc is turned off as well as ErrorHandling since procedures called
'   by Tracing Procedure have error handling and therefore reset the Err object

Dim udtErrorDetails As udtError

    With Err
        udtErrorDetails.Description = .Description
        udtErrorDetails.HelpContext = .HelpContext
        udtErrorDetails.HelpFile = .HelpFile
        udtErrorDetails.LastDllError = .LastDllError
        udtErrorDetails.Number = .Number
        udtErrorDetails.Source = .Source
        udtErrorDetails.ErrLine = Erl
    End With

    fGetErrorUdt = udtErrorDetails

End Function

Public Function fGetLastWord(ByVal strSentence As String, ByVal sDelimiter As String) As String
    
    Dim strLastWord As String
    Dim trimmedSentence As String
    Dim a As Integer
    Dim intLetterCount As Integer
    Dim intCountDown As Integer
    
    trimmedSentence = Trim(strSentence)
    intLetterCount = Len(trimmedSentence)
    intCountDown = intLetterCount
    For a = 1 To intLetterCount
        strLastWord = Mid(trimmedSentence, (intCountDown), 1)
        If strLastWord = sDelimiter Or strLastWord = Null Then
            Exit For
        End If
        intCountDown = intCountDown - 1
    Next a
    fGetLastWord = Right(trimmedSentence, intLetterCount - intCountDown)
 
End Function

Public Function fsErrDetail(ByVal pProcName As String, Optional ByVal pExtraMsg As String = vbNullString) As String
'!  vbwNoErrorHandler vbwNoTraceProc !
'!  DO NOT ALTER ABOVE LINE - INCLUDES TAG FOR VB WATCH ERROR HANDLING AND MUST BE FIRST LINE IN THE PROCEDURE !

'   VB Watch tags stop VB Watch inserting error handling that would reset the
'   Err object and thus clear information in it required by this procedure
'   TraceProc is turned off as well ErrorHandling since procedures called
'   by Tracing Procedure have error handling and therefore reset the Err object
Dim strResult As String
Dim strLineNo As String
Dim strSource As String
Dim strExtraMsg As String

    If Len(pExtraMsg) <> 0 Then
        strExtraMsg = " (" & pExtraMsg & ")"
    End If
    
    If Erl Then strLineNo = "line " & Erl
    If Len(Err.Source) Then strSource = "Source=" & Err.Source
'   If Err.LastDllError Then strLastDll = " LastDllError " & Err.LastDllError

    strResult = pProcName & " " & _
                strLineNo & ": " & _
                strSource & " " & _
                "Err " & Err.Number & ": " & _
                Err.Description & strExtraMsg

'   fsVersion() called post collecting error information b/c error handling in fsVersion() clears Err object
    strResult = "Version " & fsVersion() & " " & strResult
    
    fsErrDetail = strResult

End Function

Public Function fsYesterdaysDate() As String
    fsYesterdaysDate = Format$(DateAdd(Interval:="d", Number:=-1, Date:=Date), gconStandardDateFormat)
End Function

Public Function GetDate_FromTsgDateFld(ByRef pFld As ADODB.Field) As Date
'   pFldVal is passed ByRef for speed
'   A TsgDateFild is either old format (string of ddMmmYy),
'   or a date type to which Tsg programs are migrating
Dim dtmResult As Date
   
    If pFld.Type = ADODB.DataTypeEnum.adDate Then
        dtmResult = pFld.Value
    Else
        dtmResult = CnvDdMmmYyToDate(pFld.Value)
    End If

    GetDate_FromTsgDateFld = dtmResult

End Function

'Public Sub subIgnoreErr_ClearDbOpenedbyFld(ByRef prstRemoteDefaults As DAO.Recordset, ByVal pLogIfAvoidsBug As Boolean, ByVal pCalledFromTag As String)
'
'On Error Resume Next
'
'    With prstRemoteDefaults
'        .Edit
'            .Fields!DatabaseOpenedBy = vbNullString
'        .Update
'    End With
'
'    If pLogIfAvoidsBug Then
'        If Err.Number Then
'            LogBugFix fGetErrorUdt(), pCalledFromTag:=pCalledFromTag
'        End If
'    End If
'
'End Sub

'Public Sub subIgnoreErr_CloseRst(ByRef pRst As DAO.Recordset, ByVal pLogIfAvoidsBug As Boolean, ByVal pCalledFromTag As String)
''   Initially written to avoid errors when closing recordsets within error handlers of subCaptureData
''   Errors encountered within the error handler were causing the O/N runtime errors
''   To facilitate logging when this procedure averts a run-time error the pLogError parameter was added
''   Where this procedure replaces in-line error handling the pLogError parameter is passed as False
'
'On Error Resume Next
'
'    pRst.Close
'
'    If pLogIfAvoidsBug Then
'        If Err.Number Then
'            LogBugFix fGetErrorUdt(), pCalledFromTag:=pCalledFromTag
'        End If
'    End If
'
'    Set pRst = Nothing
'
'End Sub

Public Function GetDateFrom_dd_mm_yyyy(ByVal pdd_mm_yyyy As String, ByRef pIsDate As Boolean) As Date
Dim strDd As Variant
Dim strMm As Variant
Dim strMmm As Variant
Dim strYyyy As Variant

    If Len(pdd_mm_yyyy) = 10 Then
        strDd = Left(String:=pdd_mm_yyyy, Length:=2)
        strMm = Mid(pdd_mm_yyyy, start:=4, Length:=2)
        strYyyy = Right(pdd_mm_yyyy, Length:=4)
        
        If IsEachCharADigit(strDd) Then
            If IsEachCharADigit(strMm) Then
                If IsEachCharADigit(strYyyy) Then
                    Select Case CInt(strMm)
                        Case 1: strMmm = "JAN"
                        Case 2: strMmm = "FEB"
                        Case 3: strMmm = "MAR"
                        Case 4: strMmm = "APR"
                        Case 5: strMmm = "MAY"
                        Case 6: strMmm = "JUN"
                        Case 7: strMmm = "JUL"
                        Case 8: strMmm = "AUG"
                        Case 9: strMmm = "SEP"
                        Case 10: strMmm = "OCT"
                        Case 11: strMmm = "NOV"
                        Case 12: strMmm = "DEC"
                        Case Else
                            strMmm = "ZZZ"
                    End Select
                    
                    If IsDate(strDd & " " & strMmm & " " & strYyyy) Then
                        pIsDate = True
                        GetDateFrom_dd_mm_yyyy = DateSerial(Year:=CInt(strYyyy), Month:=CInt(strMm), Day:=CInt(strDd))
                    End If
                End If
            End If
        End If
    End If
    
End Function

Public Function GetDateFrom_ddmmmyy(ByVal pDdMmmYyyy As String) As Date
Dim intDay As Long
Dim intMonth As Long
Dim intYear As Long

    intDay = Val(Left$(pDdMmmYyyy, 2))
    intYear = Val(Right$(pDdMmmYyyy, 2))
    Select Case VBA.UCase$(Mid$(pDdMmmYyyy, 3, 3))
        Case "JAN": intMonth = 1
        Case "FEB": intMonth = 2
        Case "MAR": intMonth = 3
        Case "APR": intMonth = 4
        Case "MAY": intMonth = 5
        Case "JUN": intMonth = 6
        Case "JUL": intMonth = 7
        Case "AUG": intMonth = 8
        Case "SEP": intMonth = 9
        Case "OCT": intMonth = 10
        Case "NOV": intMonth = 11
        Case "DEC": intMonth = 12
    End Select
    
    GetDateFrom_ddmmmyy = CDate(DateSerial(Year:=intYear, Month:=intMonth, Day:=intDay))

End Function

Public Function GetDateFrom_yyyymmdd(ByVal pYyyyMmDd As String, ByRef pIsDate As Boolean) As Date
Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim strMonth As String

    If Len(pYyyyMmDd) = 8 Then
        If IsEachCharADigit(pYyyyMmDd) Then
            intYear = CInt(Left$(pYyyyMmDd, 4))
            intMonth = CInt(Mid$(pYyyyMmDd, start:=5, Length:=2))
            intDay = CInt(Right$(pYyyyMmDd, Length:=2))
            
            Select Case intMonth
                Case 1: strMonth = "JAN"
                Case 2: strMonth = "FEB"
                Case 3: strMonth = "MAR"
                Case 4: strMonth = "APR"
                Case 5: strMonth = "MAY"
                Case 6: strMonth = "JUN"
                Case 7: strMonth = "JUL"
                Case 8: strMonth = "AUG"
                Case 9: strMonth = "SEP"
                Case 10: strMonth = "OCT"
                Case 11: strMonth = "NOV"
                Case 12: strMonth = "DEC"
                Case Else
                    strMonth = "ZZZ"
            End Select
    
            If IsDate(intDay & " " & strMonth & " " & intYear) Then
                pIsDate = True
                GetDateFrom_yyyymmdd = DateSerial(Year:=intYear, Month:=intMonth, Day:=intDay)
            End If
        
        End If
    End If
    
End Function

Public Function GetFilePath(ByVal FileName As String) As String
Dim strResult As String
Dim astr() As String

    astr = Split(FileName, "\")
    ReDim Preserve astr(UBound(astr) - 1)
    strResult = Join(astr, "\")
    Erase astr()
    
    GetFilePath = strResult
    
End Function

Public Function GetFranName(ByVal pFranID As Long) As String
' This function should somehow be optimised for use in MySQL
' Very inefficoent to call the DB just for a field. Historically was used everywhere
Dim strSQL As String
Dim strResult As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    strSQL = "SELECT FranchiseBusinessName FROM Franchises WHERE FranchiseIDTSG  = " & pFranID
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)

'   Empty rst may be necessary b/c in all probability seeking a name for a fran we know exists
    If Not (rst.BOF And rst.EOF) Then
        strResult = rst!FranchiseBusinessName
    End If
    
    rst.Close
    Set rst = Nothing

    GetFranName = strResult
    
End Function

Public Function gfsSplitDate(ByVal sDate As String) As String
'   AUrban One day there will be no need for this function, but as things stand the program is littered with its use
    gfsSplitDate = Left(sDate, 2) & gconSpace & Mid(sDate, 3, 3) & gconSpace & Right(sDate, 2)
End Function

Sub gsubAddSubItemToListview(ByVal sMessage As String, ByVal iSubItem As Integer)
    If Not IsNull(sMessage) Then
        gvListItem.SubItems(iSubItem) = sMessage
    End If
End Sub

Sub gsubAddToLocalEventLog(ByVal pEvent As String, ByVal pFranchise As String)
Const kSpace As String = " "
Const kValListSep As String = ", " & vbNewLine
Dim lngLogMsgLen As Long
Dim lngSegment As Long
Dim dtmNow As Date
Dim strSQL As String
Dim strValList As String
Dim strLogMsg As String
Dim astrLogMsg() As String

    If g.bMaster Then   ' Exclude all but Master instance from writing to database
        dtmNow = Now
        
        strLogMsg = Trim$(pEvent)           ' Trim leading and trailing spaces
        Do While InStr(strLogMsg, Space(2)) ' Replace multiple embedded spaceswith single spaces
            strLogMsg = Replace$(Expression:=strLogMsg, Find:=Space(2), Replace:=kSpace)
        Loop
            
        lngLogMsgLen = Len(strLogMsg)
        If lngLogMsgLen > g.lngEventLogEventFldSize Then
            astrLogMsg = SplitText(pText:=strLogMsg, pMaxLength:=g.lngEventLogEventFldSize - 1)
        Else
            ReDim astrLogMsg(0 To 0)
            astrLogMsg(0) = strLogMsg
        End If
    
    '   Log each array element working backwards through array so that events broken into multiple
    '   lines that can be read easily in event log which is diplayed from most recent to oldest
        strSQL = "INSERT INTO " & vbNewLine & _
                 "EventLog(DateTime, Event, Franchise, lngDate) " & vbNewLine & _
                 "VALUES " & vbNewLine

        For lngSegment = UBound(astrLogMsg, 1) To 0 Step -1
            strValList = strValList & "(" & MySqlDateTime(dtmNow) & ", "
            If lngSegment = 0 Then
            ''' strValList = strValList & SqlQ(astrLogMsg(lngSegment)) & ", " ' 1st Segment -> No continuation char     ''' V401
                strValList = strValList & MySqlQ(astrLogMsg(lngSegment)) & ", " ' 1st Segment -> No continuation char   ''' V401
            Else
            ''' strValList = strValList & SqlQ(("_" & astrLogMsg(lngSegment))) & ", " ' Prepend continuation char ("_") ''' V401
                strValList = strValList & MySqlQ(("_" & astrLogMsg(lngSegment))) & ", " ' Prepend continuation char ("_") ''' V401
            End If
            strValList = strValList & SQ(Left$(pFranchise, g.lngEventLogFranFldSize)) & ", " & _
                                      CStr(Fix(CDbl(dtmNow))) & ")" & kValListSep
        Next lngSegment
        strValList = Left$(strValList, Len(strValList) - Len(kValListSep))
        strSQL = strSQL & strValList

        CnnDwExecute pCommandText:=strSQL
    End If

End Sub

Sub gsubRefreshEventLogDisplay()
'   Would event log would be better off with a grid rather than a listview?
Dim intPrevMousePointer As Integer
Dim lngDate As Long
Dim strSQL As String
Dim strErrMsg As String
Dim strFormattedDateTime As String
Dim rst As ADODB.Recordset
    
    intPrevMousePointer = SetMousePointer(vbHourglass)

''' Review '' Factors to take in to account when rewriting gsubRefreshEventLogDisplay to avoid un-necessary processing and possibly use new flags.
''' Review Should only refresh EventLog if current date is selected for Event Log Display (ie tdpEventLogDate=Today)
''' Review (otherwise there will be no changes that need to be reflected) and if first tab (tab containing eventlog display)
''' Review is the current tab, or a pForce parameter is passed so that
''' Review refreshing can be forced for special cases such as changing the date selected for EventLog dispaly etc.
''' Review  If (frmTSGDataWarehouse.tabMain.Tab = TabEnum.eDataCaptureTab) Then
''' Review  ACTUALLY REQUIRES SOME IN DEPTH THINKING AND WOULD PROBABLY BEST BE FIXED BY THOUGHTFUL CALLING OF THIS PROCEDURE

    If gbEventLogRefreshIsNotAlreadyInProgress Then
    
        gbEventLogRefreshIsNotAlreadyInProgress = False
    
    '   Interrogate database
        lngDate = Fix(CDbl(frmTSGDataWarehouse.tdpEventLogDate.Value))
        strSQL = "SELECT * FROM EventLog WHERE (lngDate = " & lngDate & ") ORDER BY Sequence DESC"
        Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            
    '.   Update display
    '
        frmTSGDataWarehouse.lvwEventLog.ListItems.Clear
        If rst.BOF And rst.EOF Then
            Set gvListItem = frmTSGDataWarehouse.lvwEventLog.ListItems.Add()
            gvListItem.Text = "No events this date"
        Else
            Do While Not rst.EOF
                strFormattedDateTime = Format$(rst!DateTime, "dd mmm  hh:nn:ss am/pm")
                Set gvListItem = frmTSGDataWarehouse.lvwEventLog.ListItems.Add(Text:=strFormattedDateTime)
            '   gvListItem.Text = rst(gconEventTableDateTimeField)
                gsubAddSubItemToListview rst!Franchise, 1
                gsubAddSubItemToListview rst!Event, 2
                rst.MoveNext
            Loop
        End If
        
        rst.Close
        Set rst = Nothing
        
        frmTSGDataWarehouse.Refresh
        
        gbEventLogRefreshIsNotAlreadyInProgress = True
    End If
        
    SetMousePointer intPrevMousePointer

End Sub

Public Function IsDateFmtOk() As Boolean
'   AUrban: Written 11Jan2006
'   Unfortunately the program uses strings to store and manipulate dates rather than date types.
'   Until this is rectified the current procedure has been written and will be used.
'   This function will progressively replace fbTheSystemDateFormatIsCorrect() which doesn't do exactly
'   as indicated by its MsgBox. Eventually both functions will be removed when they are no longer needed
Dim dtm As Date

    dtm = Date
    IsDateFmtOk = Format$(dtm) = Format$(dtm, "dd/MM/yy")

End Function

Public Sub LogBugFix(pUdtError As udtError, ByVal pCalledFromTag As String) '(pErrRef As String, ByVal pErrObjectCopy As ErrObject)
' Perhaps uncomment and use fGetErrorUdt
'Dim z As New ErrObject
'   pErrRef: Reference to reporting of bug. eg. Zip file name etc.
'LogBugFixed - Logging Previous run-time error (nnn) handled by changes to procedure Nnnn.
'(Anthony Urban- ddmmmyyyy) - may use fsErrDetail- should give a reference to the zipfile date-time-error
'   - this can be written as a comment in code, however the logging function is better than a comments
'(more concise) and the function can be easily disabled

Dim strMsg As String
    
    With pUdtError
        strMsg = "PLEASE INFORM ANTHONY URBAN - Logging BugFix -" & pCalledFromTag & _
                 " Line: " & .ErrLine & _
                 " ErrNo: " & .Number & _
                 " Error: " & .Description
    End With
    
    StatusBar strMsg

End Sub

Public Sub MoveFileOverWrite(ByVal pSource As String, ByVal pDest As String)
Dim fso As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject
    
    If fso.FileExists(pDest) Then
        fso.DeleteFile pDest, Force:=True
    End If
    fso.MoveFile Source:=pSource, Destination:=pDest

    Set fso = Nothing

End Sub

Public Function openNewFile(ByVal sFile As String) As Integer
    
    Call DeleteFile(sFile)
    
    openNewFile = FreeFile
    Open sFile For Output As #openNewFile
    
End Function

Public Function SafeDivide(ByVal pNumerator As Variant, ByVal pDenominator As Variant, Optional pAnswerForZeroDividedByZero As Variant = "") As Variant
'   Wrapper function to avert run-time error in Sales Report.
Dim vntResult As Variant

    If pDenominator <> 0 Then
        vntResult = pNumerator / pDenominator
    Else
    '   Divide by zero errro handled according to passed parameters
        Select Case pNumerator
            Case 0: vntResult = pAnswerForZeroDividedByZero
            Case Else
            '   Other errors UNHANDLED because this is a wrapper function to
            '   fix a specific bug. We don't want to mask un-anticipated bugs.
            '   Function may however evolve in to a generic function.
                vntResult = pNumerator / pDenominator
        End Select
    End If
    
    SafeDivide = vntResult
    
End Function

Private Function SplitText(ByVal pText As String, ByVal pMaxLength As Long) As String()
' LATER ADDITIONS COULD BE OPTIONAL FLAGS FOR REMOVING DOUBLE SPACES ETC.
'   Algorithm for splitting text into a string array
'   Set markers in text breaking it into segments <=  pMaxLength (put markers at <= every pMaxLength char pos)
'   Don't break words (except where words exceed pMaxLength and in that case break word into largest possible segments)
'   Split segments into an array
Const kDelimiter As String = "`"
Const kSpace As String = " "
Dim lngCurrentPos As Long
Dim lngLastPos As Long
Dim lngDelimiterPos As Long
Dim lngTextLength As Long
Dim strText As String

'   Set delimiters in string if required
    strText = pText
    lngTextLength = Len(strText)
    
    If lngTextLength > pMaxLength Then
        Do
            lngCurrentPos = InStr(lngCurrentPos + 1, strText, kSpace)
            Select Case True
            
            '   No spaces left in strText.
                Case lngCurrentPos = 0
                    If lngTextLength - lngLastPos > pMaxLength Then
                    '   Current segment is too long -> split it
                        lngDelimiterPos = lngDelimiterPos + pMaxLength + 1
                        lngCurrentPos = lngDelimiterPos
                        strText = Left$(strText, lngDelimiterPos - 1) & kDelimiter & Right$(strText, lngTextLength - lngDelimiterPos + 1)
                    ElseIf lngTextLength - lngDelimiterPos > pMaxLength Then
                    ''' THIS CLAUSE ADDED TO FIX A BUG. SHOULD BE REVIEWED AT SOME STAGE
                    '   SHOULD BE DOING AS PER CASE BELOW?
                    '   Remaining segment from lngDelimiterPos is too long -> split
                        If lngLastPos <> lngDelimiterPos Then
                        '   lngLastPos is a suiteable breakpoint -> replace space at lngLastPos with delimiter
                            lngDelimiterPos = lngLastPos
                            lngCurrentPos = lngDelimiterPos
                            strText = Left$(strText, lngDelimiterPos - 1) & kDelimiter & Right$(strText, lngTextLength - lngDelimiterPos)
                       Else
'*' NEED TO WORK OUT WHETHER AN ELSE CLAUSE IS NEEDED.
'   FORCE REPPORTING OF SORTS WITHOUT CRASHING IF ELSE CAUSE IS EVER EXCERCISED
                            strText = "1. PLEASE REPORT ERROR IN SplitText() Code Date 08July2009"
                            lngTextLength = Len(strText)
                            lngDelimiterPos = 0
                            lngCurrentPos = lngDelimiterPos
                        End If
                    End If
                    
            '   Found a space, but exceeded segment length -> insert a delimiter
                Case (lngCurrentPos - 1 - lngDelimiterPos) > pMaxLength
                    If lngLastPos <> lngDelimiterPos Then
                    '   lngLastPos is a suiteable breakpoint -> replace space at lngLastPos with delimiter
                        lngDelimiterPos = lngLastPos
                        lngCurrentPos = lngDelimiterPos
                        strText = Left$(strText, lngDelimiterPos - 1) & kDelimiter & Right$(strText, lngTextLength - lngDelimiterPos)
                    Else
'*' NEED TO WORK OUT WHETHER ELSE CLAUSE IS NEEDED.
'*' DOESN'T APPEAR TO BE NEEDED AS ABOVE AND PROBABLY NOT NEEDED HERE EITHER
'   FORCE REPPORTING OF SORTS WITHOUT CRASHING IF ELSE CAUSE IS EVER EXCERCISED
                        strText = "2. PLEASE REPORT ERROR IN SplitText() Code Date 08July2009"
                        lngTextLength = Len(strText)
                        lngDelimiterPos = 0
                        lngCurrentPos = lngDelimiterPos
                    End If
                    
            '   Found a space: Have EQUALLED segment length
            '   We have found a suiteable breakpoint at exactly the segment length => insert delimeter BY REPLACING SPACE AT THIS POS
                Case ((lngCurrentPos - lngDelimiterPos) = (pMaxLength))
                    lngDelimiterPos = lngCurrentPos
                    strText = Left$(strText, lngDelimiterPos - 1) & kDelimiter & Right$(strText, lngTextLength - lngDelimiterPos)
            
            End Select
            lngLastPos = lngCurrentPos
            
        Loop Until (lngCurrentPos = 0) And ((lngTextLength - lngDelimiterPos) <= pMaxLength) ' (End of spaces) AND (last segment <= pMaxLength)
    End If

    SplitText = Split(Expression:=strText, Delimiter:=kDelimiter)
    
End Function

Sub StatusBar(ByVal pMsg As String, _
     Optional ByVal pFranchise As String = vbNullString, _
     Optional ByVal pLog As Boolean = True, _
     Optional ByVal pRefreshEventLogDisplay As Boolean = True)
''' Review  '''AUrban Look at changing default value of pLog to False. TRUE will be the exception where
''' Review programmer deliberately chooses to log somethnig exceptional (eg error) or that needs to be logged.
''' Review Will be moving more to logging the unusaul/exceptional rather than using it as a development log.

'? pRefreshLog will ultimately be mandatory and will replace gsubRefreshEventLogDisplay ?'
'? In some cases, such as O/N processing, we may want to add to the log but not refresh the event
'? log display. If the program were better written we would set a flag to refresh the event log
'? display whenever we added to the log but were not on that tab. That way when we were next on
'? that tab (most likely from user interaction) we could then refresh the display.
'? Whole area needs some analysis. Is really to do with how to handle and if to rewrite for gbEventLogRefreshIsEnabled
'? As at Version 3.0.91 pRefreshLog makes no difference and could be removed because the only time it is
'? explicitly called it is always in agreement with pLog parameter. The logic for when to refresh the
'? logged event can either reside in gsubRefreshEventLogDisplay or in the call to it.

Dim strMsg As String
    
    If Len(pFranchise) = 0 Then
        strMsg = pMsg
    Else
        strMsg = pFranchise & " - " & pMsg
    End If
    
    frmTSGDataWarehouse.stb.SimpleText = strMsg
    
    If pLog Then
        gsubAddToLocalEventLog pMsg, pFranchise
        If gbEventLogRefreshIsEnabled And pRefreshEventLogDisplay Then
            gsubRefreshEventLogDisplay
        End If
    End If
        
    DoEvents    ' Assist screen refreshing

End Sub
 
''''  V401 Renamed subAddToRemoteEventLog() to AddToRemoteEventLog(), made it a private procedure
''''       and move it to the main form
'''Public Sub subAddToRemoteEventLog(ByVal pEvent As String, ByVal pFranName As String, ByRef pCnnRemote As ADODB.Connection)
'''Dim strErrMsg As String
'''Dim rstEventLog As ADODB.Recordset
'''
'''On Error GoTo Procedure_Error
'''
'''    Set rstEventLog = GetRst(pCnn:=pCnnRemote, _
'''                             pSource:="EventLog", _
'''                             pSourceType:=adCmdTable, _
'''                             pRstType:=eEditableFwdOnly, _
'''                             pErrMsg:=strErrMsg)
'''
'''    rstEventLog.AddNew
'''        rstEventLog!DateTime = Now
'''        rstEventLog!Event = Left$(Trim$(pEvent), rstEventLog!Event.DefinedSize)
'''        rstEventLog.Update
'''    rstEventLog.Close
'''
'''    Set rstEventLog = Nothing
'''
'''Procedure_Exit:
'''    Exit Sub
'''
'''Procedure_Error:
'''''' gsubAddToLocalEventLog "Remote EventLog Error " & Err.Number & ": " & Err.Description & " EventToLog: " & pEvent, _ ''' V397
''''''                        pFranName                                                                                    ''' V397
'''    StatusBar "Remote EventLog Error " & Err.Number & ": " & Err.Description & " EventToLog: " & pEvent, _
'''              pFranName, _
'''              pRefreshEventLogDisplay:=False                                                                             ''' V397
'''    Resume Procedure_Exit
'''
'''End Sub

Public Sub subPurgeTable(ByVal pTableName As String)
Const kProcName As String = "subPurgeTable"

'   g.cnnDW.BeginTrans
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx

'   An ExecuteComplete event will be issued when Execute operation concludes.
''' V389 Start
''' g.cnnDW.Execute CommandText:="DELETE FROM " & pTableName, _
                    Options:=CommandTypeEnum.adCmdText + ExecuteOptionEnum.adExecuteNoRecords
    CnnDwExecute pCommandText:="DELETE FROM " & pTableName
''' V389 End
'   g.cnnDW.CommitTrans
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx
   
End Sub

Public Sub ZipFiles(ByRef pColFullnamesToZip As VBA.Collection, ByVal pZipFile As String)
'   For information on command line options run pkzipc.exe
'   This procedure calls pkzip asynchronsouly since zipped files are emailed manually and not subsequently used by the program
''' Version 360 Noticed that Zipped files Noticed are now uploaded automatically by the programm to AZTEC
''' Version 360 To be safe this procedure should probably be converted to making a synchronous call to
''' Version 360 ensure that zipping of files is always complete before upload is attempted.
''' Version 360 Since introduction of automatic uploading of files to AZTEC we must just have been lucky enough for
''' Version 360 the zipping process to be quick enough

'   Minimal validation of passed parameters

'PKZIP(R)  Version 4.00  FAST! Compression Utility for Windows
'Usage: PKZIPC [command] [options] zipfile [@list] [files...]
'   View .zip file contents: PKZIPC zipfile
'   Create a .zip file:      PKZIPC -add zipfile file(s)...
'   Extract files from .zip: PKZIPC -extract zipfile

'The above usages are only basic examples of PKZIP's capability.
'   PKZIP Commands:
'
'Add             Default         Header           Listfile       Test
'Comment         Delete          Help             ListSfxTypes   Version
'Configuration   Extract         License          Print          View
'Console         Fix             ListCertificates Sfx
'
'   PKZIP Options:
'
'204             Directories     Lowercase        OptionChar     Smaller
'After           Encode          Mask             Overwrite      Sort
'Ascii           Exclude         Maximum          Password       Span
'Attributes      Fast            More             Path           Speed
'Authenticity    Hash            Move             Preview        Store
'Before          Header          NameSfx          Recurse        Temp
'Binary          Include         Newer            Runafter       Times
'Certificate     Larger          NoExtended       Sfx            Translate
'Comment         Level           NoFix            Shortname      Volume
'DCLImplode      ListChar        Normal           Sign           Warning
'Decode          Locale          NoZipExtension   Silent         Zipdate
'Deflate64       Logfile         Older

Dim strFilesToZip As String
Dim strCommand As String
Dim vntFileToZip As Variant
Dim fso As Scripting.FileSystemObject

    If Not pColFullnamesToZip Is Nothing Then
    '   Use custom function GetFilePath() beuause GetFolder method of FileSystemObject
    '   requires that the file exists if a file is provided as the argument
        Set fso = New Scripting.FileSystemObject
        If fso.FolderExists(GetFilePath(pZipFile)) Then
            For Each vntFileToZip In pColFullnamesToZip
                If FileExists(Trim$(vntFileToZip)) Then
                    strFilesToZip = strFilesToZip & " " & DoubleQuote(vntFileToZip)
                End If
            Next vntFileToZip
        End If
        Set fso = Nothing
    End If
    
    If Len(strFilesToZip) <> 0 Then
        strCommand = g.strPkZipCExe & " -add " & DoubleQuote(pZipFile) & " " & strFilesToZip
        ExecCmd cmdline:=strCommand, pTimeoutInMinutes:=40
    End If
    
End Sub

''' ''' V397 Start
'''Public Function IsVpnAvailable() As Boolean
'''Dim bVPN As Boolean
'''Dim lngAttempts As Long
'''Dim lngMaxAttempts As Long
'''Dim strSQL As String
'''Dim strErrMsg As String
'''Dim rstVpnFranchises As ADODB.Recordset
'''
'''
'''    If IsIDE() Then
'''        lngMaxAttempts = 1
'''    Else
'''        lngMaxAttempts = 3
'''    End If
'''
'''    strSQL = "SELECT * FROM Franchises " & vbNewLine & _
'''             "WHERE FranchiseIncludedInStatistics AND Live AND (VpnIpAddress <> '')"
'''
'''    Set rstVpnFranchises = GetRst(pCnn:=g.cnnDW, _
'''                                  pSource:=strSQL, _
'''                                  pSourceType:=adCmdText, _
'''                                  pErrMsg:=strErrMsg)
'''
''''   Give Ping 3 chances to succeed increasing timeout on each iteration
''''   We don't want to incorrectly diagnose the VPN as not available
'''    With rstVpnFranchises
'''        Do While (Not .EOF) And (lngAttempts < lngMaxAttempts) And (bVPN = False)
'''            lngAttempts = lngAttempts + 1
'''                If IsIPAddress(.Fields!VpnIpAddress) Then
'''                    bVPN = IsPingSuccessful(pIpAddress:=.Fields!VpnIpAddress.Value, pTimeout:=lngAttempts * 1000)
'''                End If
'''            .MoveNext
'''        Loop
'''        .Close
'''    End With
'''
'''    Set rstVpnFranchises = Nothing
'''    IsVpnAvailable = bVPN
'''
'''End Function
''' ''' V397 End

''' V397 Start
'''Public Sub AddRecordToEventLog(ByRef pCnnRStats As ADODB.Connection, strEvent As String)
'''Dim lngRetryCounter As Long
'''Dim rstEventLog As ADODB.Recordset
'''Dim strErrMsg As String
'''
'''On Error GoTo Retry
'''
'''    Set rstEventLog = GetRst(pCnn:=pCnnRStats, _
'''                             pSource:="EventLog", _
'''                             pSourceType:=adCmdTable, _
'''                             pRstType:=eEditableFwdOnly, _
'''                             pErrMsg:=strErrMsg)
'''    rstEventLog.AddNew
'''        rstEventLog!DateTime = Now
'''        rstEventLog!Event = Left$(strEvent, rstEventLog!Event.DefinedSize)
'''    rstEventLog.Update
'''    rstEventLog.Close
'''    Exit Sub
'''
'''Retry:
'''    lngRetryCounter = lngRetryCounter + 1
'''    Call Sleep(1000)
'''    If lngRetryCounter < 10 Then
'''        Resume
'''    End If
'''    MsgBox "Event log entry was not added, table is locked", vbExclamation
'''
'''End Sub
''' V397 End


