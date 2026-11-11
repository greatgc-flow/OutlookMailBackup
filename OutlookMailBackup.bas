Attribute VB_Name = "OutlookMailBackup"
Option Explicit

' ============================================================================
' Outlook Mail Backup 3.5.0 - one standard module, Windows classic Outlook
' Built: 2026-11-11. ASCII source / CRLF for VBA File > Import.
'
' START -- ONE VISIBLE MACRO
' 1. Replace the previous OutlookMailBackup module, then Debug > Compile.
' 2. Run MB_01_RunBackup. It is the ONLY public no-argument macro in this module.
' 3. First use in each Outlook VBA session runs disposable native checks for
'    hashing, locale/UTC, file moves, MSG comparison and metadata recovery.
'    Synthetic UNSENT messages and owned temporary metadata are used. No real
'    mailbox item is saved, sent, moved or deleted by the acceptance checks.
' 4. Select a base/account directory only when no usable remembered path exists.
'    Setup, safe recovery, year/month maintenance, legacy linking and adaptive
'    backup are automatic. There is no separate sequence of public commands.
' 5. The same macro while busy requests cancellation. STOP.txt in MailBackup
'    (the selected managed root) also cancels. Wait for handle cleanup.
'
' DIRECTORIES -- NEW ROOTS USE VERSIONLESS ROLE NAMES
'   <selected base>\MailBackup\State\
'       journal.tsv / moves.tsv / account-roots.tsv / stages.tsv
'       coverage_*.tsv / source_audit_*.tsv / archive_audit_*.tsv
'       baseline_*.tsv / setup_*.tsv / reconcile_*.tsv / runs\<run>\
'   <selected account>\yyyy\mm\<message>.msg
'   <selected account>\_Recovery\
'       recovery-owner.tsv / snapshots\ / chunks\ / stage-intents.tsv
'       restore-intents.tsv / reconcile_*.tsv / recovery-events.tsv / overlays\
' State is the active operating directory; _Recovery holds independent metadata.
' Existing SafeV2/_MailBackup_Data/_MailBackup_Recovery bindings are read in place.
' Prior _state TSVs migrate by verified copy; their original history is retained.
' No user-specific drive or absolute output directory is embedded in this module.
' New MBS4 journal rows use account-relative file paths. At the I/O boundary,
' they expand against one exact account-root binding; old MBS2/MBS3 rows remain
' readable without rewrite. Roots, recovery/move proof and out-of-root bindings
' still use explicit absolute paths. Relative records do not authorize moving
' a bound archive by hand or silently changing its ownership/history.
' Root-level owner.txt, operating-format.tsv and operating-store.tsv, plus the
' account's _mb_account.tsv, retain ownership/binding outside the primary state.
' One AUTO managed root belongs to one Outlook StoreID. Another store requires
' another managed root. Foreign-store proof is rejected before archive work.
'
' PATH SETTINGS / DELETION
' A checked per-store locator also remembers base, managed and account paths:
'   %LOCALAPPDATA%\OutlookMailBackup\locations\base_<StoreID SHA256>.tsv
' It complements HKCU/account maps; it never overrides a conflicting owner.
' Previous locator bytes are preserved before a locked, checked replacement.
' Exact account/recovery markers can rediscover custom-named sibling accounts.
' With all exact references gone, select the existing base/account once. An
' unrelated hardcoded default is never treated as a remembered selection.
' Existing active state evidence prevents stale _state migration overwrites.
'
' OPERATING-STATE RECOVERY
' After writers close, checked primary TSV files become an independent complete
' recovery generation under the selected account. Metadata uses content-addressed
' 1 MiB chunks; unchanged chunks are reused. MSG payloads are NOT copied there.
' Missing/damaged primary state restores a fully checked generation. BEFORE
' replacement, surviving primary TSVs and the independent stage mirror receive
' exact immutable copies under overlays. A checked plan binds those raw hashes
' and the original snapshot; durable START/DONE resumes that SAME plan. Newer
' journal, baseline, move/publication proof and unresolved audit REVIEW records
' are retained. A same-key expected-hash conflict stops recovery, never rebases
' changed bytes. Move histories require an ordered-prefix relation. A checked
' unknown schema or orphan semantic reference is retained for explicit review.
' A legacy interrupted restore with unbound preserved histories also stops for
' review. No automatic choice can prove which conflicting history is correct.
' Snapshot order uses a monotonic generation, not only wall-clock time.
' Missing recovery owner can be rebuilt only from a complete exactly bound
' snapshot; folder names alone never authorize ownership or destructive actions.
' If snapshots are unavailable, surviving checked proof is still retained. The
' REAL index validator checks source/store bindings and prior R references. Raw
' malformed framing remains preserved; losing a checksum is not proof of a
' trustworthy historical hash. An odd-byte or divergent move history stops.
' Unindexed readable MSGs receive MSG_ONLY_REBASELINED records. This proves their
' observed bytes NOW; it cannot reconstruct lost old hashes/IDs or prove historical
' equality. Indexed/published expected hashes always take precedence over a newly
' observed baseline. Missing files remain recorded; changed bytes are not accepted
' as replacements merely because Outlook can open them.
' Recovery sets durable NEEDS reconciliation tokens in both metadata locations.
' Cancellation/restart keeps this requirement. Both become DONE only after the
' complete current-mail inventory succeeds and its coverage is durably recorded.
' Deleting both metadata locations cannot reconstruct deleted MSGs. An additional
' independent backup of the whole archive is still needed for disk/media loss.
'
' DATA PUBLICATION / RESTART
' Unicode MSG is written to an exclusively owned same-directory stage. Reopen
' comparison, SHA-256 and size precede publication. RESERVE and PREPARED intents
' are flushed to two metadata locations before their dependent action.
' Stage bytes are flushed; MoveFileExW uses WRITE_THROUGH without COPY_ALLOWED or
' REPLACE_EXISTING. Final bytes are flushed and checked before index commitment.
' A read guard denies new write/delete handles during final verification/linking.
' A published verified old revision is indexed even if the live source changes
' immediately afterward; the CURRENT source still becomes a retry outcome.
' PREPARED orphan output can be recovered only against its prior hash/size.
' Readable RESERVE-only MSG remains a candidate requiring content comparison.
' Incomplete owned stages are preserved with a .held suffix, never silently
' deleted or passed as successful mail. Unproven unrelated fragments need REVIEW.
' Rename recovery/flushes do not claim immunity to hardware or total power loss.
'
' YEAR / MONTH AND FILENAMES
' Existing and future messages use the SAME selected account\yyyy\mm folder.
' Valid SentOn, then ReceivedTime, then CreationTime chooses the filing month.
' Recognized original YYYYMMDD_HHNNSS names preserve their original filing date.
' No filesystem modification time chooses the filing month. Partial year/month
' organization is recursively inventoried and refreshed on subsequent runs.
' A checked complete plan precedes moves. Expected bytes are verified; both
' colliding files survive with a suffix. Empty old folders remain. Moves have
' durable PREPARE/COMMIT records; redirects bind both old path AND expected hash.
' Exact unique same-account bytes can relink a missing reference after a manual
' relocation. An existing changed file is never hidden by this missing-path rule.
' Existing names are retained, apart from necessary collision suffixes.
' New names: yyyyMMdd_HHmmss_sender_subject.msg. Sender uses the historic first
' 10 UTF-16 units, before "("; neither sender nor subject splits a surrogate pair.
' Subjects use remaining space within the conservative 235 UTF-16 full-path
' budget, reserving collision suffix room. Illegal Windows characters are cleaned.
' No fixed 40-character subject cap or always-present 64-hex suffix is added.
' Very long subjects can still be truncated; full Subject remains inside MSG.
' Everything filename search cannot recover characters absent from the filename.
' Separate attachment extraction defaults OFF; embedded attachments remain in MSG.
'
' LEGACY / RESTORED MAIL COMPATIBILITY
' Existing verified MSGs are reused without MailItem.SaveAs. MBS2/MBS3/MBS4 index
' records remain readable. Original file names and sent-time keys are CANDIDATE
' hints only; matched sender/recipients/headers/body/HTML/RTF and supported ordinary
' attachment bytes/inline metadata must pass the declared comparison policy.
' A filename-tail fallback tolerates timezone-shifted old name prefixes. Locale
' separator normalization never guesses an ambiguous calendar or date order.
' L and R final selections are guarded through their committed reference. R aliases
' require prior same-store expected bytes plus full bounded content comparison;
' Message-ID alone never proves equality. Changed EntryID can therefore reuse
' existing files, without pretending provider-rewritten dates/content are equal.
' Missing/corrupt recorded outputs are not automatically re-exported. Unsupported
' special attachments and uncertain old matches stay explicit REVIEW outcomes.
' The conservative old-archive write-time guard can defer unmatched historical
' messages. This guard is NOT proof that all old mail was previously backed up.
' v3.4 adds exact Importance, Sensitivity and Unicode category-name membership
' checks; it reads the multivalued Keywords property without locale separators.
' Any UserProperties produce limited validation for a newly saved MSG, and
' prevent an equal-content legacy/R alias. Flags, tasks, reminders, read status,
' arbitrary MAPI properties and every special attachment are not certified.
' V1/V2 legacy proof and V2 new/R proof remain readable. Old proof labels are
' rechecked before V3 promotion; a changed policy forces full comparison on the
' first upgraded successful coverage run. Existing MSGs are not bulk exported.
' A proven same-source content change may create a separate new revision, even
' when its mail date is older than the legacy cutoff. Equal content reuses an
' existing file; ambiguous, missing or corrupt prior proof remains REVIEW.
' This is forward reuse of old outputs. Old writers do not understand new state:
' owner.txt is upgraded to Schema4 before relative rows; older writers reject it.
' Do not alternate
' macro versions; an arbitrary original script cannot be retroactively controlled.
'
' ADAPTIVE COVERAGE / INDEPENDENT AUDITS
' AUTO reads lightweight metadata for all ordinary MailItems in the DEFAULT
' mounted Outlook store each run. Other PST/shared/online stores are not implicit.
' The last successful scan START is UTC. The next selective range begins 7 days
' earlier; there is no 30/90-day absence cap. Initially use 30 days plus EVERY
' unknown identity or changed modification stamp, at any historical mail date.
' A new/imported old message is therefore eligible beyond the nominal range.
' Backward clocks, changed timezone rules/policy and metadata recovery force full
' current-mail reconciliation. New dates are ASCII Gregorian; VBA Calendar is
' restored. AUTO uses typed Table dates rather than locale-formatted Jet filters.
' Every run also rotates up to 500 historical source comparisons. Unresolved
' source reviews have durable retry obligations outside the normal range/budget.
' Source audit cursor advances only after visited results/retry obligations are
' written. Unvisited/cancelled work does not falsely pass. Old review aggregates
' are refreshed; a persistent issue does not freeze the rest of the audit.
' Separately, up to 500 physical archived proofs are checked per run, including
' originals no longer in Outlook. Up to 100 pending issues retry plus rotation.
' Archive checks retain expected SHA/size, reopen MSG, and keep missing-file
' tombstones/conflicting proofs. Source REVIEW does not stop physical rotation.
' These budgets are not a promise that every possible issue is found instantly.
' Current mail coverage, physical archive integrity, baseline provenance and
' reporting are separate summary results. A broad PARTIAL label does not itself
' mean that every current-mail coverage checkpoint was rejected.
'
' VALIDATION / PERFORMANCE LIMITS
' Authoring checks are syntax/static linkage and adversarial state models.
' Actual Windows Outlook was not available during authoring. Debug > Compile is
' required on the target PC; native disposable checks then run automatically.
' Acceptance logs explicitly retain NOT_RUN for real archive compatibility,
' complete VBA project compilation, all special MAPI cases and 100k load tests.
' Actual AUTO item times are measured without extra SaveAs. Separate NEW/REUSE
' means and containing p50/p95 intervals include all timed observations. At
' least 30 ordinary timed successes are required for an indicative 100k item
' processing-cost forecast; recovery/inventory/moves/audits remain additional.
' No observations means NOT_MEASURED. This is not a measured 100k load test.
' Reuse avoids repeated SaveAs, not all I/O. Full metadata inventories, bounded
' content audits and metadata snapshots still take time. First MSG-only recovery
' hashes/reopens every unindexed MSG; it can be much slower than later increments.
' All providers/attachments are not transactional. A provider can change content
' without a useful revision stamp; rotating audits reduce, not eliminate, this risk.
' Reparse/device paths and >235-unit paths are refused; account roots <=80 units.
' Windows classic Outlook only; new Outlook, Mac and server automation unsupported.
' Reports: MailBackup\State\runs\<run>\summary.txt and checked audit/recovery logs.
' Helper/expert procedures are PRIVATE; none needs to be run by the user.
' ============================================================================

Private Const MB_VERSION As String = "3.5.0"
Private Const MB_OWNER As String = "OutlookMailBackup.Schema4"
Private Const MB_OWNER_V34 As String = "OutlookMailBackup.SafeV3.4"
Private Const MB_EXTRACT_ATTACHMENTS As Boolean = False
Private Const MB_ATTACHMENT_EXTENSIONS As String = "*"
' Alternative: "pdf;doc;docx;xls;xlsx;ppt;pptx;zip;txt;csv"
Private Const MB_REOPEN_NEW_MSG As Boolean = True
Private Const MB_NEW_CONTENT_VALIDATION As String = "REOPENED_CONTENT_SHA256_V3"
Private Const MB_NEW_CONTENT_VALIDATION_V2 As String = "REOPENED_CONTENT_SHA256_V2"
Private Const MB_NEW_BASIC_VALIDATION As String = "REOPENED_METADATA_SHA256_V3"
Private Const MB_NEW_BASIC_VALIDATION_V2 As String = "REOPENED_METADATA_SHA256_V2"
Private Const MB_VERIFY_EXISTING_HASH As Boolean = True
Private Const MB_JOURNAL_FLUSH_EVERY As Long = 25
Private Const MB_BENCHMARK_COUNT As Long = 300
Private Const MB_ERR As Long = vbObjectError + 7600
Private Const MB_CHANGE_KEY As String = "http://schemas.microsoft.com/mapi/proptag/0x65E20102"

Private gBusy As Boolean
Private gCancel As Boolean
Private gFatal As Boolean
Private gFatalText As String
Private gRoot As String
Private gRunId As String
Private gRunDir As String
Private gRunStart As Date
Private gSince As Date
Private gUntil As Date
Private gClockStart As Double
Private gLastPulse As Double
Private gLastProgress As Double
Private gPhase As String
Private gMode As String
Private gFSO As Object
Private gIndex As Object
Private gDirs As Object
Private gEvents As Object
Private gJournal As Object
Private gJournalPath As String
Private gJournalWrites As Long
Private gJournalIssues As Long
Private gFoldersScanned As Long
Private gFolderFailures As Long
Private gPolicySkipped As Long
Private gCandidates As Long
Private gScanChanged As Long
Private gPlanned As Long
Private gFullPlanned As Long
Private gProcessed As Long
Private gMailNew As Long
Private gMailReused As Long
Private gMailFailed As Long
Private gMailChanged As Long
Private gMailPolicy As Long
Private gContentReview As Long
Private gMailPending As Long
Private gAttFound As Long
Private gAttNew As Long
Private gAttReused As Long
Private gAttExcluded As Long
Private gAttFailed As Long
Private gAttPending As Long
Private gBytesNew As Double
Private gScanSeconds As Double
Private gIndexSeconds As Double
Private gExportSeconds As Double
Private gSampleSeconds() As Double
Private gSampleCount As Long
Private gScope As String
Private gCleanupIssue As String
Private gModifiedWindow As Boolean

' Currency is a signed 64-bit storage type scaled by 10,000. Both QPC values
' use that same storage/scaling, which cancels in the division.
#If VBA7 Then
Private Declare PtrSafe Function MB_QueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare PtrSafe Function MB_QueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
#Else
Private Declare Function MB_QueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare Function MB_QueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
#End If


' MailBackup helper declarations. Insert before ALL procedures in the final module.
' Windows classic Outlook; VBA7 32/64-bit. CNG-managed hash objects require Windows 7+.
' No added VBA references: Scripting.FileSystemObject / Dictionary are late bound.
Private Const MBH_MAX_PATH_UNITS As Long = 235
Private Const MBH_ERR_BASE As Long = vbObjectError + 7100
Private Const MBH_GENERIC_READ As Long = &H80000000
Private Const MBH_GENERIC_WRITE As Long = &H40000000
Private Const MBH_FILE_SHARE_READ As Long = &H1
Private Const MBH_CREATE_NEW As Long = 1
Private Const MBH_OPEN_EXISTING As Long = 3
Private Const MBH_OPEN_ALWAYS As Long = 4
Private Const MBH_FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const MBH_FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const MBH_ERROR_FILE_EXISTS As Long = 80
Private Const MBH_ERROR_ALREADY_EXISTS As Long = 183
Private MBH_Stages As Object
Private MBH_FileSystemCache As Object
Private MBH_NormalizedPaths As Object
Private MBH_PathCacheHits As Double
Private MBH_PathCacheMisses As Double
Private MBH_OpenHashes As Long

#If VBA7 Then
Private MBH_Algorithm As LongPtr
Private MBH_RunLock As LongPtr
Private Type MBH_HashContext
    Handle As LongPtr
End Type
Private Declare PtrSafe Function MBH_BCryptOpenAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptOpenAlgorithmProvider" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function MBH_BCryptCreateHash Lib "bcrypt.dll" Alias "BCryptCreateHash" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByVal pbHashObject As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function MBH_BCryptHashData Lib "bcrypt.dll" Alias "BCryptHashData" (ByVal hHash As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function MBH_BCryptFinishHash Lib "bcrypt.dll" Alias "BCryptFinishHash" (ByVal hHash As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function MBH_BCryptDestroyHash Lib "bcrypt.dll" Alias "BCryptDestroyHash" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function MBH_BCryptCloseAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptCloseAlgorithmProvider" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function MBH_CreateFileW Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function MBH_ReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function MBH_CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function MBH_MoveFileW Lib "kernel32" Alias "MoveFileW" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr) As Long
Private Declare PtrSafe Function MBH_DeleteFileW Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As LongPtr) As Long
#Else
Private MBH_Algorithm As Long
Private MBH_RunLock As Long
Private Type MBH_HashContext
    Handle As Long
End Type
Private Declare Function MBH_BCryptOpenAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptOpenAlgorithmProvider" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function MBH_BCryptCreateHash Lib "bcrypt.dll" Alias "BCryptCreateHash" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function MBH_BCryptHashData Lib "bcrypt.dll" Alias "BCryptHashData" (ByVal hHash As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function MBH_BCryptFinishHash Lib "bcrypt.dll" Alias "BCryptFinishHash" (ByVal hHash As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function MBH_BCryptDestroyHash Lib "bcrypt.dll" Alias "BCryptDestroyHash" (ByVal hHash As Long) As Long
Private Declare Function MBH_BCryptCloseAlgorithmProvider Lib "bcrypt.dll" Alias "BCryptCloseAlgorithmProvider" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function MBH_CreateFileW Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function MBH_ReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function MBH_CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function MBH_MoveFileW Lib "kernel32" Alias "MoveFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function MBH_DeleteFileW Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
#End If


' Path settings persist per Windows user, independent of the imported module file.
Private Const MBP_SETTINGS_APP As String = "OutlookMailBackup"
Private Const MBP_SETTINGS_SECTION As String = "PathsV2"
Private Const MBP_BASE_KEY As String = "BackupBase"
Private Const MBP_MANAGED_LEAF As String = "MailBackup"
Private Const MBP_MANAGED_LEAF_LEGACY As String = "SafeV2"
Private Const MBP_LOCATOR_SCHEMA As String = "MBBASELOCATION1"
Private mbpLocatorPath As String
Private mbpLocatorHash As String
Private mbpLocatorObserved As Boolean
Private mbpLocatorStore As String
Private mbpLocatorManaged As String
Private mbpLocatorAccount As String



' Legacy catalog uses filename keys, with read-only SentOn fallback for renamed MSGs.
Private mblByName As Object
Private mblByStamp As Object
Private mblVisited As Object
Private gLegacyFileCount As Long
Private gLegacyFolderCount As Long
Private gLegacyUnrecognizedNames As Long
Private gLegacyNewestFileTime As Date
Private gLegacyScanComplete As Boolean
Private Const MBL_REPARSE_POINT As Long = 1024
Private Const MBL_ID_UNICODE As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const MBL_ID_ANSI As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Private Const MBL_MAPI_NOT_FOUND As Long = &H8004010F
Private Const MBL_CID_UNICODE As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Private Const MBL_CID_ANSI As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
Private Const MBL_LOCATION_UNICODE As String = "http://schemas.microsoft.com/mapi/proptag/0x3713001F"
Private Const MBL_LOCATION_ANSI As String = "http://schemas.microsoft.com/mapi/proptag/0x3713001E"
Private Const MBL_MIME_UNICODE As String = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
Private Const MBL_MIME_ANSI As String = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
Private Const MBL_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


' Existing-export compatibility is conservative: an uncertain old export is
' deferred for review, never silently labelled complete or automatically replaced.
Private Const MB_LEGACY_VALIDATION As String = "LEGACY_COMPARE_SHA256_V3"
Private Const MB_LEGACY_VALIDATION_V2 As String = "LEGACY_COMPARE_SHA256_V2"
Private Const MB_LEGACY_VALIDATION_V1 As String = "LEGACY_COMPARE_SHA256_V1"
Private gBaseRoot As String
Private gSelectedBase As String
Private gLegacyRoot As String
Private gLegacyBindings As Object
Private gLegacySources As Object
Private gLegacyOwners As Object
Private gMailLegacy As Long
Private gMailReview As Long
Private gLegacyUnmatched As Long
Private gLegacySeconds As Double



' Account roots are bound to StoreID and this managed state root, not display names.
Private Const MBY_ACCOUNT_MARKER As String = "_mb_account.tsv"
Private Const MBY_MAP_FILE As String = "account-roots.tsv"
Private Const MBY_ATTACHMENT_EXTENSIONS As String = ";pdf;doc;docx;xls;xlsx;ppt;pptx;zip;txt;"
Private mbyRoots As Object
Private mbyValidated As Object
Private mbyLoadedState As String



' Year/month move engine: durable intent records, no replace, no copy/delete.
Private mbmOps As Object
Private mbmRedirects As Object
Private mbmAccountLocks As Object
Private mbmLoaded As Boolean
Private gMovePlanned As Long
Private gMoveKept As Long
Private gMoveSkipped As Long
Private gMoveApplied As Long
Private gMoveAlready As Long
Private gMoveRestored As Long
Private gMoveRecovered As Long
Private gMoveReview As Long
Private gMovePlanReview As Long
Private gMovePlanPath As String
#If VBA7 Then
Private Declare PtrSafe Function MBM_MoveFileExW Lib "kernel32" Alias "MoveFileExW" (ByVal source As LongPtr, ByVal target As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function MBM_FlushFileBuffers Lib "kernel32" Alias "FlushFileBuffers" (ByVal handle As LongPtr) As Long
#Else
Private Declare Function MBM_MoveFileExW Lib "kernel32" Alias "MoveFileExW" (ByVal source As Long, ByVal target As Long, ByVal flags As Long) As Long
Private Declare Function MBM_FlushFileBuffers Lib "kernel32" Alias "FlushFileBuffers" (ByVal handle As Long) As Long
#End If



' MBS4 stores file references relative to an exact account-root binding.
' MBS3 absolute records are still readable; normalization is only in memory.
' M/A keys remain the lookup namespace so MBS2 backups are still reusable.
Private gExternalRoots As Object
Private gRecordMeta As Object
Private gStoreTokens As Object
Private gManagedFilePaths As Object
Private gLegacyAccountStats As Object
Private gArchiveRoots As Collection
Private gFilePlan As String
Private gFileAccount As String
Private gFileDetail As String
Private gReuseBlocked As Boolean
Private gReuseReason As String

Private gFileExtraReview As Long


' One visible command. Persistent setup states are bound to account + archive.
Private gAutoNeedsRegistration As Boolean
Private gAutoStoreId As String
Private gAutoAccount As String
Private gAutoKey As String
Private gAutoSetupPhase As String
Private gAutoSetupNote As String
Private gAutoRegistrationSummary As String
Private gAutoReviewCount As Long
Private gAutoFailureCount As Long
Private gAutoFileReview As Long
Private gAutoLegacyReview As Long
Private gAutoSetupSeconds As Double
Private gAutoTestReport As String



' Acceptance reports describe their tested scope; unexecuted operational gates
' remain NOT_RUN. No global pass flag is stored here or in the Windows registry.
Private Const MBT_REPORT_SCHEMA As String = "MBTEST1"



' Locale/time helpers. Keep before every procedure in the assembled module.
' All state/control stamps are ASCII Gregorian; UTC stamps explicitly end in Z.
Private Const MBC_ERR As Long = vbObjectError + 7600
Private Const MBC_MODIFIED_UTC As String = "http://schemas.microsoft.com/mapi/proptag/0x30080040"
Private Const MBC_RECEIVED_UTC As String = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
Private Const MBC_SENT_UTC As String = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
Private Const MBC_CREATED_UTC As String = "http://schemas.microsoft.com/mapi/proptag/0x30070040"

Private Type MBC_SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type MBC_DYNAMIC_TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As MBC_SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As MBC_SYSTEMTIME
    DaylightBias As Long
    TimeZoneKeyName(0 To 127) As Integer
    DynamicDaylightTimeDisabled As Byte
    ReservedPadding(0 To 2) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Sub MBC_GetSystemTime Lib "kernel32" Alias "GetSystemTime" (ByRef value As MBC_SYSTEMTIME)
Private Declare PtrSafe Function MBC_GetDynamicTimeZoneInformation Lib "kernel32" Alias "GetDynamicTimeZoneInformation" (ByRef value As MBC_DYNAMIC_TIME_ZONE_INFORMATION) As Long
#Else
Private Declare Sub MBC_GetSystemTime Lib "kernel32" Alias "GetSystemTime" (ByRef value As MBC_SYSTEMTIME)
Private Declare Function MBC_GetDynamicTimeZoneInformation Lib "kernel32" Alias "GetDynamicTimeZoneInformation" (ByRef value As MBC_DYNAMIC_TIME_ZONE_INFORMATION) As Long
#End If


' Search-friendly names for NEW MSG files only. Existing paths remain indexed.
' Keep the original timestamp_sender_subject ordering and sender width.
Private Const MBN_SENDER_UNITS As Long = 10
Private Const MBN_COLLISION_RESERVE As Long = 8



' Restored identities are aliases of bounded, content-verified existing bytes.
' R uses the current root-bound journal schema and the ordinary M|revision cache.
' Candidate subjects are exact Unicode strings; hashes select, never prove.
Private Const MBI_MAX_CANDIDATES As Long = 128
Private gIdentityBuckets As Object
Private gIdentitySeen As Object
Private gIdentityHeaders As Object
Private gIdentityLookupState As Object
Private gIdentityAliases As Long
Private gIdentityHeaderReads As Long
Private gIdentityHeaderHits As Long


' Adaptive coverage: full lightweight inventory, selective content work.
Private Const MBW_POLICY As String = "AUTO_INVENTORY_UTC_V2_FIDELITY_V3"
Private Const MBW_OVERLAP_DAYS As Long = 7
Private Const MBW_INITIAL_DAYS As Long = 30
Private Const MBW_AUDIT_LIMIT As Long = 500
Private gWindowStore As String
Private gWindowKnown As Object
Private gWindowAudit As Object
Private gWindowStartUTC As Date
Private gWindowPreviousUTC As Date
Private gWindowContext As String
Private gWindowPath As String
Private gWindowCursor As String
Private gWindowNextCursor As String
Private gWindowFull As Boolean
Private gWindowReady As Boolean
Private gWindowInventoryComplete As Boolean
Private gWindowCommitted As Boolean
Private gWindowNote As String
Private gWindowNew As Long
Private gWindowChanged As Long
Private gWindowPolicyQueued As Long
Private gWindowRecent As Long
Private gWindowAuditQueued As Long
Private gWindowFullQueued As Long
Private gWindowHistoricalSkipped As Long
Private gWindowBadRows As Long

Private gWindowAuditOrder As Collection
Private gWindowAuditSeen As Object
Private gWindowAuditResults As Object
Private gWindowSourcePending As Object
Private gWindowAuditPath As String
Private gWindowAuditCommitted As Boolean
Private gWindowAuditStateIssues As Long
Private gWindowAuditResultsWritten As Long
Private gWindowRetryQueued As Long
Private gWindowAuditPass As Long
Private gWindowAuditLoaded As Boolean


' Versionless names for new roots; prior absolute SafeV2 bindings stay in place.
Private Const MBS_DATA_LEAF As String = "State"
Private Const MBS_DATA_LEAF_LEGACY As String = "_MailBackup_Data"
Private Const MB_OWNER_PREVIOUS As String = "OutlookMailBackup.SafeV3.3"
Private Const MB_OWNER_LEGACY As String = "OutlookMailBackup.SafeV2"
Private gStateRoot As String
Private gStateWasMissing As Boolean
Private gStateMigrated As Boolean
Private gStageEnabled As Boolean
Private gStageRecords As Object
Private gStageExpected As Object
Private gStageMirror As String
Private gStageAccount As String
Private gStageReview As Long
Private gStageRecovered As Long
Private gStageHeld As Long
Private gStageLedgerRepaired As Long
Private gStageProofCopies As Long
Private gMaintenanceReady As Boolean
Private gAutoMaintenance As Boolean
Private gEarlyAccount As String



' Recoverable metadata: primary state is replaceable; MSG bytes are never rebuilt.
Private Const MBR_FOLDER As String = "_Recovery"
Private Const MBR_FOLDER_LEGACY As String = "_MailBackup_Recovery"
Private Const MBR_MIRROR_LEAF As String = "_independent_stage_mirror.tsv"
Private Const MBR_OWNER_FILE As String = "recovery-owner.tsv"
Private Const MBR_CHUNK_BYTES As Long = 1048576
Private gRecoveryForceFull As Boolean
Private gRecoveryRebased As Boolean
Private gRecoveryFiles As Object
Private gRecoverySentHints As Object
Private gRecoveryReviewCount As Long
Private gRecoveryRestoredCount As Long
Private gRecoveryNewBaselines As Long
Private gRecoveryNote As String
Private mbrActiveRecoveryRoot As String
Private mbrBaselineWrites As Long
Private gRecoverySnapshotsSaved As Long
Private gRecoverySalvagedRows As Long
Private gRecoveryMergedRows As Long
Private gRecoveryOverlayRejected As Long
Private gRecoveryOverlayFiles As Long
Private mbrDigestFiles As Object

#If VBA7 Then
Private Declare PtrSafe Function MBR_WriteFile Lib "kernel32" Alias "WriteFile" (ByVal handle As LongPtr, ByRef buffer As Any, ByVal count As Long, ByRef written As Long, ByVal overlapped As LongPtr) As Long
#Else
Private Declare Function MBR_WriteFile Lib "kernel32" Alias "WriteFile" (ByVal handle As Long, ByRef buffer As Any, ByVal count As Long, ByRef written As Long, ByVal overlapped As Long) As Long
#End If


' Physical file audit does not require a surviving Outlook source item.
Private Const MBA_FILE_BUDGET As Long = 500
Private Const MBA_RETRY_BUDGET As Long = 100
Private gArchiveAuditEntries As Object
Private gArchiveAuditIssues As Object
Private gArchiveAuditCheckedPaths As Object
Private gArchiveAuditVisited As Object
Private gArchiveAuditStatePath As String
Private gArchiveAuditCursor As String
Private gArchiveAuditStore As String
Private gArchiveAuditAccount As String
Private gArchiveAuditChecked As Long
Private gArchiveAuditPassed As Long
Private gArchiveAuditFailed As Long
Private gArchiveAuditPending As Long
Private gArchiveAuditFilesRead As Long
Private gArchiveAuditReferenceReview As Long
Private gArchiveAuditStateIssues As Long
Private gArchiveAuditTotal As Long
Private gArchiveAuditCommitted As Boolean


' Same-source revisions use a direct history head, independent of Subject/date.
' This authorizes a new version only after an unchanged prior MSG is compared.
Private Const MBV_NONE As Long = 0
Private Const MBV_REUSED As Long = 1
Private Const MBV_REVIEW As Long = 2
Private Const MBV_CANCELLED As Long = 3
Private Const MBV_CHANGED As Long = 4
Private gRevisionHeads As Object
Private gRevisionCompared As Long
Private gRevisionReused As Long
Private gRevisionDifferent As Long
Private gRevisionReview As Long


' Actual AUTO-run observations. Fixed memory; no extra SaveAs or mailbox reads.
' Counts are Double to avoid a 32-bit counter overflow in a long-running session.
Private Const MBF_GROUP_COUNT As Long = 12
Private Const MBF_BIN_COUNT As Long = 26
Private Const MBF_MIN_FORECAST_COUNT As Long = 30
Private Const MBF_FIRST_UPPER_SECONDS As Double = 0.001
Private gMetricObserved(1 To MBF_GROUP_COUNT) As Double
Private gMetricTimed(1 To MBF_GROUP_COUNT) As Double
Private gMetricSeconds(1 To MBF_GROUP_COUNT) As Double
Private gMetricMinimum(1 To MBF_GROUP_COUNT) As Double
Private gMetricMaximum(1 To MBF_GROUP_COUNT) As Double
Private gMetricZeros(1 To MBF_GROUP_COUNT) As Double
Private gMetricHistogram(1 To MBF_GROUP_COUNT, 1 To MBF_BIN_COUNT) As Double
Private gMetricFault As Boolean
Private gMetricFaultText As String


' Bounded stable metadata comparison. Arbitrary MAPI/custom fields are not equal
' merely because Subject/Body/attachments compare equal.
Private Const MBE_CATEGORY_KEYWORDS As String = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/Keywords"
Private Const MBE_MAX_CATEGORIES As Long = 1024


' ====================== Single-command user workflow ========================
Public Sub MB_01_RunBackup()
    Static nativeTestsPassed As Boolean
    Static nativeReport As String
    Dim detail As String, n As Long
    If gBusy Then
        gCancel = True
        Debug.Print "Cancellation requested. The current safe unit will finish first."
        Exit Sub
    End If
    On Error GoTo Failed
    If Not nativeTestsPassed Then
        If Not MBT_RunAcceptanceReport(nativeReport, detail) Then
            MsgBox "Environment check did not pass. No archive workflow was started." & _
                   vbCrLf & detail & vbCrLf & "Report: " & nativeReport, vbExclamation, "Mail Backup"
            Exit Sub
        End If
        nativeTestsPassed = True
    End If
    gAutoTestReport = nativeReport
    MB_Run "AUTO", False, True, False
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_ReleaseCrypto
    On Error GoTo 0
    MsgBox "The backup command stopped: " & CStr(n) & " - " & detail, vbExclamation, "Mail Backup"
End Sub

Private Sub MBU_Reset()
    gAutoNeedsRegistration = False
    gAutoStoreId = "": gAutoAccount = "": gAutoKey = ""
    gAutoSetupPhase = "": gAutoSetupNote = "": gAutoRegistrationSummary = ""
    gAutoReviewCount = 0: gAutoFailureCount = 0: gAutoFileReview = 0: gAutoLegacyReview = 0: gAutoSetupSeconds = 0
End Sub

Private Function MBU_ValidationPolicy() As String
    MBU_ValidationPolicy = MB_LEGACY_VALIDATION & "|" & MB_NEW_CONTENT_VALIDATION & "|MAINTENANCE_3.3"
End Function

Private Sub MBU_SaveSetup(ByVal phase As String, ByVal planPath As String, ByVal note As String)
    Dim payload As String
    If gCancel Or gFatal Then Exit Sub
    payload = MB_EncodeFields(Array("MBAUTO1", gRoot, gAutoStoreId, gAutoAccount, phase, planPath, note, _
                                   CStr(gAutoFileReview), CStr(gAutoLegacyReview), MBU_ValidationPolicy()))
    MBS_WriteAtomic MBU_SetupPath(), payload & vbTab & MBH_HashText(payload) & vbTab & "END" & vbCrLf
    MBP_SaveSettingText gAutoKey, payload & vbTab & MBH_HashText(payload) & vbTab & "END"
    gAutoSetupPhase = phase: gAutoSetupNote = note
End Sub

Private Function MBU_SetupPath() As String
    MBU_SetupPath = MB_StatePath("setup_" & Left$(MBH_HashText(LCase$(gAutoStoreId)), 24) & ".tsv")
End Function

Private Function MBU_PrepareAccount() As Boolean
    Dim saved As String, fields As Variant, extra As Collection, count As Long
    Dim stream As Object, needRegistration As Boolean, n As Long, reason As String
    gAutoStoreId = Application.Session.DefaultStore.StoreID
    gAutoAccount = MBY_AccountRoot(gAutoStoreId)
    If Len(gAutoAccount) = 0 Or gCancel Then Exit Function
    If gJournalIssues > 0 Then Err.Raise MB_ERR + 650, , "Invalid backup index rows require recovery before setup."
    gAutoKey = "AutomaticSetup_" & MBH_HashText(MB_EncodeFields(Array(LCase$(gRoot), gAutoStoreId, LCase$(gAutoAccount))))
    If gFSO.FileExists(MBU_SetupPath()) Then
        Set stream = gFSO.OpenTextFile(MBU_SetupPath(), 1, False, -1)
        If Not stream.AtEndOfStream Then saved = stream.ReadLine
        stream.Close: Set stream = Nothing
    Else
        saved = MBP_GetSettingText(gAutoKey, "")
    End If
    needRegistration = (Len(saved) = 0 Or gRecoveryForceFull Or gStateWasMissing)
    If Len(saved) > 0 Then
        On Error Resume Next
        Err.Clear
        fields = MB_ReadCheckedLine(saved, 10)
        n = Err.Number: reason = Err.Description
        On Error GoTo 0
        If n <> 0 Then
            needRegistration = True
            MB_LogEvent "WARNING", "SETUP_STATE_REBUILT", "", MBU_SetupPath(), n, reason
        Else
            If fields(0) <> "MBAUTO1" Or StrComp(CStr(fields(1)), gRoot, vbTextCompare) <> 0 Or _
               StrComp(CStr(fields(2)), gAutoStoreId, vbBinaryCompare) <> 0 Or _
               StrComp(CStr(fields(3)), gAutoAccount, vbTextCompare) <> 0 Then
                Err.Raise MB_ERR + 651, , "Automatic setup state belongs to a different archive/account."
            End If
            gAutoSetupPhase = CStr(fields(4)): gFilePlan = CStr(fields(5)): gAutoSetupNote = CStr(fields(6))
            If gAutoSetupPhase <> "APPLY" And gAutoSetupPhase <> "REGISTER" And gAutoSetupPhase <> "READY" Then _
                Err.Raise MB_ERR + 652, , "Unknown setup phase."
            If gAutoSetupPhase <> "READY" Or CStr(fields(9)) <> MBU_ValidationPolicy() Then needRegistration = True
            ' An old aggregate review is seeded into the new per-source ledger
            ' once; after that its issue is retried without rescanning all mail.
            If MB_ByteCount(CStr(fields(8))) > 0 Then needRegistration = True
        End If
    End If
    gFileAccount = gAutoAccount
    gAutoFileReview = 0: gAutoLegacyReview = 0: gAutoReviewCount = 0
    If gAutoSetupPhase = "APPLY" Then
        If Not gFSO.FileExists(gFilePlan) Or gRecoveryForceFull Then gAutoSetupPhase = ""
    End If
    If gAutoSetupPhase <> "APPLY" Then
        gAutoMaintenance = (Not needRegistration)
        gPhase = "Checking existing files for incremental year/month organization"
        gFileExtraReview = 0
        Set extra = MBX_ManagedFiles(gAutoStoreId, gAutoAccount)
        If gCancel Then Exit Function
        gFilePlan = MB_UnusedFinal(MBH_JoinPath(gRunDir, "organize_plan.tsv"))
        count = MBM_BuildPlan(gAutoAccount, gFilePlan, extra)
        If gCancel Then Exit Function
        gAutoFileReview = gFileExtraReview
        MBP_SaveSettingText MBX_PlanKey(gAutoStoreId), gFilePlan
        MBU_SaveSetup "APPLY", gFilePlan, "Checked current filesystem; recurring maintenance plan."
    End If
    gPhase = "Applying/resuming verified year/month moves"
    MBM_RequirePlanAccount gFilePlan, gAutoAccount
    If gCancel Then Exit Function
    MBP_SaveSettingText MBX_PlanKey(gAutoStoreId, True), gFilePlan
    MBM_ApplyPlan gFilePlan
    If gCancel Then Exit Function
    gAutoFileReview = gAutoFileReview + gMovePlanReview + gMoveReview
    gAutoReviewCount = gAutoFileReview
    gAutoSetupNote = "Current organization review items=" & CStr(gAutoFileReview) & "; plan=" & gFilePlan
    MBU_WriteText MB_UnusedFinal(MBH_JoinPath(gRunDir, "organization_summary.txt")), gAutoSetupNote & vbCrLf & MBM_Summary()
    MBX_BuildManagedPaths
    If gCancel Then Exit Function
    gAutoNeedsRegistration = needRegistration
    If needRegistration Then
        MBU_SaveSetup "REGISTER", gFilePlan, gAutoSetupNote
    Else
        MBU_SaveSetup "READY", gFilePlan, gAutoSetupNote
    End If
    MBU_PrepareAccount = True
End Function

Private Function MBU_RegisterExisting(ByVal roots As Collection) As Boolean
    Dim queuePath As String, started As Double, savedMode As String
    Dim savedSince As Date, summary As String, note As String, n As Long, detail As String
    Dim completeScan As Boolean, priorSummaryAt As Long
    On Error GoTo Failed
    savedMode = gMode: savedSince = gSince
    If gLegacyFileCount = 0 Then
        MBU_SaveSetup "READY", gFilePlan, gAutoSetupNote & "; no legacy MSG files to link."
        MBU_RegisterExisting = True
        Exit Function
    End If
    gMode = "REGISTER_LEGACY": gSince = 0
    gPhase = "Initial setup: linking existing MSG files; no mail SaveAs"
    queuePath = MBH_JoinPath(gRunDir, "initial_registration_queue.tsv")
    started = MB_Clock()
    gPlanned = MBP_BuildQueue(roots, 0, gUntil, True, queuePath)
    gFullPlanned = gPlanned
    gScanSeconds = gScanSeconds + MB_Clock() - started
    If Not gCancel Then
        MB_OpenJournal
        started = MB_Clock()
        MB_ProcessQueue queuePath
        gExportSeconds = gExportSeconds + MB_Clock() - started
    End If
    gMailPending = gPlanned - gProcessed
    If gProcessed <> gMailReused + gMailFailed + gMailChanged + gMailPolicy + gMailReview + gLegacyUnmatched Or _
       gMailNew <> 0 Or gMailPending < 0 Then
        Err.Raise MB_ERR + 653, , "Initial no-export registration outcomes do not reconcile."
    End If
    MB_FlushJournal
    If gFSO.FileExists(gJournalPath) Then MBM_FlushPath gJournalPath
    MBW_FinishAudit
    summary = MB_Summary()
    MBU_WriteText MBH_JoinPath(gRunDir, "initial_registration_summary.txt"), summary
    gAutoRegistrationSummary = "Initial linking: planned=" & CStr(gPlanned) & "; linked/reused=" & CStr(gMailReused) & _
        "; unmatched=" & CStr(gLegacyUnmatched) & "; review=" & CStr(gMailReview) & "; failed=" & CStr(gMailFailed) & _
        "; changed=" & CStr(gMailChanged) & "; pending=" & CStr(gMailPending)
    gAutoLegacyReview = gMailReview + gContentReview
    gAutoReviewCount = gAutoFileReview + gAutoLegacyReview
    gAutoFailureCount = gMailFailed + gMailChanged + gFolderFailures + gScanChanged + gMailPending
    completeScan = (Not gCancel And Not gFatal And gAutoFailureCount = 0 And gJournalIssues = 0)
    If gCancel Then Exit Function
    gMode = savedMode: gSince = savedSince
    note = gAutoSetupNote
    priorSummaryAt = InStr(1, note, "; Initial linking:", vbBinaryCompare)
    If priorSummaryAt > 0 Then note = Left$(note, priorSummaryAt - 1)
    note = note & "; " & gAutoRegistrationSummary
    If completeScan Then
        ' READY means the initial pass completed, not proof of a full historical backup.
        ' Unmatched/review items remain explicitly counted and are never force-exported.
        MBU_SaveSetup "READY", gFilePlan, note
    Else
        MBU_SaveSetup "REGISTER", gFilePlan, note
    End If
    ' Save the initial phase separately; the normal summary counts the adaptive pass.
    gFoldersScanned = 0: gFolderFailures = 0: gPolicySkipped = 0: gCandidates = 0: gScanChanged = 0
    gPlanned = 0: gFullPlanned = 0: gProcessed = 0: gMailPending = 0
    gMailNew = 0: gMailReused = 0: gMailFailed = 0: gMailChanged = 0
    gMailPolicy = 0: gMailReview = 0: gMailLegacy = 0: gLegacyUnmatched = 0: gContentReview = 0
    MBU_RegisterExisting = True
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    Err.Raise n, "MBU_RegisterExisting", detail
End Function

Private Sub MBU_WriteText(ByVal path As String, ByVal contents As String)
    Dim output As Object, n As Long, detail As String
    On Error GoTo Failed
    Set output = gFSO.CreateTextFile(path, False, True)
    output.Write contents & vbCrLf
    output.Close
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close
    On Error GoTo 0
    Err.Raise n, "MBU_WriteText", detail
End Sub


' ======================= Native acceptance report ===========================
' Called by the single public command, once per VBA session after a successful
' result. Uses only newly owned local test files and unsent, undisplayed messages.
' No MailItem.Save, Send, real archive migration or original-mail reads occur.
Private Function MBT_RunAcceptanceReport(ByRef reportPath As String, ByRef detail As String) As Boolean
    Dim fso As Object, stream As Object, base As String, runFolder As String
    Dim testDetail As String, envOK As Boolean, helperOK As Boolean, moveOK As Boolean
    Dim plainOK As Boolean, htmlOK As Boolean, propertyOK As Boolean, localeOK As Boolean
    Dim recoveryOK As Boolean, relativeOK As Boolean
    Dim started As Double, failures As Long, savedCalendar As Long, calendarEntered As Boolean
    Dim n As Long, errorText As String, ownsBusy As Boolean
    reportPath = "": detail = ""
    If gBusy Then
        detail = "A backup or native test is already running."
        Exit Function
    End If
    On Error GoTo Failed
    gBusy = True: ownsBusy = True
    MBC_EnterGregorian savedCalendar, calendarEntered
    MB_Reset
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = Environ$("LOCALAPPDATA")
    If Len(base) = 0 Then Err.Raise MB_ERR + 700, , "LOCALAPPDATA is unavailable for the native test report."
    base = MBH_JoinPath(MBH_NormalizePath(base), "OutlookMailBackupValidation")
    MBH_EnsureDirectory base
    runFolder = MBH_JoinPath(base, MBC_FileStamp(Now) & "_" & fso.GetTempName)
    ' CreateFolder must claim a new directory. Never reuse/delete a prior report.
    fso.CreateFolder runFolder
    reportPath = MBH_JoinPath(runFolder, "acceptance.tsv")
    Set stream = fso.CreateTextFile(reportPath, False, True)
    stream.WriteLine "schema" & vbTab & "module_version" & vbTab & "local_time" & vbTab & _
                     "check" & vbTab & "status" & vbTab & "seconds" & vbTab & "detail"
    stream.Close: Set stream = Nothing

    started = Timer
    envOK = MBT_Environment(testDetail)
    MBT_ReportRow reportPath, "OUTLOOK_ENVIRONMENT", MBT_PassFail(envOK), MBT_Elapsed(started), testDetail
    If Not envOK Then failures = failures + 1

    started = Timer
    helperOK = MBT_HelperSelfTest(testDetail)
    If gCancel Then helperOK = False: testDetail = testDetail & " Cancellation requested."
    MBT_ReportRow reportPath, "NATIVE_HELPERS", MBT_PassFail(helperOK), MBT_Elapsed(started), testDetail
    If Not helperOK Then failures = failures + 1

    If helperOK And Not gCancel Then
        started = Timer
        localeOK = MBC_LocaleSelfTest(testDetail)
        MBT_ReportRow reportPath, "LOCALE_GREGORIAN_UTC", MBT_PassFail(localeOK), MBT_Elapsed(started), testDetail
        If Not localeOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "LOCALE_GREGORIAN_UTC", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If

    If helperOK And Not gCancel Then
        started = Timer
        relativeOK = MBX_RelativeSelfTest(MBH_JoinPath(runFolder, "account"), testDetail)
        If gCancel Then relativeOK = False: testDetail = testDetail & " Cancellation requested."
        MBT_ReportRow reportPath, "RELATIVE_JOURNAL_CODEC", MBT_PassFail(relativeOK), MBT_Elapsed(started), testDetail
        If Not relativeOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "RELATIVE_JOURNAL_CODEC", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If

    If helperOK And Not gCancel Then
        started = Timer
        propertyOK = MBT_PropertySelfTest(testDetail)
        MBT_ReportRow reportPath, "MAPI_PROPERTY_CONTRACT", MBT_PassFail(propertyOK), MBT_Elapsed(started), testDetail
        If Not propertyOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "MAPI_PROPERTY_CONTRACT", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If

    If helperOK And propertyOK And Not gCancel Then
        started = Timer
        moveOK = MBT_YearMonthSelfTest(testDetail)
        If gCancel Then moveOK = False: testDetail = testDetail & " Cancellation requested."
        MBT_ReportRow reportPath, "NATIVE_YEAR_MONTH", MBT_PassFail(moveOK), MBT_Elapsed(started), testDetail
        If Not moveOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "NATIVE_YEAR_MONTH", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If

    If helperOK And moveOK And Not gCancel Then
        started = Timer
        recoveryOK = MBS_NativeRecoverySelfTest(testDetail)
        If gCancel Then recoveryOK = False: testDetail = testDetail & " Cancellation requested."
        MBT_ReportRow reportPath, "NATIVE_METADATA_RECOVERY", MBT_PassFail(recoveryOK), MBT_Elapsed(started), testDetail
        If Not recoveryOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "NATIVE_METADATA_RECOVERY", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If

    If envOK And helperOK And propertyOK And moveOK And recoveryOK And Not gCancel Then
        started = Timer
        plainOK = MBT_SyntheticMsgSelfTest(olFormatPlain, testDetail)
        If gCancel Then plainOK = False: testDetail = testDetail & " Cancellation requested."
        MBT_ReportRow reportPath, "SYNTHETIC_MSG_PLAIN", MBT_PassFail(plainOK), MBT_Elapsed(started), testDetail
        If Not plainOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "SYNTHETIC_MSG_PLAIN", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If
    If envOK And helperOK And propertyOK And moveOK And recoveryOK And plainOK And Not gCancel Then
        started = Timer
        htmlOK = MBT_SyntheticMsgSelfTest(olFormatHTML, testDetail)
        If gCancel Then htmlOK = False: testDetail = testDetail & " Cancellation requested."
        MBT_ReportRow reportPath, "SYNTHETIC_MSG_HTML", MBT_PassFail(htmlOK), MBT_Elapsed(started), testDetail
        If Not htmlOK Then failures = failures + 1
    Else
        MBT_ReportRow reportPath, "SYNTHETIC_MSG_HTML", "NOT_RUN", 0, "Prerequisite failed or cancellation requested."
    End If

    ' These gates require the actual account/archive and must not be reported PASS
    ' from a small synthetic test. Running a macro is not a full VBA project compile.
    MBT_ReportRow reportPath, "FULL_VBA_PROJECT_COMPILE", "NOT_RUN", 0, "Debug > Compile is a separate manual gate."
    MBT_ReportRow reportPath, "REAL_ARCHIVE_COMPATIBILITY", "NOT_RUN", 0, "No original mail or archive files were used by these tests."
    MBT_ReportRow reportPath, "SPECIAL_ATTACHMENTS_AND_ALL_MAPI", "NOT_RUN", 0, "Synthetic tests cover ordinary by-value attachments only."
    MBT_ReportRow reportPath, "100000_ITEM_PERFORMANCE", "NOT_RUN", 0, "No large real-account workload was executed."
    MBT_RunAcceptanceReport = envOK And helperOK And localeOK And relativeOK And propertyOK And moveOK And recoveryOK And plainOK And htmlOK And Not gCancel
    detail = "Built-in native checks: " & MBT_PassFail(MBT_RunAcceptanceReport) & _
             "; failed checks=" & CStr(failures) & ". Four operational gates remain NOT_RUN."
    MBT_ReportRow reportPath, "RUN_RESULT", MBT_PassFail(MBT_RunAcceptanceReport), 0, detail
    MBC_LeaveGregorian savedCalendar, calendarEntered
    gBusy = False
    Exit Function
Failed:
    n = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(reportPath) > 0 Then
        MBT_ReportRow reportPath, "RUN_RESULT", "FAIL", 0, "Test runner error " & CStr(n) & ": " & errorText
    End If
    MBH_ReleaseCrypto
    MBC_LeaveGregorian savedCalendar, calendarEntered
    On Error GoTo 0
    If ownsBusy Then gBusy = False
    detail = "Native acceptance could not complete: " & CStr(n) & " - " & errorText
    MBT_RunAcceptanceReport = False
End Function

Private Function MBT_PropertySelfTest(ByRef detail As String) As Boolean
    On Error GoTo Failed
    If Not MBL_PropertySelfTest() Then Err.Raise MB_ERR + 712, , "Optional MAPI property contract test failed."
    If Not MBE_CategorySelfTest() Then Err.Raise MB_ERR + 1125, , "Category metadata contract test failed."
    detail = "PASS: exact category sets, MAPI Error variants, Unicode/ANSI fallback, missing/access failure and Null/Empty/type handling."
    MBT_PropertySelfTest = True
    Exit Function
Failed:
    detail = "FAIL: " & CStr(Err.Number) & " - " & Err.Description
End Function

Private Function MBT_PassFail(ByVal passed As Boolean) As String
    If passed Then MBT_PassFail = "PASS" Else MBT_PassFail = "FAIL"
End Function

Private Function MBT_Elapsed(ByVal started As Double) As Double
    MBT_Elapsed = CDbl(Timer) - started
    If MBT_Elapsed < 0 Then MBT_Elapsed = MBT_Elapsed + 86400#
End Function

Private Sub MBT_ReportRow(ByVal path As String, ByVal check As String, ByVal status As String, _
                          ByVal seconds As Double, ByVal detail As String)
    Dim fso As Object, stream As Object, n As Long, errorText As String
    On Error GoTo Failed
    If status <> "PASS" And status <> "FAIL" And status <> "NOT_RUN" Then
        Err.Raise MB_ERR + 701, , "Unknown acceptance status."
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(path, 8, False, -1)
    stream.WriteLine MB_EncodeFields(Array(MBT_REPORT_SCHEMA, MB_VERSION, MB_Stamp(Now), check, status, _
                      Trim$(Str$(seconds)), detail))
    stream.Close
    Exit Sub
Failed:
    n = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBT_ReportRow", errorText
End Sub

Private Function MBT_Environment(ByRef detail As String) As Boolean
    Dim stores As Long, bitness As String
    On Error GoTo Failed
#If Win64 Then
    bitness = "64"
#Else
    bitness = "32"
#End If
    stores = Application.Session.Stores.Count
    If stores < 1 Then Err.Raise MB_ERR + 702, , "No Outlook store is mounted."
    detail = "Outlook=" & Application.Version & "; VBA bits=" & bitness & _
             "; mounted stores=" & CStr(stores) & "; local time=" & MB_Stamp(Now)
    MBT_Environment = True
    Exit Function
Failed:
    detail = "Environment error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MBT_SyntheticMsgSelfTest(ByVal bodyFormat As Long, ByRef detail As String) As Boolean
    Dim testBase As String, payload As String, msgPath As String, copyPath As String
    Dim stagePath As String, expectedSubject As String, unicodeText As String, expectedHash As String
    Dim source As Outlook.MailItem, reopened As Outlook.MailItem, obj As Object, attachment As Outlook.Attachment
    Dim stream As Object, owned As Boolean, complete As Boolean, n As Long, errorText As String, cleanup As String
#If VBA7 Then
    Dim readGuard As LongPtr
#Else
    Dim readGuard As Long
#End If
    On Error GoTo Failed
    MB_Reset
    gMode = "SELFTEST_MSG": gPhase = "Testing disposable Outlook MSG files"
    gClockStart = MB_Clock(): gRunStart = Now: gUntil = gRunStart
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gDirs = CreateObject("Scripting.Dictionary"): gDirs.CompareMode = vbTextCompare
    testBase = MBH_JoinPath(gFSO.GetSpecialFolder(2).Path, "MBMSG_" & gFSO.GetTempName)
    gFSO.CreateFolder testBase: owned = True
    gRoot = testBase: gBaseRoot = testBase: gRunDir = testBase
    gRunId = "synthetic_msg"
    payload = MBH_JoinPath(testBase, "payload.txt")
    unicodeText = ChrW(&HD55C) & ChrW(&HAE00) & ChrW(&HD83D) & ChrW(&HDE00)
    expectedSubject = "Mail Backup synthetic test " & unicodeText
    Set stream = gFSO.CreateTextFile(payload, False, True)
    stream.Write "Owned synthetic attachment" & vbCrLf & unicodeText & vbTab & "1234567890"
    stream.Close: Set stream = Nothing
    expectedHash = MBH_HashFile(payload)
    Set source = Application.CreateItem(olMailItem)
    source.Subject = expectedSubject
    source.BodyFormat = bodyFormat
    If bodyFormat = olFormatHTML Then
        source.HTMLBody = "<html><body><p>Synthetic HTML " & unicodeText & "</p></body></html>"
    Else
        source.Body = "Synthetic plain text " & unicodeText & vbCrLf & "Second line."
    End If
    Set attachment = source.Attachments.Add(payload, olByValue)
    Set attachment = Nothing
    msgPath = MBH_JoinPath(testBase, "original.msg")
    stagePath = MBH_NewStagePath(msgPath)
    source.SaveAs stagePath, olMSGUnicode
    MBH_CommitStage stagePath, msgPath: stagePath = ""
    source.Close olDiscard: Set source = Nothing
    If gCancel Then Err.Raise MB_ERR + 703, , "Synthetic test cancelled."

    ' Attachment.Type on an unsaved item can report olOLE. Reopen the owned MSG
    ' before validating types or invoking the production content comparator.
    Set obj = Application.Session.OpenSharedItem(msgPath)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 704, , "Synthetic MSG did not reopen as MailItem."
    Set reopened = obj: Set obj = Nothing
    If reopened.Sent Or reopened.Subject <> expectedSubject Or reopened.BodyFormat <> bodyFormat Then
        Err.Raise MB_ERR + 705, , "Synthetic subject/sent-state/body-format did not survive SaveAs."
    End If
    If InStr(1, reopened.Body, unicodeText, vbBinaryCompare) = 0 Then
        Err.Raise MB_ERR + 706, , "Synthetic Unicode body did not survive SaveAs."
    End If
    If reopened.Attachments.Count <> 1 Then Err.Raise MB_ERR + 707, , "Synthetic attachment count changed."
    Set attachment = reopened.Attachments.Item(1)
    If attachment.Type <> olByValue Or attachment.FileName <> "payload.txt" Then
        Err.Raise MB_ERR + 708, , "Synthetic ordinary attachment metadata changed."
    End If
    If MBL_AttachmentHash(attachment) <> expectedHash Then
        Err.Raise MB_ERR + 709, , "Synthetic attachment bytes changed."
    End If
    Set attachment = Nothing
    copyPath = MBH_JoinPath(testBase, "roundtrip.msg")
    stagePath = MBH_NewStagePath(copyPath)
    reopened.SaveAs stagePath, olMSGUnicode
    MBH_CommitStage stagePath, copyPath: stagePath = ""
    readGuard = MBH_GuardFile(copyPath)
    MB_CompareMsg reopened, copyPath, complete
    MBH_CloseGuard readGuard
    If Not complete Then Err.Raise MB_ERR + 710, , "Ordinary synthetic message did not receive full bounded content validation."
    If Not MBI_SyntheticComparatorTest(reopened, copyPath, detail) Then _
        Err.Raise MB_ERR + 713, , "Synthetic restored-identity comparison failed: " & detail
    If Not MBE_SyntheticMetadataTest(reopened, copyPath, detail) Then _
        Err.Raise MB_ERR + 1125, , "Synthetic metadata comparison failed: " & detail
    If gCancel Or gFatal Then Err.Raise MB_ERR + 711, , "Synthetic verification was cancelled or stopped."
    GoTo CleanUp
Failed:
    n = Err.Number: errorText = Err.Description
CleanUp:
    On Error Resume Next
    MBH_CloseGuard readGuard
    If Not stream Is Nothing Then stream.Close
    Set attachment = Nothing
    Err.Clear
    If Not reopened Is Nothing Then reopened.Close olDiscard
    If Err.Number <> 0 Then cleanup = " Reopened MSG close: " & Err.Description
    Set reopened = Nothing: Set obj = Nothing
    Err.Clear
    If Not source Is Nothing Then source.Close olDiscard
    If Err.Number <> 0 Then cleanup = cleanup & " Synthetic source close: " & Err.Description
    Set source = Nothing
    If Len(stagePath) > 0 Then
        Err.Clear
        MBH_DeleteStage stagePath
        If Err.Number <> 0 Then cleanup = cleanup & " Stage cleanup: " & Err.Description
    End If
    Err.Clear
    MBH_ReleaseCrypto
    If Err.Number <> 0 Then cleanup = cleanup & " Crypto cleanup: " & Err.Description
    If owned Then
        Err.Clear
        gFSO.DeleteFolder testBase, True
        If Err.Number <> 0 Then cleanup = cleanup & " Owned test directory preserved: " & testBase & ": " & Err.Description
    End If
    Set gDirs = Nothing: Set gFSO = Nothing
    gRoot = "": gBaseRoot = "": gRunDir = "": gRunId = ""
    On Error GoTo 0
    If n = 0 And Len(cleanup) = 0 And Not gCancel Then
        detail = "PASS: unsent Unicode MSG roundtrip, ordinary attachment bytes, equal-content reuse and changed-body rejection."
        MBT_SyntheticMsgSelfTest = True
    Else
        detail = "FAIL: " & CStr(n) & " - " & errorText & cleanup
    End If
End Function



' ========================== Public entry points =============================
Private Sub MB_BackupRecentMonth()
    MB_Run "MONTH", False, True, False
End Sub

Private Sub MB_BackupModifiedMonth()
    MB_Run "MODIFIED_MONTH", False, True, False
End Sub

Private Sub MB_BackupAll()
    MB_Run "ALL", False, True, False
End Sub

Private Sub MB_BackupAllStoresMonth()
    MB_Run "MONTH", True, True, False
End Sub

Private Sub MB_BackupPickedFolderMonth()
    MB_Run "MONTH", False, True, True
End Sub

Private Sub MB_BackupPickedFolderOnly()
    MB_Run "MONTH", False, False, True
End Sub

Private Sub MB_BackupSelectedMessages()
    MB_Run "SELECTED", False, False, False
End Sub

Private Sub MB_Benchmark300()
    MB_Run "BENCHMARK", False, True, False
End Sub

Private Sub MB_VerifyAll()
    MB_Run "VERIFY", False, False, False
End Sub

Private Sub MB_CancelBackup()
    If gBusy Then gCancel = True
End Sub

Private Sub MB_Run(ByVal mode As String, ByVal allStores As Boolean, _
                   ByVal includeChildren As Boolean, ByVal pickFolder As Boolean)
    Dim roots As Collection, folder As Outlook.Folder, st As Outlook.Store
    Dim ex As Outlook.Explorer, selected As Outlook.Selection
    Dim baseRoot As String, queuePath As String, summary As String
    Dim startTick As Double, errNo As Long, errText As String, rootIndex As Long
    Dim lockHeld As Boolean, attached As String, autoReady As Boolean
    Dim savedCalendar As Long, calendarEntered As Boolean
    If gBusy Then
        MsgBox "A backup/verification is already running. Use MB_CancelBackup.", vbInformation
        Exit Sub
    End If
    On Error GoTo Failed
    gBusy = True
    MBC_EnterGregorian savedCalendar, calendarEntered
    MB_Reset
    gMode = mode
    gRunStart = Now
    gUntil = gRunStart
    gClockStart = MB_Clock()
    If mode = "MONTH" Or mode = "MODIFIED_MONTH" Then gSince = MBC_ShiftMonths(gUntil, -1)
    If mode = "AUTO" Then MBW_Start Application.Session.DefaultStore.StoreID
    gModifiedWindow = (mode = "MODIFIED_MONTH")
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gIndex = CreateObject("Scripting.Dictionary")
    gIndex.CompareMode = vbBinaryCompare
    Set gExternalRoots = CreateObject("Scripting.Dictionary")
    Set gRecordMeta = CreateObject("Scripting.Dictionary")
    Set gStoreTokens = CreateObject("Scripting.Dictionary")
    Set gManagedFilePaths = CreateObject("Scripting.Dictionary")
    Set gLegacyAccountStats = CreateObject("Scripting.Dictionary")
    gStoreTokens.CompareMode = vbTextCompare
    gManagedFilePaths.CompareMode = vbTextCompare
    gLegacyAccountStats.CompareMode = vbTextCompare
    Set gLegacyBindings = CreateObject("Scripting.Dictionary")
    Set gLegacySources = CreateObject("Scripting.Dictionary")
    Set gLegacyOwners = CreateObject("Scripting.Dictionary")
    gLegacyBindings.CompareMode = vbBinaryCompare
    gLegacySources.CompareMode = vbBinaryCompare
    gLegacyOwners.CompareMode = vbTextCompare
    Set gDirs = CreateObject("Scripting.Dictionary")
    gDirs.CompareMode = vbTextCompare
    Set roots = New Collection
    If mode <> "VERIFY" And Not MBX_IsFileMode(mode) Then
        If mode = "SELECTED" Then
            Set ex = Application.ActiveExplorer
            If ex Is Nothing Then Err.Raise MB_ERR + 1, "MB_Run", "No active Explorer."
            Set selected = ex.Selection
            If selected.Count = 0 Then Err.Raise MB_ERR + 2, "MB_Run", "No selected items."
            gScope = "Current Explorer selection"
        ElseIf pickFolder Then
            Set folder = Application.Session.PickFolder
            If folder Is Nothing Then
                gCancel = True
                GoTo CleanUp
            End If
            roots.Add folder
            gScope = folder.FolderPath
        ElseIf allStores Then
            For Each st In Application.Session.Stores
                Set folder = st.GetRootFolder
                roots.Add folder
            Next st
            gScope = "All currently mounted stores (" & CStr(roots.Count) & ")"
        Else
            Set folder = Application.Session.DefaultStore.GetRootFolder
            roots.Add folder
            gScope = folder.FolderPath
        End If
    Else
        gScope = "Every committed file in the selected managed archive"
    End If
    baseRoot = MBP_ResolveBackupRoot(gSelectedBase)
    If Len(baseRoot) = 0 Then
        gCancel = True
        GoTo CleanUp
    End If
    ' Recheck after selection: the derived root/ancestors may have changed
    ' since preflight. No link target may be initialized as an owned archive.
    MBP_NoReparse baseRoot, gFSO
    MBH_EnsureDirectory baseRoot
    MBP_NoReparse MBH_JoinPath(baseRoot, "backup.lock"), gFSO
    MBH_AcquireRunLock MBH_JoinPath(baseRoot, "backup.lock")
    lockHeld = True
    gRoot = baseRoot
    gBaseRoot = baseRoot
    MBP_NoReparse baseRoot, gFSO
    MB_CheckOwnership gRoot
    gLegacyRoot = MBP_ResolveLegacyRoot(gSelectedBase, gBaseRoot)
    If gFSO.FileExists(MBH_JoinPath(baseRoot, "STOP.txt")) Then
        Err.Raise MB_ERR + 4, "MB_Run", "STOP.txt exists. Remove it before starting a new run."
    End If
    gRunId = MBC_FileStamp(gRunStart) & "_" & _
             Left$(MBH_HashText(gFSO.GetTempName & CStr(Timer)), 10)
    If mode = "BENCHMARK" Then
        gRoot = MBH_JoinPath(MBH_JoinPath(baseRoot, "_benchmarks"), gRunId)
        MBH_EnsureDirectory gRoot
        MB_CheckOwnership gRoot
    End If
    If mode = "AUTO" Then MBS_BindStore Application.Session.DefaultStore.StoreID
    MBP_PersistBackupRoot gSelectedBase
    MBS_Initialize
    gRunDir = MBH_JoinPath(MB_StatePath("runs"), gRunId)
    MB_EnsureDir gRunDir
    Set gEvents = gFSO.CreateTextFile(MBH_JoinPath(gRunDir, "events.tsv"), False, True)
    gEvents.WriteLine "time" & vbTab & "level" & vbTab & "stage" & vbTab & _
                      "key" & vbTab & "path" & vbTab & "error" & vbTab & "detail"
    MB_LogEvent "INFO", "START", "", gRoot, 0, _
        "version=" & MB_VERSION & "; mode=" & mode & "; scope=" & gScope & _
        "; since=" & MB_Stamp(gSince) & "; until=" & MB_Stamp(gUntil) & _
        "; descendants=" & CStr(includeChildren) & _
        "; extract=" & CStr(MB_EXTRACT_ATTACHMENTS) & _
        "; rehash_existing=" & CStr(MB_VERIFY_EXISTING_HASH) & _
        "; selected_base=" & gSelectedBase & "; legacy_root=" & gLegacyRoot
    gJournalPath = MB_StatePath("journal.tsv")
    If mode = "AUTO" Then
        gPhase = "Recovering operating records before loading the index"
        gEarlyAccount = MBS_EarlyAccount(Application.Session.DefaultStore.StoreID)
        If gCancel Then GoTo Finish
        MBR_RestoreState gEarlyAccount, Application.Session.DefaultStore.StoreID
        MBS_CheckRestoredStore Application.Session.DefaultStore.StoreID
        MBY_Reset
    End If
    gPhase = "Loading committed-file index"
    startTick = MB_Clock()
    MB_LoadIndex
    If mode = "AUTO" And gJournalIssues > 0 Then
        MB_LogEvent "WARNING", "SEMANTIC_INDEX_RECOVERY", "", gJournalPath, 0, _
                    "Checked rows failed semantic validation; restoring/rebuilding state once."
        MBR_RestoreState gEarlyAccount, Application.Session.DefaultStore.StoreID, True
        MBS_CheckRestoredStore Application.Session.DefaultStore.StoreID
        MBS_ClearLoadedIndex Application.Session.DefaultStore.StoreID
        MB_LoadIndex
        If gJournalIssues > 0 Then Err.Raise MB_ERR + 926, , "Recovered index still contains invalid semantic records."
    End If
    gIndexSeconds = MB_Clock() - startTick
    If gCancel Then GoTo Finish
    gPhase = "Checking recoverable file moves"
    MBM_LoadMoves
    If gCancel Then GoTo Finish
    If mode = "AUTO" Then
        gEarlyAccount = MBY_ValidateAccountRoot(gEarlyAccount, Application.Session.DefaultStore.StoreID, True)
        If gCancel Or Len(gEarlyAccount) = 0 Then GoTo Finish
        MBY_SaveRoot Application.Session.DefaultStore.StoreID, gEarlyAccount
    End If
    MBX_BuildManagedPaths
    If gCancel Then GoTo Finish
    If MBX_IsFileMode(mode) Then
        MBX_RunFileMode
        GoTo Finish
    End If
    If mode = "VERIFY" Then
        gPhase = "Verifying indexed files"
        startTick = MB_Clock()
        MB_AuditIndex
        gExportSeconds = MB_Clock() - startTick
        GoTo Finish
    End If
    If mode <> "BENCHMARK" Then
        gPhase = "Finding existing MSG files (no export)"
        startTick = MB_Clock()
        MBX_ResolveAccounts roots, selected
        If gCancel Then GoTo Finish
        If mode = "AUTO" Then
            MBS_RememberAccount Application.Session.DefaultStore.StoreID, gEarlyAccount
            MBS_EnableStages gEarlyAccount, Application.Session.DefaultStore.StoreID
            If gCancel Then GoTo Finish
            startTick = MB_Clock()
            autoReady = MBU_PrepareAccount()
            gAutoSetupSeconds = MB_Clock() - startTick
            If Not autoReady Then GoTo Finish
            startTick = MB_Clock()
            MBR_BuildBaseline gAutoAccount, gAutoStoreId
            If gCancel Then GoTo Finish
            If MBS_RelinkMissingReferences(gAutoAccount, gAutoStoreId) > 0 Then
                MBX_BuildManagedPaths
                autoReady = MBU_PrepareAccount()
                If Not autoReady Then GoTo Finish
            End If
            gMaintenanceReady = True
        End If
        MBL_BuildAccountIndexes gArchiveRoots, gBaseRoot
        gLegacySeconds = MB_Clock() - startTick
        If gCancel Then GoTo Finish
        If Not gLegacyScanComplete Then
            Err.Raise MB_ERR + 273, "MB_Run", "Old archive scan is incomplete; new exports are blocked."
        End If
        MBP_PersistLegacyRoot gBaseRoot, gLegacyRoot
        MB_LogEvent "INFO", "LEGACY_SCAN", "", gLegacyRoot, 0, _
                    "files=" & CStr(gLegacyFileCount) & "; unrecognized_names=" & _
                    CStr(gLegacyUnrecognizedNames) & "; newest_file_write=" & MB_Stamp(gLegacyNewestFileTime)
    End If
    If mode = "AUTO" Then
        MBW_Prepare
        If gCancel Then GoTo Finish
        If gAutoNeedsRegistration Then
            If Not MBU_RegisterExisting(roots) Then GoTo Finish
            MBW_Prepare
            If gCancel Then GoTo Finish
        End If
    End If
    queuePath = MBH_JoinPath(gRunDir, "queue.tsv")
    gPhase = "Collecting fixed target list"
    startTick = MB_Clock()
    If mode = "SELECTED" Then
        gPlanned = MB_QueueSelection(selected, queuePath)
    Else
        gPlanned = MBP_BuildQueue(roots, gSince, gUntil, includeChildren, queuePath)
    End If
    If mode = "AUTO" Then
        gWindowInventoryComplete = (Not gCancel And Not gFatal And gFolderFailures = 0 And gScanChanged = 0)
    End If
    gScanSeconds = gScanSeconds + MB_Clock() - startTick
    gFullPlanned = gPlanned
    If gCancel Then GoTo Finish
    If mode = "BENCHMARK" Then
        queuePath = MB_MakeSample(queuePath, MB_BENCHMARK_COUNT)
    End If
    MB_OpenJournal
    gPhase = "Exporting / verifying messages"
    startTick = MB_Clock()
    MB_ProcessQueue queuePath
    gExportSeconds = gExportSeconds + MB_Clock() - startTick
Finish:
    gMailPending = gPlanned - gProcessed
    If gMailPending < 0 Then
        Err.Raise MB_ERR + 5, "MB_Run", "Internal count mismatch: processed exceeds plan."
    End If
    If mode <> "VERIFY" And Not MBX_IsFileMode(mode) Then
        If gProcessed <> gMailNew + gMailReused + gMailFailed + gMailChanged + gMailPolicy + gMailReview + gLegacyUnmatched Then
            Err.Raise MB_ERR + 6, "MB_Run", "Internal message outcome reconciliation failed."
        End If
        If gAttFound <> gAttNew + gAttReused + gAttExcluded + gAttFailed + gAttPending Then
            Err.Raise MB_ERR + 7, "MB_Run", "Internal attachment outcome reconciliation failed."
        End If
    End If
    If mode = "AUTO" And gMaintenanceReady Then
        MBW_FinishAudit
        If Not gCancel And Not gFatal Then MBA_Run
        ' Current unresolved mail is itemized in the source-audit ledger.
        ' Do not perpetuate an old aggregate registration-review counter.
        If gWindowInventoryComplete Then
            gAutoLegacyReview = 0
            gAutoReviewCount = gAutoFileReview
            If gAutoSetupPhase = "READY" Then MBU_SaveSetup "READY", gFilePlan, gAutoSetupNote
        End If
    End If
    MB_FlushJournal
    If mode = "AUTO" Then MBW_PrepareCommit
    summary = MB_Summary()
    MB_LogEvent "INFO", "END", "", gRunDir, 0, Replace(summary, vbCrLf, " | ")
    GoTo CleanUp
Failed:
    errNo = Err.Number
    errText = Err.Description
    gFatal = True
    gFatalText = "Error " & CStr(errNo) & ": " & errText
    On Error Resume Next
    If Not gEvents Is Nothing Then MB_LogEvent "ERROR", "FATAL", "", gRunDir, errNo, errText
    gMailPending = gPlanned - gProcessed
    If gMailPending < 0 Then gMailPending = 0
    summary = MB_Summary()
    On Error GoTo 0
CleanUp:
    On Error Resume Next
    MBC_LeaveGregorian savedCalendar, calendarEntered
    If Err.Number <> 0 Then gCleanupIssue = gCleanupIssue & " Calendar restore: " & Err.Description
    On Error GoTo 0
    MB_CloseResources
    ' Coverage advances under the same locks, only after the durable message
    ' journal and both streams closed successfully. Interrupted work replays.
    On Error Resume Next
    Err.Clear
    If mode = "AUTO" Then MBW_CommitAfterClose
    If mode = "AUTO" And gWindowCommitted Then MBR_CompleteReconcile gAutoAccount, gAutoStoreId
    If Err.Number <> 0 Then
        gFatal = True
        gFatalText = gFatalText & " UTC checkpoint: " & Err.Description
    End If
    On Error GoTo 0
    If mode = "AUTO" And Len(gRunDir) > 0 Then summary = MB_Summary()
    If mode = "AUTO" And gMaintenanceReady And gJournalIssues = 0 Then
        On Error Resume Next
        Err.Clear
        MBR_SaveSnapshot gAutoAccount, gAutoStoreId
        If Err.Number <> 0 Then
            gFatal = True
            gFatalText = gFatalText & " Recovery snapshot: " & Err.Description
        End If
        On Error GoTo 0
        summary = MB_Summary()
    End If
    gStageEnabled = False
    On Error Resume Next
    Err.Clear
    MBM_ReleaseAccountLocks
    If Err.Number <> 0 Then gCleanupIssue = gCleanupIssue & " Account lock release: " & Err.Description
    On Error GoTo 0
    If lockHeld Then
        On Error Resume Next
        Err.Clear
        MBH_ReleaseRunLock
        If Err.Number <> 0 Then gCleanupIssue = gCleanupIssue & " Lock release: " & Err.Description
        Err.Clear
        MBH_ReleaseCrypto
        If Err.Number <> 0 Then gCleanupIssue = gCleanupIssue & " Crypto release: " & Err.Description
        On Error GoTo 0
    End If
    If Len(gCleanupIssue) > 0 Then
        gFatal = True
        gFatalText = gFatalText & " CLEANUP ERROR: " & gCleanupIssue
        summary = MB_Summary()
    End If
    If Len(summary) = 0 Then
        If gCancel Then
            summary = "Cancelled before export."
        Else
            summary = gFatalText
        End If
    End If
    ' The durable summary is written only after journal/event streams have closed.
    If Len(gRunDir) > 0 Then
        On Error Resume Next
        Err.Clear
        MB_WriteSummary summary
        If Err.Number <> 0 Then
            gFatal = True
            summary = "FAILED TO WRITE FINAL SUMMARY: " & Err.Description & vbCrLf & summary
        End If
        On Error GoTo 0
    End If
    MBL_Reset
    MBY_Reset
    MBM_Reset
    MBI_Reset
    MBW_Reset
    MBA_Reset
    MBR_Reset
    Set gExternalRoots = Nothing: Set gRecordMeta = Nothing: Set gStoreTokens = Nothing
    Set gManagedFilePaths = Nothing: Set gLegacyAccountStats = Nothing: Set gArchiveRoots = Nothing
    Set gLegacyBindings = Nothing: Set gLegacySources = Nothing: Set gLegacyOwners = Nothing
    Set gIndex = Nothing: Set gDirs = Nothing: Set gFSO = Nothing
    On Error Resume Next
    MBC_LeaveGregorian savedCalendar, calendarEntered
    On Error GoTo 0
    gBusy = False
    Debug.Print summary
    If Len(summary) > 900 Then
        summary = Left$(summary, 350) & vbCrLf & "Phase: " & Left$(gPhase, 100)
        If gFatal Then summary = summary & vbCrLf & "Cause: " & Left$(gFatalText, 250)
        summary = summary & vbCrLf & "Full results: " & gRunDir & "\summary.txt"
    End If
    MsgBox summary, IIf(gFatal Or Len(gCleanupIssue) > 0, vbExclamation, vbInformation), _
           "Outlook Mail Backup " & MB_VERSION
End Sub

Private Sub MB_Reset()
    gCancel = False: gFatal = False: gFatalText = "": gCleanupIssue = ""
    gBaseRoot = "": gSelectedBase = "": gLegacyRoot = ""
    gMailLegacy = 0: gMailReview = 0: gLegacyUnmatched = 0: gLegacySeconds = 0
    gFilePlan = "": gFileAccount = "": gFileDetail = "": gFileExtraReview = 0
    gReuseBlocked = False: gReuseReason = ""
    MBP_Reset
    MBF_Reset
    MBU_Reset
    MBS_Reset
    MBR_Reset
    MBA_Reset
    MBI_Reset
    MBW_Reset
    MBL_Reset
    MBY_Reset
    MBM_Reset
    gRoot = "": gRunId = "": gRunDir = "": gSince = 0: gUntil = 0
    gPhase = "Starting": gMode = "": gScope = "": gModifiedWindow = False
    gLastPulse = 0: gLastProgress = 0
    MBH_PathCacheHits = 0: MBH_PathCacheMisses = 0
    gJournalWrites = 0: gJournalIssues = 0
    gFoldersScanned = 0: gFolderFailures = 0: gPolicySkipped = 0
    gCandidates = 0: gScanChanged = 0
    gPlanned = 0: gFullPlanned = 0: gProcessed = 0
    gMailNew = 0: gMailReused = 0: gMailFailed = 0
    gMailChanged = 0: gMailPolicy = 0: gMailPending = 0: gContentReview = 0
    gAttFound = 0: gAttNew = 0: gAttReused = 0
    gAttExcluded = 0: gAttFailed = 0: gAttPending = 0
    gBytesNew = 0: gScanSeconds = 0: gIndexSeconds = 0: gExportSeconds = 0
    gSampleCount = 0
    ReDim gSampleSeconds(1 To MB_BENCHMARK_COUNT)
    Set gEvents = Nothing: Set gJournal = Nothing
End Sub

Private Sub MB_CheckOwnership(ByVal rootPath As String)
    Dim p As String, ts As Object, value As String, f As Object, subFolder As Object
    p = MBH_JoinPath(rootPath, "owner.txt")
    If Not gFSO.FileExists(p) Then
        If MBS_CanRecoverOwner(rootPath, gFSO) Then
            MBS_WriteAtomic p, MB_OWNER & vbCrLf
            Exit Sub
        End If
    End If
    If gFSO.FileExists(p) Then
        Set ts = gFSO.OpenTextFile(p, 1, False, -1)
        value = ts.ReadAll
        ts.Close
        If value <> MB_OWNER & vbCrLf And value <> MB_OWNER_LEGACY & vbCrLf And _
           value <> MB_OWNER_PREVIOUS & vbCrLf And value <> MB_OWNER_V34 & vbCrLf Then
            Err.Raise MB_ERR + 10, "MB_CheckOwnership", "Output owner marker is invalid: " & p
        End If
        If value <> MB_OWNER & vbCrLf Then
            ' Preserve the old marker, then reject unsupported downgrade writers.
            If Not gFSO.FileExists(MBH_JoinPath(rootPath, "owner-previous-" & Left$(MBH_HashText(value), 16) & ".txt")) Then
                Set ts = gFSO.CreateTextFile(MBH_JoinPath(rootPath, "owner-previous-" & Left$(MBH_HashText(value), 16) & ".txt"), False, True)
                ts.Write value: ts.Close
            End If
            MBS_WriteAtomic p, MB_OWNER & vbCrLf
        End If
        Exit Sub
    End If
    For Each f In gFSO.GetFolder(rootPath).Files
        If LCase$(f.Name) <> "backup.lock" Then
            Err.Raise MB_ERR + 11, "MB_CheckOwnership", _
                      "Unowned output folder contains files. Choose a new empty managed backup folder."
        End If
    Next f
    For Each subFolder In gFSO.GetFolder(rootPath).SubFolders
        Err.Raise MB_ERR + 12, "MB_CheckOwnership", _
                  "Unowned output folder contains directories. Choose a new empty managed backup folder."
    Next subFolder
    Set ts = gFSO.CreateTextFile(p, False, True)
    ts.WriteLine MB_OWNER
    ts.Close
End Sub

' ======================= Journal / metadata utilities =======================
Private Function MB_Stamp(ByVal value As Date) As String
    MB_Stamp = MBC_Stamp(value)
End Function

Private Function MB_ParseStamp(ByVal value As String) As Date
    MB_ParseStamp = MBC_ParseLegacyStamp(value)
End Function

Private Function MB_Escape(ByVal value As String) As String
    value = Replace(value, "%", "%25")
    value = Replace(value, vbTab, "%09")
    value = Replace(value, vbCr, "%0D")
    value = Replace(value, vbLf, "%0A")
    ' Avoid formula interpretation when human-readable TSV is opened in Excel.
    If Len(value) > 0 Then
        Select Case Left$(value, 1)
            Case "=": value = "%3D" & Mid$(value, 2)
            Case "+": value = "%2B" & Mid$(value, 2)
            Case "-": value = "%2D" & Mid$(value, 2)
            Case "@": value = "%40" & Mid$(value, 2)
        End Select
    End If
    MB_Escape = value
End Function

Private Function MB_Unescape(ByVal value As String) As String
    value = Replace(value, "%09", vbTab)
    value = Replace(value, "%0D", vbCr)
    value = Replace(value, "%0A", vbLf)
    value = Replace(value, "%3D", "=")
    value = Replace(value, "%2B", "+")
    value = Replace(value, "%2D", "-")
    value = Replace(value, "%40", "@")
    MB_Unescape = Replace(value, "%25", "%")
End Function

Private Function MB_EncodeFields(ByVal values As Variant) As String
    Dim i As Long, fields() As String
    ReDim fields(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        fields(i) = MB_Escape(CStr(values(i)))
    Next i
    MB_EncodeFields = Join(fields, vbTab)
End Function

Private Sub MB_WriteCheckedLine(ByVal stream As Object, ByVal payload As String)
    stream.WriteLine payload & vbTab & MBH_HashText(payload) & vbTab & "END"
End Sub

Private Function MB_ReadCheckedLine(ByVal line As String, ByVal fieldCount As Long) As Variant
    Dim raw As Variant, fields() As String, payload As String, i As Long
    raw = Split(line, vbTab)
    If UBound(raw) <> fieldCount + 1 Then
        Err.Raise MB_ERR + 15, "MB_ReadCheckedLine", "Wrong journal/queue field count."
    End If
    If raw(fieldCount + 1) <> "END" Then
        Err.Raise MB_ERR + 16, "MB_ReadCheckedLine", "Incomplete journal/queue record."
    End If
    For i = 0 To fieldCount - 1
        If i > 0 Then payload = payload & vbTab
        payload = payload & CStr(raw(i))
    Next i
    If MBH_HashText(payload) <> CStr(raw(fieldCount)) Then
        Err.Raise MB_ERR + 17, "MB_ReadCheckedLine", "Journal/queue checksum mismatch."
    End If
    ReDim fields(0 To fieldCount - 1)
    For i = 0 To fieldCount - 1
        fields(i) = MB_Unescape(CStr(raw(i)))
    Next i
    MB_ReadCheckedLine = fields
End Function

Private Function MB_SourceKey(ByVal storeId As String, ByVal entryId As String) As String
    If Len(storeId) = 0 Or Len(entryId) = 0 Then
        Err.Raise MB_ERR + 18, "MB_SourceKey", "Missing source identity."
    End If
    storeId = LCase$(storeId): entryId = LCase$(entryId)
    MB_SourceKey = MBH_HashText("MBSOURCE2|" & CStr(Len(storeId)) & ":" & storeId & _
                              "|" & CStr(Len(entryId)) & ":" & entryId)
End Function

Private Function MB_IsHexKey(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    If Len(value) <> 64 Then Exit Function
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If InStr(1, "0123456789abcdef", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    MB_IsHexKey = True
End Function

Private Function MB_ByteCount(ByVal value As String) As Double
    Dim i As Long, ch As String
    If Len(value) = 0 Or Len(value) > 16 Then
        Err.Raise MB_ERR + 19, "MB_ByteCount", "Invalid byte count."
    End If
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch < "0" Or ch > "9" Then Err.Raise MB_ERR + 19, "MB_ByteCount", "Invalid byte count."
    Next i
    MB_ByteCount = CDbl(value)
End Function

Private Function MB_RelPath(ByVal relativePath As String) As String
    Dim parts As Variant, part As Variant, p As String
    parts = Split(relativePath, "\")
    If UBound(parts) < 1 Then Err.Raise MB_ERR + 20, "MB_RelPath", "Invalid relative output path."
    If parts(0) <> "messages" And parts(0) <> "attachments" Then
        Err.Raise MB_ERR + 21, "MB_RelPath", "Unexpected output namespace."
    End If
    p = gRoot
    For Each part In parts
        p = MBH_JoinPath(p, CStr(part))
    Next part
    MB_RelPath = p
End Function

Private Sub MB_EnsureDir(ByVal directoryPath As String)
    If gDirs.Exists(directoryPath) Then Exit Sub
    MBH_EnsureDirectory directoryPath
    gDirs.Add directoryPath, True
End Sub

Private Sub MB_LoadIndex()
    Dim stream As Object, line As String, rowNumber As Long, errNo As Long, errText As String
    On Error GoTo Failed
    If Not gFSO.FileExists(gJournalPath) Then Exit Sub
    If CDbl(gFSO.GetFile(gJournalPath).Size) Mod 2 <> 0 Then
        Err.Raise MB_ERR + 22, "MB_LoadIndex", "Journal has an odd UTF-16 byte count; preserve and inspect it."
    End If
    Set stream = gFSO.OpenTextFile(gJournalPath, 1, False, -1)
    Do Until stream.AtEndOfStream
        If Not MB_ContinueWork() Then Exit Do
        line = stream.ReadLine
        rowNumber = rowNumber + 1
        If Len(line) > 0 Then MB_LoadIndexRow line, rowNumber
    Loop
    stream.Close
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise errNo, "MB_LoadIndex", errText
End Sub

Private Sub MB_LoadIndexRow(ByVal line As String, ByVal rowNumber As Long)
    Dim f As Variant, errNo As Long, errText As String
    On Error GoTo Invalid
    f = MBR_ParseIndexFields(line)
    If Not MB_IsHexKey(f(2)) Or Not MB_IsHexKey(f(5)) Or Not MB_IsHexKey(f(6)) Then
        Err.Raise MB_ERR + 25, , "Invalid index key/hash."
    End If
    MBX_CacheRecord f
    Exit Sub
Invalid:
    errNo = Err.Number: errText = Err.Description
    gJournalIssues = gJournalIssues + 1
    MB_LogEvent "ERROR", "INVALID_INDEX_ROW", CStr(rowNumber), gJournalPath, errNo, errText
End Sub

Private Sub MB_OpenJournal()
    If gJournal Is Nothing Then
        Set gJournal = gFSO.OpenTextFile(gJournalPath, 8, True, -1)
        ' Leading newline separates a possibly truncated previous final line.
        gJournal.WriteLine ""
        gJournalWrites = 0
    End If
End Sub

Private Sub MB_FlushJournal()
    If Not gJournal Is Nothing Then
        gJournal.Close
        Set gJournal = Nothing
    End If
    gJournalWrites = 0
End Sub

Private Sub MB_RecordCommit(ByVal kind As String, ByVal key As String, ByVal relativePath As String, _
                            ByVal bytes As Double, ByVal digest As String, ByVal q As Variant, _
                            ByVal subject As String, ByVal attNo As Long, _
                            ByVal originalName As String, ByVal validation As String, _
                            Optional ByVal accountRoot As String = "")
    Dim payload As String, fields As Variant, indexKey As String
    On Error GoTo Failed
    If kind = "N" Or kind = "T" Or kind = "R" Then
        fields = Array("MBS4", kind, key, MBX_AccountRelative(relativePath, accountRoot), Format$(bytes, "0"), digest, _
                       q(7), q(1), q(2), q(3), q(4), q(6), subject, CStr(attNo), originalName, validation, accountRoot)
    ElseIf kind = "L" Then
        fields = Array("MBS4", kind, key, MBX_AccountRelative(relativePath, originalName), Format$(bytes, "0"), digest, _
                       q(7), q(1), q(2), q(3), q(4), q(6), subject, CStr(attNo), "", validation, originalName)
    Else
        fields = Array("MBS2", kind, key, relativePath, Format$(bytes, "0"), digest, _
                       q(7), q(1), q(2), q(3), q(4), q(6), subject, CStr(attNo), originalName, validation)
    End If
    ' Reject an incompatible repeat before it becomes durable journal history.
    MBR_RejectIndexConflict fields
    payload = MB_EncodeFields(fields)
    MB_OpenJournal
    MB_WriteCheckedLine gJournal, payload
    gJournalWrites = gJournalWrites + 1
    If gJournalWrites >= MB_JOURNAL_FLUSH_EVERY Then MB_FlushJournal
    MBX_CacheRecord fields
    indexKey = MBX_IndexKey(kind, key)
    If kind <> "L" Then gManagedFilePaths(MBX_IndexedPath(indexKey, MB_IndexFields(indexKey))) = True
    Exit Sub
Failed:
    gFatal = True
    gFatalText = "Cannot record committed file: " & Err.Description
    Err.Raise MB_ERR + 27, "MB_RecordCommit", gFatalText
End Sub

Private Function MB_IndexFields(ByVal key As String) As Variant
    Dim parts As Variant, i As Long
    parts = Split(CStr(gIndex(key)), vbTab)
    If UBound(parts) <> 3 Then Err.Raise MB_ERR + 28, , "Invalid compact index state."
    For i = 0 To 3
        parts(i) = MB_Unescape(CStr(parts(i)))
    Next i
    MB_IndexFields = parts
End Function

Private Function MB_ExistingValid(ByVal key As String, ByRef relativePath As String) As Boolean
    Dim fields As Variant, p As String, reason As String, errNo As Long
    On Error GoTo Invalid
    gReuseBlocked = False: gReuseReason = ""
    If Not gIndex.Exists(key) Then Exit Function
    fields = MB_IndexFields(key)
    relativePath = CStr(fields(0))
    p = MBX_IndexedPath(key, fields, True)
    If Not gFSO.FileExists(p) Then
        reason = "Indexed file is missing.": GoTo InvalidResult
    End If
    If CDbl(gFSO.GetFile(p).Size) <> MB_ByteCount(CStr(fields(1))) Then
        reason = "Indexed file size mismatch.": GoTo InvalidResult
    End If
    If Left$(key, 2) = "M|" And MB_REOPEN_NEW_MSG Then
        If fields(3) <> "REOPENED_SHA256" And fields(3) <> MB_NEW_CONTENT_VALIDATION And _
           fields(3) <> MB_NEW_BASIC_VALIDATION And fields(3) <> MB_NEW_CONTENT_VALIDATION_V2 And _
           fields(3) <> MB_NEW_BASIC_VALIDATION_V2 Then
            reason = "Existing MSG does not have required reopen validation.": GoTo InvalidResult
        End If
    End If
    If MB_VERIFY_EXISTING_HASH Then
        If MBH_HashFile(p) <> CStr(fields(2)) Then
            reason = "Indexed SHA-256 mismatch.": GoTo InvalidResult
        End If
    End If
    MB_ExistingValid = True
    Exit Function
Invalid:
    errNo = Err.Number: reason = Err.Description
InvalidResult:
    MB_LogEvent "WARNING", "INDEXED_FILE_REVIEW_NO_REEXPORT", key, p, errNo, reason
    ' Keep the recorded identity, including moved/missing files. No automatic SaveAs.
    gReuseBlocked = True: gReuseReason = reason
    relativePath = ""
End Function

' =========================== Exporting messages =============================
Private Sub MB_ProcessQueue(ByVal queuePath As String)
    Dim stream As Object, line As String, q As Variant, result As Long
    Dim tick As Double, elapsed As Double, errNo As Long, errText As String
    Dim priorContentReview As Long, metricInFlight As Boolean
    On Error GoTo Failed
    Set stream = gFSO.OpenTextFile(queuePath, 1, False, -1)
    Do Until stream.AtEndOfStream
        If Not MB_ContinueWork() Then Exit Do
        line = stream.ReadLine
        If Len(line) > 0 Then
            q = MB_ReadCheckedLine(line, 8)
            If q(0) <> "MBQ2" Then Err.Raise MB_ERR + 30, , "Wrong queue schema."
            If MB_SourceKey(q(1), q(2)) <> q(7) Then Err.Raise MB_ERR + 31, , "Queue identity mismatch."
            tick = MB_Clock()
            priorContentReview = gContentReview
            result = 0: metricInFlight = True
            result = MB_ProcessOne(q)
            elapsed = MB_Clock() - tick
            metricInFlight = False
            MBF_Record gMode, result, elapsed, (gContentReview > priorContentReview)
            If result = 5 Then Exit Do
            MBW_RecordAuditResult CStr(q(7)), result, (gContentReview > priorContentReview)
            gProcessed = gProcessed + 1
            Select Case result
                Case 1: gMailNew = gMailNew + 1
                Case 2: gMailReused = gMailReused + 1
                Case 3: gMailFailed = gMailFailed + 1
                Case 4: gMailChanged = gMailChanged + 1
                Case 6: gMailPolicy = gMailPolicy + 1
                Case 7: gMailReview = gMailReview + 1
                Case 8: gLegacyUnmatched = gLegacyUnmatched + 1
                Case Else: Err.Raise MB_ERR + 32, , "Unknown message result."
            End Select
            If gMode = "BENCHMARK" And result = 1 Then
                gSampleCount = gSampleCount + 1
                If gSampleCount <= MB_BENCHMARK_COUNT Then gSampleSeconds(gSampleCount) = elapsed
            End If
        End If
    Loop
    stream.Close
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If metricInFlight Then
        elapsed = -1#
        elapsed = MB_Clock() - tick
        MBF_Record gMode, result, elapsed, (gContentReview > priorContentReview)
    End If
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise errNo, "MB_ProcessQueue", errText
End Sub

Private Function MB_Revision(ByVal mail As Outlook.MailItem, ByVal sourceKey As String) As String
    Dim changeKey As String, modified As Date, itemSize As Long, precise As String
    modified = mail.LastModificationTime
    itemSize = mail.Size
    precise = Replace(Format$(CDbl(modified), "0.0000000000"), ",", ".")
    changeKey = MB_GetChangeKey(mail)
    MB_Revision = MBH_HashText("MBREV2|" & sourceKey & "|" & precise & "|" & _
                             CStr(itemSize) & "|" & changeKey)
End Function

Private Function MB_GetChangeKey(ByVal mail As Outlook.MailItem) As String
    Dim value As Variant, i As Long, result As String
    ' Optional provider property. Absence is not a missing email.
    On Error GoTo NotProvided
    value = mail.PropertyAccessor.GetProperty(MB_CHANGE_KEY)
    If Not IsArray(value) Then GoTo NotProvided
    For i = LBound(value) To UBound(value)
        result = result & Right$("0" & Hex$(CByte(value(i))), 2)
    Next i
    MB_GetChangeKey = LCase$(result)
    Exit Function
NotProvided:
    MB_GetChangeKey = "NO_PROVIDER_CHANGEKEY"
End Function

Private Function MB_SourceStable(ByVal mail As Outlook.MailItem, ByVal q As Variant, _
                                  ByVal revision As String) As Boolean
    Dim parent As Outlook.Folder
    Set parent = mail.Parent
    If MB_SourceKey(parent.StoreID, mail.EntryID) <> CStr(q(7)) Then Exit Function
    If StrComp(parent.FolderPath, CStr(q(3)), vbBinaryCompare) <> 0 Then Exit Function
    If MB_Stamp(mail.LastModificationTime) <> CStr(q(4)) Then Exit Function
    If CDbl(mail.Size) <> MB_ByteCount(CStr(q(5))) Then Exit Function
    If Len(revision) > 0 Then
        If MB_Revision(mail, CStr(q(7))) <> revision Then Exit Function
    End If
    MB_SourceStable = True
End Function

Private Function MB_MessageLeaf(ByVal mail As Outlook.MailItem, ByVal q As Variant, _
                                ByVal revision As String, ByVal parent As String) As String
    MB_MessageLeaf = MBN_MessageLeaf(MBC_FileStamp(MBY_MessageDate(mail)), _
                                     MB_SenderLabel(mail), mail.Subject, parent)
End Function

Private Function MB_SenderLabel(ByVal mail As Outlook.MailItem) As String
    On Error GoTo Missing
    MB_SenderLabel = mail.SenderName
    If Len(MB_SenderLabel) = 0 Then MB_SenderLabel = "unknown_sender"
    Exit Function
Missing:
    MB_SenderLabel = "unknown_sender"
End Function

Private Function MB_ProcessOne(ByVal q As Variant) As Long
    Dim obj As Object, mail As Outlook.MailItem, originalFolder As Outlook.Folder
    Dim revision As String, relativePath As String, finalPath As String, stagePath As String
    Dim digest As String, bytes As Double, subject As String, stage As String
    Dim wasNew As Boolean, attCount As Long, attHandled As Long, i As Long, r As Long
    Dim attAccounted As Boolean, messageDir As String, wasLegacy As Boolean
    Dim legacyResult As Long, legacyReason As String, accountRoot As String, recordKind As String
    Dim errNo As Long, errText As String, sourceRevisionAfter As String, validation As String
    Dim contentComplete As Boolean, verifiedKnownChange As Boolean
#If VBA7 Then
    Dim publishedGuard As LongPtr
#Else
    Dim publishedGuard As Long
#End If
    On Error GoTo Failed
    stage = "GET_ITEM"
    Set obj = Application.Session.GetItemFromID(CStr(q(2)), CStr(q(1)))
    If Not TypeOf obj Is Outlook.MailItem Then
        MB_LogEvent "WARNING", "TYPE_CHANGED", q(7), q(3), 0, "Queued item is no longer MailItem."
        MB_ProcessOne = 4
        Exit Function
    End If
    Set mail = obj
    Set originalFolder = mail.Parent
    If MB_SourceKey(originalFolder.StoreID, mail.EntryID) <> CStr(q(7)) Then GoTo Changed
    If StrComp(originalFolder.FolderPath, CStr(q(3)), vbBinaryCompare) <> 0 Then
        ' Search folders are virtual views; an item belongs to its real parent.
        ' Same-store moves that preserve the source revision are also recorded.
        MB_LogEvent "INFO", "SOURCE_LOCATION_RESOLVED", q(7), originalFolder.FolderPath, 0, _
                    "Queued view/location: " & CStr(q(3)) & "; recording actual parent at export."
        q(3) = originalFolder.FolderPath
    End If
    If Not MB_SourceStable(mail, q, "") Then GoTo Changed
    stage = "DOWNLOAD_STATE"
    If mail.DownloadState = olHeaderOnly Then
        Err.Raise MB_ERR + 33, , "Header-only message. Download it fully, then rerun."
    End If
    revision = MB_Revision(mail, CStr(q(7)))
    subject = mail.Subject
    stage = "VERIFY_EXISTING_MSG"
    If MB_ExistingValid("M|" & revision, relativePath) Then
        stage = "UPGRADE_EXISTING_VALIDATION_NO_SAVEAS"
        validation = MB_UpgradeMsgValidation(mail, q, revision)
        If validation = MB_NEW_BASIC_VALIDATION Then gContentReview = gContentReview + 1
        wasNew = False
    Else
        If gReuseBlocked Then
            legacyReason = "Recorded backup needs review: " & gReuseReason
            GoTo LegacyReview
        End If
        If gMode <> "BENCHMARK" Then
            stage = "VERIFY_KNOWN_SOURCE_REVISION_NO_SAVEAS"
            legacyResult = MBV_TryKnownRevision(mail, q, revision, legacyReason)
            Select Case legacyResult
                Case MBV_REUSED
                    wasNew = False
                    GoTo MessageReady
                Case MBV_REVIEW
                    GoTo LegacyReview
                Case MBV_CANCELLED
                    MB_ProcessOne = 5
                    Exit Function
                Case MBV_CHANGED
                    verifiedKnownChange = True
                Case MBV_NONE
                    verifiedKnownChange = False
                Case Else
                    legacyReason = "Unknown known-revision decision; prior files retained."
                    GoTo LegacyReview
            End Select
            stage = "VERIFY_RESTORED_IDENTITY_NO_SAVEAS"
            legacyResult = MBI_TryAlias(mail, q, revision, legacyReason)
            If legacyResult = 1 Then
                wasNew = False
                GoTo MessageReady
            ElseIf legacyResult = 3 Then
                MB_ProcessOne = 5
                Exit Function
            ElseIf legacyResult = 2 Then
                GoTo LegacyReview
            End If
            If Not verifiedKnownChange Then
                stage = "VERIFY_LEGACY_MSG"
                legacyResult = MB_ReuseLegacy(mail, q, revision, legacyReason)
                If legacyResult = 1 Then
                    wasNew = False: wasLegacy = True
                    GoTo MessageReady
                ElseIf legacyResult = 3 Then
                    MB_ProcessOne = 5
                    Exit Function
                ElseIf legacyResult = 2 Then
                    GoTo LegacyReview
                End If
            End If
        End If
        If gMode = "REGISTER_LEGACY" Then
            MB_LogEvent "INFO", "LEGACY_UNMATCHED_NO_EXPORT", q(7), q(3), 0, _
                        "No verified existing MSG match; registration mode never saves new messages."
            MB_ProcessOne = 8
            Exit Function
        End If
        If Not MB_NewExportAllowed(mail, legacyReason, verifiedKnownChange) Then GoTo LegacyReview
        stage = "SAVE_MSG"
        If gMode = "BENCHMARK" Then
            messageDir = MBH_JoinPath(MBH_JoinPath(gRoot, "messages"), Left$(revision, 2))
            recordKind = "M"
        Else
            accountRoot = MBY_AccountRoot(CStr(q(1)))
            If Len(accountRoot) = 0 Then MB_ProcessOne = 5: Exit Function
            messageDir = MBY_MonthDirectory(accountRoot, MBY_MessageDate(mail))
            recordKind = "N"
        End If
        finalPath = MBH_JoinPath(messageDir, MB_MessageLeaf(mail, q, revision, messageDir))
        MB_EnsureDir messageDir
        MBY_NoReparse messageDir, True
        finalPath = MB_UnusedFinal(finalPath)
        If recordKind = "N" Then
            relativePath = finalPath
        Else
            relativePath = Mid$(finalPath, Len(gRoot) + 2)
        End If
        stagePath = MBH_NewStagePath(finalPath)
        mail.SaveAs stagePath, olMSGUnicode
        If CDbl(gFSO.GetFile(stagePath).Size) = 0 Then Err.Raise MB_ERR + 34, , "Empty exported MSG."
        stage = "REOPEN_MSG"
        validation = "SHA256"
        If MB_REOPEN_NEW_MSG Then
            MB_CompareMsg mail, stagePath, contentComplete
            If contentComplete Then
                validation = MB_NEW_CONTENT_VALIDATION
            Else
                validation = MB_NEW_BASIC_VALIDATION
            End If
        End If
        stage = "HASH_MSG"
        digest = MBH_HashFile(stagePath)
        bytes = CDbl(gFSO.GetFile(stagePath).Size)
        If Not MB_SourceStable(mail, q, revision) Then GoTo Changed
        stage = "COMMIT_MSG"
        MBH_CommitStage stagePath, finalPath, digest, bytes
        stagePath = ""
        publishedGuard = MBH_GuardFile(finalPath)
        If MBH_HashFile(finalPath) <> digest Or CDbl(gFSO.GetFile(finalPath).Size) <> bytes Then _
            Err.Raise MB_ERR + 923, , "Published MSG changed before index commit."
        ' The stage was content-checked against this captured revision before
        ' publication. Record those verified historical bytes even if the live
        ' item changed immediately afterward. MessageReady still refuses to
        ' count the CURRENT revision as complete in that case.
        MB_RecordCommit recordKind, revision, relativePath, bytes, digest, q, subject, 0, "", validation, accountRoot
        MBH_CloseGuard publishedGuard
        gBytesNew = gBytesNew + bytes
        If validation = MB_NEW_BASIC_VALIDATION Then
            gContentReview = gContentReview + 1
            MB_LogEvent "WARNING", "SPECIAL_ATTACHMENT_CONTENT_REVIEW", q(7), finalPath, 0, _
                        "MSG saved and reopened; special attachment or custom metadata equivalence is not fully proven."
        End If
        wasNew = True
    End If
MessageReady:
    If Not MB_SourceStable(mail, q, revision) Then GoTo Changed
    If gMode = "REGISTER_LEGACY" Then
        If wasLegacy Then gMailLegacy = gMailLegacy + 1
        MB_ProcessOne = 2
        Exit Function
    End If
    If MB_EXTRACT_ATTACHMENTS Then
        attCount = mail.Attachments.Count
        gAttFound = gAttFound + attCount
        For i = 1 To attCount
            If Not MB_ContinueWork() Then
                Exit For
            End If
            stage = "ATTACHMENT_" & CStr(i)
            r = MB_ProcessAttachment(mail, i, revision, q, subject)
            Select Case r
                Case 1: gAttNew = gAttNew + 1
                Case 2: gAttReused = gAttReused + 1
                Case 3: gAttFailed = gAttFailed + 1
                Case 4: gAttExcluded = gAttExcluded + 1
                Case Else: Err.Raise MB_ERR + 35, , "Unknown attachment result."
            End Select
            attHandled = attHandled + 1
        Next i
        gAttPending = gAttPending + (attCount - attHandled)
        attAccounted = True
    End If
    If Not MB_SourceStable(mail, q, revision) Then GoTo Changed
    If wasLegacy Then gMailLegacy = gMailLegacy + 1
    If wasNew Then
        MB_ProcessOne = 1
    Else
        MB_ProcessOne = 2
    End If
    Exit Function
LegacyReview:
    MB_LogEvent "WARNING", "LEGACY_REVIEW_REQUIRED", q(7), q(3), 0, legacyReason
    MB_ProcessOne = 7
    Exit Function
Changed:
    MBH_CloseGuard publishedGuard
    If Not attAccounted Then gAttPending = gAttPending + (attCount - attHandled)
    If Len(stagePath) > 0 Then MBH_DeleteStage stagePath
    MB_LogEvent "WARNING", "SOURCE_CHANGED_RETRY", q(7), q(3), 0, _
                "Source changed/moved after queue capture or during export; rerun to capture its current revision."
    MB_ProcessOne = 4
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    If Not attAccounted Then gAttPending = gAttPending + (attCount - attHandled)
    On Error Resume Next
    MBH_CloseGuard publishedGuard
    If Len(stagePath) > 0 Then MBH_DeleteStage stagePath
    On Error GoTo 0
    MB_LogEvent "ERROR", stage, q(7), q(3), errNo, errText
    If gFatal Then Err.Raise errNo, "MB_ProcessOne", errText
    MB_ProcessOne = 3
End Function

Private Function MB_UnusedFinal(ByVal requested As String) As String
    Dim candidate As String, parent As String, stem As String, ext As String, n As Long
    If Not gFSO.FileExists(requested) And Not gFSO.FolderExists(requested) Then
        MB_UnusedFinal = requested
        Exit Function
    End If
    parent = gFSO.GetParentFolderName(requested)
    stem = gFSO.GetBaseName(requested)
    ext = gFSO.GetExtensionName(requested)
    If Len(ext) > 0 Then ext = "." & ext
    For n = 1 To 99999
        candidate = MBH_JoinPath(parent, stem & "_r" & Format$(n, "00000") & ext)
        If Not gFSO.FileExists(candidate) And Not gFSO.FolderExists(candidate) Then
            MB_UnusedFinal = candidate
            Exit Function
        End If
    Next n
    Err.Raise MB_ERR + 36, "MB_UnusedFinal", "Too many preserved path conflicts."
End Function

Private Function MB_UpgradeMsgValidation(ByVal mail As Outlook.MailItem, ByVal q As Variant, _
                                         ByVal revision As String) As String
    Dim key As String, fields As Variant, path As String, bound As String
    Dim validation As String, complete As Boolean, relative As String, kind As String
    Dim n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    key = "M|" & revision
    fields = MB_IndexFields(key)
    validation = CStr(fields(3))
    If MB_REOPEN_NEW_MSG And (validation <> MB_NEW_CONTENT_VALIDATION Or MBW_RequireContentCheck(CStr(q(7)))) Then
        path = MBX_IndexedPath(key, fields, True)
        guard = MBH_GuardFile(path)
        MB_CompareMsg mail, path, complete
        If MBH_HashFile(path) <> CStr(fields(2)) Then Err.Raise MB_ERR + 80, , "Existing MSG changed during validation upgrade."
        If Not MB_SourceStable(mail, q, revision) Then Err.Raise MB_ERR + 81, , "Source changed during validation upgrade."
        If complete Then validation = MB_NEW_CONTENT_VALIDATION Else validation = MB_NEW_BASIC_VALIDATION
        bound = MBM_ResolvedRoot(MBX_IndexedOriginal(key, fields), CStr(fields(2)), MBX_IndexedRoot(key))
        If StrComp(bound, gRoot, vbTextCompare) = 0 Then
            kind = "M": relative = Mid$(path, Len(gRoot) + 2)
        Else
            kind = "N": relative = path
        End If
        MB_RecordCommit kind, revision, relative, MB_ByteCount(CStr(fields(1))), CStr(fields(2)), _
                        q, mail.Subject, 0, "", validation, bound
        MBH_CloseGuard guard
        MB_LogEvent "INFO", "VALIDATION_UPGRADED_NO_SAVEAS", q(7), path, 0, validation
    End If
    MB_UpgradeMsgValidation = validation
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    Err.Raise n, "MB_UpgradeMsgValidation", detail
End Function

Private Sub MB_CompareMsg(ByVal source As Outlook.MailItem, ByVal path As String, _
                          Optional ByRef contentComplete As Boolean = False)
    Dim obj As Object, copy As Outlook.MailItem, a As Outlook.Attachment, b As Outlook.Attachment
    Dim ra As Outlook.Recipient, rb As Outlook.Recipient
    Dim i As Long, errNo As Long, errText As String, digest As String
    Dim sourceId As String, copyId As String, hasSourceId As Boolean, hasCopyId As Boolean
    On Error GoTo Failed
    contentComplete = True
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 37, , "Export is not MailItem."
    Set copy = obj
    If Not MBE_MetadataComplete(source, copy) Then contentComplete = False
    If StrComp(source.MessageClass, copy.MessageClass, vbBinaryCompare) <> 0 Or _
       source.Sent <> copy.Sent Or source.BodyFormat <> copy.BodyFormat Then
        Err.Raise MB_ERR + 74, , "Message class/sent state/body format mismatch."
    End If
    If source.Sent Then
        If MB_Stamp(source.SentOn) <> MB_Stamp(copy.SentOn) Or _
           MB_Stamp(source.ReceivedTime) <> MB_Stamp(copy.ReceivedTime) Then
            Err.Raise MB_ERR + 75, , "Sent/received timestamp mismatch."
        End If
        If StrComp(source.SenderEmailType, copy.SenderEmailType, vbTextCompare) <> 0 Or _
           StrComp(source.SenderEmailAddress, copy.SenderEmailAddress, vbTextCompare) <> 0 Or _
           StrComp(source.SenderName, copy.SenderName, vbBinaryCompare) <> 0 Then
            Err.Raise MB_ERR + 76, , "Sender mismatch."
        End If
    End If
    sourceId = MBL_InternetId(source, hasSourceId): copyId = MBL_InternetId(copy, hasCopyId)
    If hasSourceId <> hasCopyId Or StrComp(sourceId, copyId, vbBinaryCompare) <> 0 Then
        Err.Raise MB_ERR + 77, , "Internet Message-ID mismatch."
    End If
    If StrComp(source.Subject, copy.Subject, vbBinaryCompare) <> 0 Then Err.Raise MB_ERR + 38, , "Subject mismatch."
    If StrComp(source.Body, copy.Body, vbBinaryCompare) <> 0 Then Err.Raise MB_ERR + 39, , "Body mismatch."
    If source.BodyFormat = olFormatHTML Then
        If StrComp(source.HTMLBody, copy.HTMLBody, vbBinaryCompare) <> 0 Then
            Err.Raise MB_ERR + 40, , "HTML body mismatch."
        End If
    ElseIf source.BodyFormat = olFormatRichText Then
        If Not MBL_ByteArraysEqual(source.RTFBody, copy.RTFBody) Then
            Err.Raise MB_ERR + 78, , "RTF body bytes mismatch."
        End If
    End If
    If source.Recipients.Count <> copy.Recipients.Count Then Err.Raise MB_ERR + 41, , "Recipient count mismatch."
    For i = 1 To source.Recipients.Count
        Set ra = source.Recipients.Item(i): Set rb = copy.Recipients.Item(i)
        If ra.Type <> rb.Type Or StrComp(ra.Address, rb.Address, vbBinaryCompare) <> 0 Then
            Err.Raise MB_ERR + 70, , "Recipient type/address mismatch at index " & CStr(i)
        End If
    Next i
    Set ra = Nothing: Set rb = Nothing
    If source.Attachments.Count <> copy.Attachments.Count Then Err.Raise MB_ERR + 42, , "Attachment count mismatch."
    For i = 1 To source.Attachments.Count
        Set a = source.Attachments.Item(i)
        Set b = copy.Attachments.Item(i)
        If a.FileName <> b.FileName Or a.Type <> b.Type Or a.Size <> b.Size Then
            Err.Raise MB_ERR + 43, , "Attachment metadata mismatch at index " & CStr(i)
        End If
        MBL_CompareAttachmentProperties a, b
        If a.Type = olByValue Then
            digest = MBL_AttachmentHash(a)
            If digest <> MBL_AttachmentHash(b) Then Err.Raise MB_ERR + 79, , "Attachment bytes mismatch at index " & CStr(i)
        Else
            contentComplete = False
        End If
    Next i
    Set a = Nothing: Set b = Nothing
    copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    Set a = Nothing: Set b = Nothing
    Set ra = Nothing: Set rb = Nothing
    If Not copy Is Nothing Then copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    On Error GoTo 0
    Err.Raise errNo, "MB_CompareMsg", errText
End Sub

Private Function MB_AttachmentAllowed(ByVal name As String) As Boolean
    Dim ext As String, dot As Long, allowed As String
    allowed = LCase$(Replace(MB_ATTACHMENT_EXTENSIONS, " ", ""))
    If allowed = "*" Then
        MB_AttachmentAllowed = True
        Exit Function
    End If
    dot = InStrRev(name, ".")
    If dot = 0 Then Exit Function
    ext = LCase$(Mid$(name, dot + 1))
    If Len(ext) = 0 Then Exit Function
    MB_AttachmentAllowed = (InStr(1, ";" & allowed & ";", ";" & ext & ";", vbBinaryCompare) > 0)
End Function

Private Function MB_AttachmentLeaf(ByVal original As String, ByVal ordinal As Long, _
                                   ByVal parent As String) As String
    Dim safe As String, dot As Long, stem As String, ext As String
    Dim prefix As String, budget As Long
    ' Split the original extension BEFORE shortening the basename.
    dot = InStrRev(original, ".")
    If dot > 1 And Len(original) - dot <= 16 Then
        stem = Left$(original, dot - 1)
        ext = MBH_SafeComponent(Mid$(original, dot + 1), 16, "bin")
        ext = "." & ext
    Else
        stem = original: ext = ""
    End If
    prefix = Format$(ordinal, "00000") & "_"
    ' Leave space for _r00001 and for staged names at the same directory depth.
    budget = MBH_MAX_PATH_UNITS - Len(parent) - 1 - Len(prefix) - Len(ext) - 8
    If budget < 8 Then Err.Raise MB_ERR + 44, , "Attachment parent path too long."
    MB_AttachmentLeaf = prefix & MBH_SafeComponent(stem, budget, "attachment") & ext
End Function

Private Function MB_ProcessAttachment(ByVal mail As Outlook.MailItem, ByVal ordinal As Long, _
                                      ByVal revision As String, ByVal q As Variant, _
                                      ByVal subject As String) As Long
    Dim attachment As Outlook.Attachment, key As String, original As String
    Dim parent As String, finalPath As String, relativePath As String, stagePath As String
    Dim digest As String, bytes As Double, errNo As Long, errText As String, stage As String
    Dim accountRoot As String, recordKind As String
#If VBA7 Then
    Dim publishedGuard As LongPtr
#Else
    Dim publishedGuard As Long
#End If
    On Error GoTo Failed
    stage = "READ_ATTACHMENT"
    Set attachment = mail.Attachments.Item(ordinal)
    original = attachment.FileName
    If Not MB_AttachmentAllowed(original) Then
        MB_LogEvent "INFO", "ATTACHMENT_POLICY_SKIP", q(7), q(3), 0, _
                    CStr(ordinal) & ": " & original
        MB_ProcessAttachment = 4
        Exit Function
    End If
    key = MBH_HashText("MBATT2|" & revision & "|" & CStr(ordinal))
    stage = "VERIFY_EXISTING_ATTACHMENT"
    If MB_ExistingValid("A|" & key, relativePath) Then
        MB_ProcessAttachment = 2
        Exit Function
    End If
    If gReuseBlocked Then Err.Raise MB_ERR + 608, , "Recorded attachment needs review: " & gReuseReason
    If gMode = "BENCHMARK" Then
        parent = MBH_JoinPath(MBH_JoinPath(MBH_JoinPath(gRoot, "attachments"), Left$(revision, 2)), revision)
        recordKind = "A"
    Else
        accountRoot = MBY_AccountRoot(CStr(q(1)))
        If Len(accountRoot) = 0 Then Err.Raise MB_ERR + 609, , "Attachment export cancelled."
        parent = MBH_JoinPath(MBH_JoinPath(MBY_MonthDirectory(accountRoot, MBY_MessageDate(mail)), "_attachments"), revision)
        recordKind = "T"
    End If
    MB_EnsureDir parent
    MBY_NoReparse parent, True
    finalPath = MBH_JoinPath(parent, MB_AttachmentLeaf(original, ordinal, parent))
    finalPath = MB_UnusedFinal(finalPath)
    If recordKind = "T" Then
        relativePath = finalPath
    Else
        relativePath = Mid$(finalPath, Len(gRoot) + 2)
    End If
    stagePath = MBH_NewStagePath(finalPath)
    stage = "SAVE_ATTACHMENT"
    attachment.SaveAsFile stagePath
    bytes = CDbl(gFSO.GetFile(stagePath).Size)
    ' Zero-byte attachments can be valid. Do not equate MAPI.Size with exported size.
    digest = MBH_HashFile(stagePath)
    If Not MB_SourceStable(mail, q, revision) Then
        Err.Raise MB_ERR + 45, , "Source changed during attachment extraction."
    End If
    stage = "COMMIT_ATTACHMENT"
    MBH_CommitStage stagePath, finalPath, digest, bytes
    stagePath = ""
    publishedGuard = MBH_GuardFile(finalPath)
    If MBH_HashFile(finalPath) <> digest Or CDbl(gFSO.GetFile(finalPath).Size) <> bytes Then _
        Err.Raise MB_ERR + 924, , "Published attachment changed before index commit."
    MB_RecordCommit recordKind, key, relativePath, bytes, digest, q, subject, ordinal, original, "SHA256", accountRoot
    MBH_CloseGuard publishedGuard
    gBytesNew = gBytesNew + bytes
    MB_ProcessAttachment = 1
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    MBH_CloseGuard publishedGuard
    If Len(stagePath) > 0 Then MBH_DeleteStage stagePath
    On Error GoTo 0
    MB_LogEvent "ERROR", stage, q(7), q(3), errNo, _
                "attachment=" & CStr(ordinal) & "; name=" & original & "; " & errText
    If gFatal Then Err.Raise errNo, "MB_ProcessAttachment", errText
    MB_ProcessAttachment = 3
End Function

' ===================== Selected mail / benchmark sample =====================
Private Function MB_QueueSelection(ByVal selection As Outlook.Selection, ByVal path As String) As Long
    Dim stream As Object, obj As Object, mail As Outlook.MailItem, folder As Outlook.Folder
    Dim seen As Object, key As String, i As Long, payload As String, dt As Date
    Dim errNo As Long, errText As String
    On Error GoTo Failed
    Set seen = CreateObject("Scripting.Dictionary")
    Set stream = gFSO.CreateTextFile(path, False, True)
    For i = 1 To selection.Count
        If Not MB_ContinueWork() Then Exit For
        Set obj = selection.Item(i)
        If TypeOf obj Is Outlook.MailItem Then
            Set mail = obj: Set folder = mail.Parent
            key = MB_SourceKey(folder.StoreID, mail.EntryID)
            If Not seen.Exists(key) Then
                seen.Add key, True
                dt = MB_EffectiveDate(mail)
                payload = MB_EncodeFields(Array("MBQ2", folder.StoreID, mail.EntryID, _
                          folder.FolderPath, MB_Stamp(mail.LastModificationTime), _
                          CStr(mail.Size), MB_Stamp(dt), key))
                MB_WriteCheckedLine stream, payload
                MB_QueueSelection = MB_QueueSelection + 1
            End If
        Else
            gPolicySkipped = gPolicySkipped + 1
        End If
    Next i
    stream.Close
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise errNo, "MB_QueueSelection", errText
End Function

Private Function MB_ValidDate(ByVal value As Date) As Boolean
    MB_ValidDate = MBC_ValidDate(value)
End Function

Private Function MB_EffectiveDate(ByVal mail As Outlook.MailItem) As Date
    Dim value As Date
    value = mail.ReceivedTime
    If MB_ValidDate(value) Then MB_EffectiveDate = value: Exit Function
    value = mail.SentOn
    If MB_ValidDate(value) Then MB_EffectiveDate = value: Exit Function
    value = mail.CreationTime
    If MB_ValidDate(value) Then MB_EffectiveDate = value: Exit Function
    Err.Raise MB_ERR + 46, , "No usable received/sent/creation timestamp."
End Function

Private Function MB_MakeSample(ByVal queuePath As String, ByVal limit As Long) As String
    Dim reservoir() As String, stream As Object, output As Object, line As String
    Dim n As Long, j As Long, count As Long, samplePath As String
    Dim errNo As Long, errText As String
    On Error GoTo Failed
    ReDim reservoir(1 To limit)
    Randomize
    Set stream = gFSO.OpenTextFile(queuePath, 1, False, -1)
    Do Until stream.AtEndOfStream
        If Not MB_ContinueWork() Then Exit Do
        line = stream.ReadLine
        If Len(line) > 0 Then
            n = n + 1
            If n <= limit Then
                reservoir(n) = line
            Else
                j = CLng(Int(Rnd * CDbl(n))) + 1
                If j <= limit Then reservoir(j) = line
            End If
        End If
    Loop
    stream.Close
    samplePath = MBH_JoinPath(gRunDir, "sample_queue.tsv")
    Set output = gFSO.CreateTextFile(samplePath, False, True)
    count = limit
    If n < count Then count = n
    For j = 1 To count
        output.WriteLine reservoir(j)
    Next j
    output.Close
    gPlanned = count
    MB_MakeSample = samplePath
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Not output Is Nothing Then output.Close
    On Error GoTo 0
    Err.Raise errNo, "MB_MakeSample", errText
End Function

' ============================= Integrity audit =============================
Private Sub MB_AuditIndex()
    Dim key As Variant, fields As Variant, p As String, ok As Boolean, binding As Variant
    gPlanned = gIndex.Count
    gFullPlanned = gPlanned
    For Each key In gIndex.Keys
        If Not MB_ContinueWork() Then Exit For
        fields = MB_IndexFields(CStr(key))
        p = MBX_IndexedPath(CStr(key), fields)
        ok = MB_AuditFile(CStr(key), p, MB_ByteCount(CStr(fields(1))), CStr(fields(2)))
        gProcessed = gProcessed + 1
        If ok Then
            gMailReused = gMailReused + 1
        Else
            gMailFailed = gMailFailed + 1
        End If
    Next key
End Sub

Private Function MB_AuditFile(ByVal key As String, ByVal path As String, _
                              ByVal bytes As Double, ByVal digest As String) As Boolean
    Dim obj As Object, errNo As Long, errText As String, binding As Variant
    On Error GoTo Failed
    path = MBX_IndexedPath(key, MB_IndexFields(key), True)
    If Not gFSO.FileExists(path) Then Err.Raise MB_ERR + 47, , "Indexed file is missing."
    If CDbl(gFSO.GetFile(path).Size) <> bytes Then Err.Raise MB_ERR + 48, , "File size mismatch."
    If MBH_HashFile(path) <> digest Then Err.Raise MB_ERR + 49, , "SHA-256 mismatch."
    If Left$(key, 2) = "M|" Or Left$(key, 2) = "L|" Then
        Set obj = Application.Session.OpenSharedItem(path)
        If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 50, , "MSG cannot reopen as MailItem."
        obj.Close olDiscard
        Set obj = Nothing
    End If
    MB_AuditFile = True
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not obj Is Nothing Then obj.Close olDiscard
    Set obj = Nothing
    On Error GoTo 0
    MB_LogEvent "ERROR", "AUDIT_FAILED", key, path, errNo, errText
End Function

' ============================= Logs and UI =================================
Private Sub MB_LogEvent(ByVal level As String, ByVal stage As String, ByVal key As String, _
                       ByVal path As String, ByVal errorNumber As Long, ByVal detail As String)
    On Error GoTo Failed
    If gEvents Is Nothing Then Err.Raise MB_ERR + 51, , "Event log is not open."
    gEvents.WriteLine MB_EncodeFields(Array(MB_Stamp(Now), level, stage, key, path, CStr(errorNumber), detail))
    Exit Sub
Failed:
    gFatal = True
    gFatalText = "Logging failed: " & Err.Description
    Err.Raise MB_ERR + 52, "MB_LogEvent", gFatalText
End Sub

Private Function MB_Clock() As Double
    Static frequency As Currency
    Dim counter As Currency
    If frequency = 0 Then
        If MB_QueryPerformanceFrequency(frequency) = 0 Then
            Err.Raise MB_ERR + 71, "MB_Clock", "Performance-counter frequency unavailable."
        End If
    End If
    If MB_QueryPerformanceCounter(counter) = 0 Then
        Err.Raise MB_ERR + 72, "MB_Clock", "Performance counter unavailable."
    End If
    MB_Clock = CDbl(counter) / CDbl(frequency)
End Function

Private Function MB_ContinueWork() As Boolean
    Dim tick As Double, ts As Object, progress As String
    If gFatal Or gCancel Then Exit Function
    tick = MB_Clock()
    If tick - gLastPulse >= 0.25 Then
        gLastPulse = tick
        DoEvents
        If gFatal Or gCancel Then Exit Function
        If Len(gRoot) > 0 Then
            If gFSO.FileExists(MBH_JoinPath(gRoot, "STOP.txt")) Then gCancel = True
            If gMode = "BENCHMARK" Then
                If gFSO.FileExists(MBH_JoinPath(gBaseRoot, "STOP.txt")) Then gCancel = True
            End If
        End If
    End If
    If tick - gLastProgress >= 2# And Len(gRunDir) > 0 Then
        gLastProgress = tick
        progress = gPhase & vbCrLf & "Processed: " & CStr(gProcessed) & "/" & CStr(gPlanned) & _
                   vbCrLf & "Folders: " & CStr(gFoldersScanned) & _
                   vbCrLf & "Elapsed seconds: " & Format$(tick - gClockStart, "0.0") & _
                   vbCrLf & "Run: " & gRunDir & vbCrLf & _
                   "Cancel: run MB_01_RunBackup again or create STOP.txt in the output root."
        Debug.Print Replace(progress, vbCrLf, " | ")
        ' Progress is optional; an unavailable progress file must not hide backup errors.
        On Error Resume Next
        Set ts = gFSO.CreateTextFile(MB_StatePath("progress.txt"), True, True)
        If Not ts Is Nothing Then ts.Write progress: ts.Close
        Set ts = Nothing
        Err.Clear
        On Error GoTo 0
    End If
    MB_ContinueWork = Not (gCancel Or gFatal)
End Function

Private Function MB_Status() As String
    If gFatal Then
        MB_Status = "FAILED"
    ElseIf gCancel Then
        MB_Status = "CANCELLED"
    ElseIf gMailFailed > 0 Or gMailChanged > 0 Or gMailPending > 0 Or gMailReview > 0 Or gContentReview > 0 Or _
           gAttFailed > 0 Or gAttPending > 0 Or gFolderFailures > 0 Or _
           gScanChanged > 0 Or gJournalIssues > 0 Or gAutoReviewCount > 0 Or gAutoFailureCount > 0 Or _
           MBA_HasReview() Or MBW_HasAuditReview() Or gStageReview > 0 Or gStageHeld > 0 Or MBR_HasReview() Then
        MB_Status = "PARTIAL"
    ElseIf gMode = "REGISTER_LEGACY" Then
        MB_Status = "LEGACY REGISTRATION FINISHED (NOT A FULL BACKUP)"
    ElseIf gMode = "BENCHMARK" Then
        MB_Status = "BENCHMARK SAMPLE COMPLETE"
    Else
        MB_Status = "COMPLETE FOR DECLARED SCOPE"
    End If
End Function

Private Function MB_Summary() As String
    Dim s As String, label As String, total As Double, mean As Double, rangeText As String
    Dim a() As Double, i As Long, j As Long, x As Double, idx95 As Long
    If MBX_IsFileMode(gMode) Then MB_Summary = MBX_FileSummary(): Exit Function
    label = "Messages"
    If gMode = "VERIFY" Then label = "Indexed files"
    If gSince = 0 Then
        rangeText = "ALL gathered items (collection starts " & MB_Stamp(gUntil) & ")"
    Else
        rangeText = MB_Stamp(gSince) & " through " & MB_Stamp(gUntil)
        If gModifiedWindow Then rangeText = "LastModificationTime: " & rangeText
    End If
    s = MB_Status() & vbCrLf & "Version: " & MB_VERSION & "; mode: " & gMode & vbCrLf & _
        "Current organization review records: " & CStr(gAutoReviewCount) & " (re-evaluated this run)" & vbCrLf & _
        "Scope: " & gScope & vbCrLf & _
        "Selected folder: " & gSelectedBase & vbCrLf & _
        "State/log root: " & gStateRoot & vbCrLf & _
        "Output layout: remembered account folder\yyyy\mm (SentOn-first)" & vbCrLf & _
        "Old archive: " & gLegacyRoot & vbCrLf & _
        "Range: " & rangeText & vbCrLf & _
        label & " planned: " & CStr(gPlanned) & "; processed: " & CStr(gProcessed) & vbCrLf & _
        "New: " & CStr(gMailNew) & "; verified existing: " & CStr(gMailReused) & _
        "; failed: " & CStr(gMailFailed) & vbCrLf & _
        "Old MSG reused in place (subset of existing): " & CStr(gMailLegacy) & vbCrLf & _
        "Review required (no SaveAs): " & CStr(gMailReview) & _
        "; registration unmatched (no SaveAs): " & CStr(gLegacyUnmatched) & vbCrLf & _
        "Source changed: " & CStr(gMailChanged) & "; policy skip: " & CStr(gMailPolicy) & _
        "; pending: " & CStr(gMailPending) & vbCrLf & _
        "Attachments: new " & CStr(gAttNew) & "; existing " & CStr(gAttReused) & _
        "; excluded " & CStr(gAttExcluded) & "; failed " & CStr(gAttFailed) & _
        "; pending " & CStr(gAttPending) & vbCrLf
    s = s & "MSGs with limited attachment/custom-metadata validation: " & CStr(gContentReview) & vbCrLf
    s = s & "Separate extraction: " & CStr(MB_EXTRACT_ATTACHMENTS) & vbCrLf & _
        "Existing SHA-256 recheck: " & CStr(MB_VERIFY_EXISTING_HASH) & _
        "; reopen new MSG: " & CStr(MB_REOPEN_NEW_MSG) & vbCrLf & _
        "Scanned folders: " & CStr(gFoldersScanned) & "; candidate mail: " & CStr(gCandidates) & _
        "; scan policy exclusions: " & CStr(gPolicySkipped) & vbCrLf & _
        "Folders failed: " & CStr(gFolderFailures) & "; scan issues: " & CStr(gScanChanged) & _
        "; invalid index rows: " & CStr(gJournalIssues) & vbCrLf & _
        "Index / scan / export seconds: " & Format$(gIndexSeconds, "0.0") & " / " & _
        Format$(gScanSeconds, "0.0") & " / " & Format$(gExportSeconds, "0.0") & vbCrLf & _
        "New output MiB: " & Format$(gBytesNew / 1048576#, "0.0") & vbCrLf & _
        "Old MSG files found: " & CStr(gLegacyFileCount) & "; unresolved filename/SentOn lookup: " & _
        CStr(gLegacyUnrecognizedNames) & "; scan seconds: " & Format$(gLegacySeconds, "0.0") & vbCrLf & _
        "Old archive newest file write (conservative guard, not coverage proof): " & _
        MB_Stamp(gLegacyNewestFileTime)
    If gMode = "BENCHMARK" And gSampleCount > 0 Then
        ReDim a(1 To gSampleCount)
        For i = 1 To gSampleCount
            a(i) = gSampleSeconds(i): total = total + a(i)
        Next i
        For i = 2 To gSampleCount
            x = a(i): j = i - 1
            Do While j >= 1
                If a(j) <= x Then Exit Do
                a(j + 1) = a(j): j = j - 1
            Loop
            a(j + 1) = x
        Next i
        mean = total / CDbl(gSampleCount)
        idx95 = CLng(Int(CDbl(gSampleCount) * 0.95))
        If idx95 < 1 Then idx95 = 1
        s = s & vbCrLf & "Sample NEW successes: " & CStr(gSampleCount) & _
            " from " & CStr(gFullPlanned) & " eligible messages." & vbCrLf & _
            "Seconds/new: mean=" & Format$(mean, "0.000") & _
            "; p50=" & Format$(a((gSampleCount + 1) \ 2), "0.000") & _
            "; p95~=" & Format$(a(idx95), "0.000") & vbCrLf & _
            "100,000 NEW estimate, hours: " & _
            Format$((gIndexSeconds + gScanSeconds + mean * 100000#) / 3600#, "0.00") & vbCrLf & _
            "Planning band (0.5x-2x variable cost), hours: " & _
            Format$((gIndexSeconds + gScanSeconds + mean * 50000#) / 3600#, "0.00") & " - " & _
            Format$((gIndexSeconds + gScanSeconds + mean * 200000#) / 3600#, "0.00") & vbCrLf & _
            "Planning band is not a statistical confidence interval. Check sample failures."
    End If
    s = s & vbCrLf & "Lexical path cache hits/misses (max 512 entries): " & _
        Format$(MBH_PathCacheHits, "0") & "/" & Format$(MBH_PathCacheMisses, "0")
    s = s & vbCrLf & "Content-verified restored/moved identity aliases (no SaveAs): " & CStr(gIdentityAliases)
    s = s & vbCrLf & "Verified identity-header reads/cache hits: " & CStr(gIdentityHeaderReads) & "/" & CStr(gIdentityHeaderHits)
    s = s & MBV_Summary() & MBF_Summary(gMode)
    s = s & MBW_Summary()
    s = s & MBA_Summary() & MBR_Summary() & MBS_Summary()
    If Len(gAutoSetupPhase) > 0 Then
        s = s & vbCrLf & "Automatic setup phase: " & gAutoSetupPhase & vbCrLf & _
            "Initial setup note: " & gAutoSetupNote & vbCrLf & _
            "Initial organization/resume seconds: " & Format$(gAutoSetupSeconds, "0.0")
        If Len(gAutoRegistrationSummary) > 0 Then s = s & vbCrLf & gAutoRegistrationSummary
    End If
    If Len(gAutoTestReport) > 0 Then s = s & vbCrLf & "Native environment test report: " & gAutoTestReport
    If gFatal Then s = s & vbCrLf & "FATAL: " & gFatalText
    If Len(gRunDir) > 0 Then s = s & vbCrLf & "Logs: " & gRunDir
    MB_Summary = s
End Function

Private Sub MB_WriteSummary(ByVal summary As String)
    Dim ts As Object
    Set ts = gFSO.CreateTextFile(MBH_JoinPath(gRunDir, "summary.txt"), True, True)
    ts.Write summary & vbCrLf
    ts.Close
End Sub

Private Sub MB_CloseResources()
    On Error Resume Next
    Err.Clear
    If Not gJournal Is Nothing Then gJournal.Close
    If Err.Number <> 0 Then gCleanupIssue = gCleanupIssue & " Journal close: " & Err.Description
    Set gJournal = Nothing
    Err.Clear
    If Not gEvents Is Nothing Then gEvents.Close
    If Err.Number <> 0 Then gCleanupIssue = gCleanupIssue & " Event log close: " & Err.Description
    Set gEvents = Nothing
    On Error GoTo 0
End Sub

' =========================== On-PC self tests ================================
Private Sub MB_SelfTest()
    Dim detail As String, ok As Boolean
    If gBusy Then Exit Sub
    gBusy = True
    ok = MBT_HelperSelfTest(detail)
    gBusy = False
    MsgBox detail, IIf(ok, vbInformation, vbExclamation), "Mail Backup helper test"
End Sub

Private Function MBT_HelperSelfTest(ByRef detail As String) As Boolean
    Dim s As String, fields As Variant, payload As String, line As String
    Dim a As String, b As String, message As String, fso As Object, testDir As String
    Dim finalPath As String, stagePath As String, ts As Object, rejected As Boolean
    Dim errNo As Long, errText As String, ownsTestDir As Boolean, clockStart As Double
    On Error GoTo Failed
    clockStart = MB_Clock()
    If Not MBH_HashSelfTest() Then Err.Raise MB_ERR + 60, , "SHA-256 known-answer test failed."
    If Not MBN_SelfTest() Then Err.Raise MB_ERR + 71, , "Searchable filename contract test failed."
    If MBH_SafeComponent("CON.txt") = "CON.txt" Then Err.Raise MB_ERR + 61, , "Reserved-name test failed."
    If MBH_SafeComponent("A:B") <> MBH_SafeComponent("A?B") Then Err.Raise MB_ERR + 62, , "Sanitizer test failed."
    a = MB_SourceKey("store", "mailA")
    b = MB_SourceKey("store", "mailB")
    If a = b Then Err.Raise MB_ERR + 63, , "Source-key separation failed."
    s = "=x" & vbTab & "%09" & vbCrLf & ChrW(&HD55C) & ChrW(&HD83D) & ChrW(&HDE00)
    payload = MB_EncodeFields(Array("TEST", s))
    line = payload & vbTab & MBH_HashText(payload) & vbTab & "END"
    fields = MB_ReadCheckedLine(line, 2)
    If fields(1) <> s Then Err.Raise MB_ERR + 64, , "Unicode TSV round-trip failed."
    rejected = MB_TestRejectRecord(Left$(line, Len(line) - 1))
    If Not rejected Then Err.Raise MB_ERR + 65, , "Truncated record was accepted."
    If DateAdd("m", -1, DateSerial(2024, 3, 31)) <> DateSerial(2024, 2, 29) Then
        Err.Raise MB_ERR + 66, , "Calendar-month boundary test failed."
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    testDir = MBH_JoinPath(fso.GetSpecialFolder(2).Path, "MBTest_" & fso.GetTempName)
    ' Claim a NEW directory; never clean up an existing directory on collision.
    fso.CreateFolder testDir
    ownsTestDir = True
    finalPath = MBH_JoinPath(testDir, "test.txt")
    stagePath = MBH_NewStagePath(finalPath)
    Set ts = fso.OpenTextFile(stagePath, 2, False, 0)
    ts.Write "first": ts.Close
    MBH_CommitStage stagePath, finalPath
    stagePath = MBH_NewStagePath(finalPath)
    Set ts = fso.OpenTextFile(stagePath, 2, False, 0)
    ts.Write "second": ts.Close
    On Error Resume Next
    Err.Clear
    MBH_CommitStage stagePath, finalPath
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    If Not rejected Then Err.Raise MB_ERR + 67, , "Existing final file was overwritten."
    Set ts = fso.OpenTextFile(finalPath, 1, False, 0)
    s = ts.ReadAll: ts.Close
    If s <> "first" Then Err.Raise MB_ERR + 68, , "Existing bytes were changed."
    MBH_DeleteStage stagePath: stagePath = ""
    fso.DeleteFolder testDir, True: testDir = "": ownsTestDir = False
    If MB_Clock() < clockStart Then Err.Raise MB_ERR + 73, , "Performance counter moved backwards."
    MBH_ReleaseCrypto
    detail = "PASS: hashes, Unicode, names, keys, journal, calendar month, no-overwrite."
    MBT_HelperSelfTest = True
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    If Len(stagePath) > 0 Then MBH_DeleteStage stagePath
    If ownsTestDir Then fso.DeleteFolder testDir, True
    MBH_ReleaseCrypto
    On Error GoTo 0
    detail = "FAIL: " & CStr(errNo) & " - " & errText
End Function

Private Function MB_TestRejectRecord(ByVal line As String) As Boolean
    Dim value As Variant
    On Error GoTo Expected
    value = MB_ReadCheckedLine(line, 2)
    Exit Function
Expected:
    MB_TestRejectRecord = True
End Function


' ================= Persistent filesystem folder selection ==================
' No hardcoded personal directory is treated as evidence of a saved selection.
' A remembered, available base is reused without showing a folder dialog.
' New bases use a MailBackup child; exact old SafeV2 paths are retained in place.
' Resolve does not create the managed root or save settings. The caller must
' acquire its run lock and validate ownership before persisting the choice.
' Settings API: HKCU\Software\VB and VBA Program Settings\OutlookMailBackup.

Private Sub MBP_Reset()
    mbpLocatorPath = "": mbpLocatorHash = "": mbpLocatorStore = ""
    mbpLocatorManaged = "": mbpLocatorAccount = ""
    mbpLocatorObserved = False
End Sub

Private Sub MBP_NoReparse(ByVal path As String, ByVal fso As Object)
    ' Folder resolution precedes gRoot initialization; use the supplied FSO.
    Dim current As String, parent As String, entry As Object
    current = MBH_NormalizePath(path)
    Do
        Set entry = Nothing
        If fso.FileExists(current) Then
            Set entry = fso.GetFile(current)
        ElseIf fso.FolderExists(current) Then
            Set entry = fso.GetFolder(current)
        End If
        If Not entry Is Nothing Then
            If (CLng(entry.Attributes) And 1024) <> 0 Then _
                Err.Raise MB_ERR + 215, , "Saved path traverses a link/reparse point: " & current
        End If
        If MBL_IsFileSystemRoot(current) Then Exit Do
        parent = CStr(fso.GetParentFolderName(current))
        If Len(parent) = 0 Or StrComp(parent, current, vbTextCompare) = 0 Then Exit Do
        current = parent
    Loop
End Sub

Private Function MBP_ReadOneLine(ByVal path As String, ByVal fso As Object) As String
    Dim stream As Object, value As String, n As Long, detail As String
    On Error GoTo Failed
    MBP_NoReparse path, fso
    If CDbl(fso.GetFile(path).Size) Mod 2 <> 0 Then _
        Err.Raise MB_ERR + 216, , "Saved binding has an incomplete UTF-16 byte: " & path
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    If stream.AtEndOfStream Then Err.Raise MB_ERR + 216, , "Saved binding is empty: " & path
    value = stream.ReadLine
    If Not stream.AtEndOfStream Then Err.Raise MB_ERR + 216, , "Saved binding has extra data: " & path
    stream.Close: Set stream = Nothing
    MBP_ReadOneLine = value
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBP_ReadOneLine", detail
End Function

Private Function MBP_LocatorFile(ByVal storeId As String) As String
    Dim base As String
    If Len(storeId) = 0 Then Err.Raise MB_ERR + 217, , "A StoreID is required for the backup locator."
    base = MBH_NormalizePath(Environ$("LOCALAPPDATA"))
    base = MBH_JoinPath(MBH_JoinPath(base, MBP_SETTINGS_APP), "locations")
    MBP_LocatorFile = MBH_JoinPath(base, "base_" & MBH_HashText(LCase$(storeId)) & ".tsv")
End Function

Private Function MBP_ReadLocator(ByVal storeId As String, ByRef base As String, _
                                 ByRef managed As String, ByRef issue As String) As Boolean
    Dim fso As Object, row As Variant, digest As String
    On Error GoTo Unusable
    base = "": managed = "": issue = ""
    mbpLocatorStore = storeId
    mbpLocatorPath = MBP_LocatorFile(storeId)
    Set fso = CreateObject("Scripting.FileSystemObject")
    MBP_NoReparse mbpLocatorPath, fso
    mbpLocatorHash = "": mbpLocatorObserved = True
    If Not fso.FileExists(mbpLocatorPath) Then Exit Function
    digest = MBH_HashFile(mbpLocatorPath)
    mbpLocatorHash = digest
    row = MB_ReadCheckedLine(MBP_ReadOneLine(mbpLocatorPath, fso), 5)
    If CStr(row(0)) <> MBP_LOCATOR_SCHEMA Or StrComp(CStr(row(1)), storeId, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 218, , "Backup locator schema or StoreID does not match."
    base = MBH_NormalizePath(CStr(row(2)))
    managed = MBH_NormalizePath(CStr(row(3)))
    If StrComp(base, managed, vbTextCompare) <> 0 Then
        If StrComp(MBH_JoinPath(base, MBP_MANAGED_LEAF), managed, vbTextCompare) <> 0 And _
           StrComp(MBH_JoinPath(base, MBP_MANAGED_LEAF_LEGACY), managed, vbTextCompare) <> 0 Then _
            Err.Raise MB_ERR + 218, , "Backup locator does not bind its base to the managed root."
    End If
    If MBH_HashFile(mbpLocatorPath) <> digest Then _
        Err.Raise MB_ERR + 218, , "Backup locator changed while being read."
    mbpLocatorManaged = managed
    mbpLocatorAccount = ""
    If Len(CStr(row(4))) > 0 Then mbpLocatorAccount = MBH_NormalizePath(CStr(row(4)))
    MBP_ReadLocator = True
    Exit Function
Unusable:
    issue = "The independent backup locator needs review; its existing bytes are retained. " & Err.Description
    base = "": managed = ""
    mbpLocatorManaged = "": mbpLocatorAccount = ""
End Function

Private Sub MBP_PersistLocator(ByVal storeId As String, ByVal selectedBase As String, ByVal managed As String, _
                               Optional ByVal account As String = "")
    Dim fso As Object, folder As String, lockPath As String, stage As String, preserved As String
    Dim contents As String, currentHash As String, nextHash As String, stream As Object
    Dim n As Long, detail As String, flags As Long, releaseError As Long
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    On Error GoTo Failed
    If Not mbpLocatorObserved Or StrComp(mbpLocatorStore, storeId, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 219, , "The backup locator was not observed before selecting this root."
    If StrComp(mbpLocatorPath, MBP_LocatorFile(storeId), vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 219, , "The per-user locator path changed during selection."
    Set fso = CreateObject("Scripting.FileSystemObject")
    MBP_NoReparse mbpLocatorPath, fso
    folder = CStr(fso.GetParentFolderName(mbpLocatorPath))
    MBH_EnsureDirectory folder
    lockPath = MBH_JoinPath(folder, "locations.lock")
    MBP_NoReparse lockPath, fso
    handle = MBH_CreateFileW(StrPtr(lockPath), MBH_GENERIC_READ Or MBH_GENERIC_WRITE, 0, 0, _
                             MBH_OPEN_ALWAYS, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If handle = -1 Then Err.Raise MB_ERR + 220, , "Backup locations are busy. Win32=" & CStr(Err.LastDllError)
    If fso.FileExists(mbpLocatorPath) Then currentHash = MBH_HashFile(mbpLocatorPath)
    If currentHash <> mbpLocatorHash Then _
        Err.Raise MB_ERR + 221, , "The saved backup location changed in another run. Re-run to read the new binding."
    If Len(account) = 0 And StrComp(mbpLocatorManaged, managed, vbTextCompare) = 0 Then account = mbpLocatorAccount
    If Len(account) > 0 Then account = MBH_NormalizePath(account)
    contents = MBS_Checked(Array(MBP_LOCATOR_SCHEMA, LCase$(storeId), _
                           MBH_NormalizePath(selectedBase), MBH_NormalizePath(managed), account))
    stage = MBH_JoinPath(folder, "_mb_locator_" & fso.GetTempName)
    Set stream = fso.CreateTextFile(stage, False, True)
    stream.Write contents: stream.Close: Set stream = Nothing
    MBM_FlushPath stage
    nextHash = MBH_HashFile(stage)
    If currentHash = nextHash Then
        fso.DeleteFile stage, False: stage = ""
    Else
        If Len(currentHash) > 0 Then
            preserved = MBH_JoinPath(folder, "preserved")
            MBP_NoReparse preserved, fso
            MBH_EnsureDirectory preserved
            preserved = MBH_JoinPath(preserved, currentHash & ".tsv")
            MBP_NoReparse preserved, fso
            If Not fso.FileExists(preserved) Then fso.CopyFile mbpLocatorPath, preserved, False
            MBM_FlushPath preserved
            If MBH_HashFile(preserved) <> currentHash Or MBH_HashFile(mbpLocatorPath) <> currentHash Then _
                Err.Raise MB_ERR + 222, , "The prior backup locator could not be preserved exactly."
        End If
        flags = 8
        If Len(currentHash) > 0 Then flags = flags Or 1
        If MBM_MoveFileExW(StrPtr(stage), StrPtr(mbpLocatorPath), flags) = 0 Then _
            Err.Raise MB_ERR + 223, , "Cannot publish backup locator. Win32=" & CStr(Err.LastDllError)
        stage = ""
        MBM_FlushPath mbpLocatorPath
        If MBH_HashFile(mbpLocatorPath) <> nextHash Then _
            Err.Raise MB_ERR + 224, , "Published backup locator hash differs."
    End If
    mbpLocatorHash = nextHash
    mbpLocatorManaged = MBH_NormalizePath(managed): mbpLocatorAccount = account
    If MBH_CloseHandle(handle) = 0 Then Err.Raise MB_ERR + 225, , "Cannot release backup-locator lock."
    handle = 0
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If handle <> 0 And handle <> -1 Then
        If MBH_CloseHandle(handle) = 0 Then releaseError = Err.LastDllError
    End If
    ' A failed stage and every differing/corrupt old locator are retained.
    On Error GoTo 0
    If releaseError <> 0 Then detail = detail & " Locator lock release failed: " & CStr(releaseError)
    Err.Raise n, "MBP_PersistLocator", detail
End Sub

Private Function MBP_LocatedAccount(ByVal storeId As String, ByVal managed As String) As String
    If StrComp(mbpLocatorStore, storeId, vbTextCompare) <> 0 Then Exit Function
    If StrComp(mbpLocatorManaged, managed, vbTextCompare) <> 0 Then Exit Function
    MBP_LocatedAccount = mbpLocatorAccount
End Function

Private Function MBP_OwnerAccountHint(ByVal storeId As String, ByVal managed As String) As String
    Dim candidate As String
    On Error GoTo Unknown
    candidate = MBP_LocatedAccount(storeId, managed)
    If Len(candidate) = 0 Then candidate = MBP_GetSettingText(MBS_AccountSettingKey(storeId, managed), "")
    If Len(candidate) > 0 Then MBP_OwnerAccountHint = MBH_NormalizePath(candidate)
Unknown:
End Function

Private Sub MBP_RememberLocatedAccount(ByVal storeId As String, ByVal account As String)
    ' The locator is a recovery hint, not authority to override an account owner.
    MBY_CheckMarker account, storeId, True
    MBP_PersistLocator storeId, gSelectedBase, gRoot, account
End Sub

Private Function MBP_GetSettingText(ByVal key As String, _
                                    Optional ByVal fallback As String = "") As String
    On Error GoTo Failed
    MBP_GetSettingText = GetSetting(MBP_SETTINGS_APP, MBP_SETTINGS_SECTION, key, fallback)
    Exit Function
Failed:
    Err.Raise MB_ERR + 201, "MBP_GetSettingText", _
              "Cannot read the remembered backup setting: " & Err.Description
End Function

Private Sub MBP_SaveSettingText(ByVal key As String, ByVal value As String)
    Dim saved As String
    On Error GoTo Failed
    saved = GetSetting(MBP_SETTINGS_APP, MBP_SETTINGS_SECTION, key, "")
    If StrComp(saved, value, vbBinaryCompare) = 0 Then Exit Sub
    SaveSetting MBP_SETTINGS_APP, MBP_SETTINGS_SECTION, key, value
    saved = GetSetting(MBP_SETTINGS_APP, MBP_SETTINGS_SECTION, key, "")
    If StrComp(saved, value, vbBinaryCompare) <> 0 Then
        Err.Raise MB_ERR + 202, , "The setting did not read back correctly."
    End If
    Exit Sub
Failed:
    Err.Raise MB_ERR + 203, "MBP_SaveSettingText", _
              "Cannot remember the selected backup folder: " & Err.Description
End Sub

Private Function MBP_ResolveBackupRoot(ByRef selectedBase As String) As String
    Dim candidate As String, managed As String, reason As String, savedBase As String
    Dim locatedBase As String, locatedRoot As String, locatorIssue As String, hasLocator As Boolean
    Dim savedRoot As String, savedReason As String, storeId As String, conflict As Boolean, savedOK As Boolean
    MBP_Reset
    storeId = Application.Session.DefaultStore.StoreID
    savedBase = MBP_GetSettingText(MBP_BASE_KEY, "")
    hasLocator = MBP_ReadLocator(storeId, locatedBase, locatedRoot, locatorIssue)
    selectedBase = ""
    If hasLocator Then
        If MBP_TryBackupBase(locatedBase, managed, reason, locatedRoot) Then
            If StrComp(managed, locatedRoot, vbTextCompare) <> 0 Then
                reason = "The selected base no longer resolves to its saved managed root."
                managed = ""
            End If
        End If
    End If
    If Len(savedBase) > 0 Then
        If hasLocator And (StrComp(savedBase, locatedBase, vbTextCompare) = 0 Or _
                           StrComp(savedBase, locatedRoot, vbTextCompare) = 0) Then
            savedOK = MBP_TryBackupBase(savedBase, savedRoot, savedReason, locatedRoot)
        Else
            savedOK = MBP_TryBackupBase(savedBase, savedRoot, savedReason)
        End If
        If savedOK Then
            If hasLocator Then
                conflict = (StrComp(savedRoot, locatedRoot, vbTextCompare) <> 0)
            End If
        ElseIf hasLocator Then
            ' An offline remembered archive is not evidence that it may be forgotten.
            ' A positively foreign-store root is irrelevant to this per-store locator.
            If InStr(1, savedReason, "[FOREIGN_STORE]", vbBinaryCompare) = 0 Then
                conflict = (StrComp(savedBase, locatedBase, vbTextCompare) <> 0 And _
                            StrComp(savedBase, locatedRoot, vbTextCompare) <> 0)
            End If
        End If
    End If
    If Not conflict Then
        If Len(managed) > 0 Then
            selectedBase = locatedBase: MBP_ResolveBackupRoot = managed
            Exit Function
        ElseIf Len(savedRoot) > 0 And Not hasLocator Then
            selectedBase = MBH_NormalizePath(savedBase): MBP_ResolveBackupRoot = savedRoot
            Exit Function
        End If
    End If
    If conflict Then
        MsgBox "Two different backup locations are remembered. Choose the archive to use." & vbCrLf & _
               savedBase & vbCrLf & locatedBase & vbCrLf & _
               "The prior locator will be preserved before the selected binding is saved.", vbInformation, "Choose backup folder"
    ElseIf Len(savedBase) > 0 Or hasLocator Or Len(locatorIssue) > 0 Then
        MsgBox "The remembered backup folder is unavailable or needs review." & vbCrLf & _
               savedBase & vbCrLf & locatedBase & vbCrLf & reason & vbCrLf & savedReason & vbCrLf & locatorIssue & _
               vbCrLf & "Choose the existing backup folder. The successful choice will be remembered.", _
               vbInformation, "Choose backup folder"
    End If
    ' No binding remains: ask once, even when a historical hardcoded default exists.
    MBP_ResolveBackupRoot = MBP_ChooseBackupRoot(selectedBase)
End Function

Private Function MBP_ChooseBackupRoot(ByRef selectedBase As String) As String
    Dim candidate As String, managed As String, reason As String, locatedBase As String, locatedRoot As String
    Dim rememberedRoot As String, fso As Object
    If Not mbpLocatorObserved Then _
        MBP_ReadLocator Application.Session.DefaultStore.StoreID, locatedBase, locatedRoot, reason
    selectedBase = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    Do
        candidate = MBP_BrowseFileSystemFolder( _
                    "Choose an existing backup folder. Existing MSG files stay in place.")
        If Len(candidate) = 0 Then Exit Function
        rememberedRoot = ""
        If Len(mbpLocatorManaged) > 0 Then
            If StrComp(MBH_NormalizePath(candidate), CStr(fso.GetParentFolderName(mbpLocatorManaged)), vbTextCompare) = 0 Then _
                rememberedRoot = mbpLocatorManaged
        End If
        If MBP_TryBackupBase(candidate, managed, reason, rememberedRoot) Then
            selectedBase = MBH_NormalizePath(candidate)
            MBP_ChooseBackupRoot = managed
            Exit Function
        End If
        MsgBox reason & vbCrLf & vbCrLf & "Choose another folder, or Cancel.", _
               vbExclamation, "Cannot use backup folder"
    Loop
End Function

Private Function MBP_BrowseFileSystemFolder(ByVal title As String) As String
    Dim shell As Object, picked As Object, entry As Object
    Set shell = CreateObject("Shell.Application")
    ' RETURNONLYFSDIRS + EDITBOX + NEWDIALOGSTYLE. Desktop root allows browsing
    ' to another drive; supplying a filesystem root would restrict navigation.
    Set picked = shell.BrowseForFolder(0, title, &H1 Or &H10 Or &H40)
    If picked Is Nothing Then Exit Function
    Set entry = picked.Self
    If Not entry.IsFileSystem Then
        Err.Raise MB_ERR + 204, "MBP_BrowseFileSystemFolder", _
                  "Select an ordinary filesystem folder, not a virtual Shell location."
    End If
    MBP_BrowseFileSystemFolder = CStr(entry.Path)
End Function

Private Function MBP_ExactOwner(ByVal rootPath As String, ByVal fso As Object) As Boolean
    Dim marker As String, stream As Object, value As String, n As Long, detail As String
    On Error GoTo Failed
    marker = MBH_JoinPath(rootPath, "owner.txt")
    If Not fso.FileExists(marker) Then
        MBP_ExactOwner = MBS_CanRecoverOwner(rootPath, fso)
        Exit Function
    End If
    MBP_NoReparse marker, fso
    If CDbl(fso.GetFile(marker).Size) Mod 2 <> 0 Then _
        Err.Raise MB_ERR + 205, , "The existing owner marker has an incomplete UTF-16 byte: " & marker
    Set stream = fso.OpenTextFile(marker, 1, False, -1)
    value = stream.ReadAll
    stream.Close
    Set stream = Nothing
    If value <> MB_OWNER & vbCrLf And value <> MB_OWNER_V34 & vbCrLf And _
       value <> MB_OWNER_PREVIOUS & vbCrLf And value <> MB_OWNER_LEGACY & vbCrLf Then
        Err.Raise MB_ERR + 205, "MBP_ExactOwner", _
                  "An invalid owner marker already exists: " & marker
    End If
    MBP_ExactOwner = True
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBP_ExactOwner", detail
End Function

Private Sub MBP_CheckUnownedEmpty(ByVal rootPath As String, ByVal fso As Object)
    Dim entry As Object, folder As Object
    Set folder = fso.GetFolder(rootPath)
    For Each entry In folder.Files
        If LCase$(CStr(entry.Name)) <> "backup.lock" Then
            Err.Raise MB_ERR + 206, "MBP_CheckUnownedEmpty", _
                      "The managed backup folder already contains unowned files: " & rootPath
        End If
    Next entry
    For Each entry In folder.SubFolders
        Err.Raise MB_ERR + 207, "MBP_CheckUnownedEmpty", _
                  "The managed backup folder already contains unowned directories: " & rootPath
    Next entry
End Sub

Private Function MBP_TryBackupBase(ByVal candidate As String, ByRef managed As String, _
                                   ByRef reason As String, Optional ByVal rememberedRoot As String = "") As Boolean
    Dim fso As Object, base As String, probeParent As String, legacyManaged As String
    Dim hasModern As Boolean, hasLegacy As Boolean
    On Error GoTo Invalid
    managed = "": reason = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = MBH_NormalizePath(candidate)
    MBP_NoReparse base, fso
    If Not fso.FolderExists(base) Then
        Err.Raise MB_ERR + 208, , "The folder does not exist or its drive/share is unavailable."
    End If
    If MBP_ExactOwner(base, fso) Then
        managed = base
        If Len(rememberedRoot) > 0 Then
            If StrComp(MBH_NormalizePath(rememberedRoot), managed, vbTextCompare) <> 0 Then _
                Err.Raise MB_ERR + 226, , "The selected base differs from its remembered managed root."
        End If
    Else
        managed = MBH_JoinPath(base, MBP_MANAGED_LEAF)
        legacyManaged = MBH_JoinPath(base, MBP_MANAGED_LEAF_LEGACY)
        hasModern = fso.FolderExists(managed) Or fso.FileExists(managed)
        hasLegacy = fso.FolderExists(legacyManaged) Or fso.FileExists(legacyManaged)
        If hasModern And hasLegacy Then Err.Raise MB_ERR + 227, , _
            "Both MailBackup and SafeV2 exist. Select the intended owned managed folder itself; neither was changed."
        If Len(rememberedRoot) > 0 Then
            rememberedRoot = MBH_NormalizePath(rememberedRoot)
            If StrComp(rememberedRoot, managed, vbTextCompare) <> 0 And _
               StrComp(rememberedRoot, legacyManaged, vbTextCompare) <> 0 Then _
                Err.Raise MB_ERR + 226, , "The remembered managed root is outside this base."
            ' A deleted remembered SafeV2 must be recovered at its same path.
            ' Do not create MailBackup merely because the old directory is gone.
            If hasModern And StrComp(rememberedRoot, legacyManaged, vbTextCompare) = 0 Then _
                Err.Raise MB_ERR + 227, , "A different managed folder exists beside the missing remembered root."
            If hasLegacy And StrComp(rememberedRoot, managed, vbTextCompare) = 0 Then _
                Err.Raise MB_ERR + 227, , "A different managed folder exists beside the missing remembered root."
            managed = rememberedRoot
        ElseIf hasLegacy Then
            managed = legacyManaged
        End If
        If fso.FileExists(managed) Then
            Err.Raise MB_ERR + 209, , "A file occupies the managed backup path: " & managed
        End If
        If fso.FolderExists(managed) Then
            If Not MBP_ExactOwner(managed, fso) Then MBP_CheckUnownedEmpty managed, fso
        End If
    End If
    ' The selected base can be ordinary while its managed child is a junction.
    ' Owner recovery intentionally treats unknown evidence as unproven, so do
    ' not rely on that probe to propagate a path-boundary failure. Check the
    ' actual managed destination before any writable-folder probe can run.
    MBP_NoReparse managed, fso
    If Len(managed) > 80 Then
        Err.Raise MB_ERR + 210, , "The managed output root must be at most 80 UTF-16 characters: " & managed
    End If
    MBS_ValidateLayout managed, fso
    MBS_CheckCandidateStore managed, Application.Session.DefaultStore.StoreID, fso
    probeParent = managed
    If Not fso.FolderExists(probeParent) Then probeParent = base
    MBP_ProbeWritable probeParent, fso
    MBP_TryBackupBase = True
    Exit Function
Invalid:
    reason = "Error " & CStr(Err.Number) & ": " & Err.Description
    managed = ""
End Function

Private Sub MBP_ProbeWritable(ByVal folderPath As String, ByVal fso As Object)
    Dim probe As String, stream As Object, created As Boolean
    Dim attempt As Long, n As Long, detail As String, cleanup As String
    On Error GoTo Failed
    For attempt = 1 To 10
        probe = MBH_JoinPath(folderPath, "_mb_write_test_" & fso.GetTempName)
        On Error Resume Next
        Err.Clear
        Set stream = fso.CreateTextFile(probe, False, False)
        n = Err.Number: detail = Err.Description
        On Error GoTo Failed
        If n = 0 Then
            created = True
            Exit For
        End If
        ' 58 means collision: retry without touching the existing file.
        If n <> 58 Then Err.Raise n, "MBP_ProbeWritable", detail
    Next attempt
    If Not created Then Err.Raise MB_ERR + 211, , "Cannot reserve a unique write probe."
    stream.Write "OutlookMailBackup write probe"
    stream.Close
    Set stream = Nothing
    fso.DeleteFile probe, False
    created = False
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If created Then
        Err.Clear
        fso.DeleteFile probe, False
        If Err.Number <> 0 Then cleanup = " Probe cleanup failed: " & probe & ": " & Err.Description
    End If
    On Error GoTo 0
    Err.Raise n, "MBP_ProbeWritable", _
              "The backup folder is not writable: " & folderPath & ". " & detail & cleanup
End Sub

Private Sub MBP_PersistBackupRoot(ByVal selectedBase As String)
    ' Called only AFTER owner/store validation and run-lock acquisition succeed.
    MBP_PersistLocator Application.Session.DefaultStore.StoreID, selectedBase, gRoot
    MBP_SaveSettingText MBP_BASE_KEY, MBH_NormalizePath(selectedBase)
End Sub

Private Function MBP_PickExistingFolder(ByVal title As String, _
                                        ByVal initialPath As String) As String
    ' Shell.BrowseForFolder has a root restriction, not an initial-selection
    ' argument. Show the previous path as context without restricting navigation.
    If Len(initialPath) > 0 Then title = title & vbCrLf & "Previous: " & initialPath
    MBP_PickExistingFolder = MBP_BrowseFileSystemFolder(title)
End Function

Private Function MBP_LegacySettingKey(ByVal managedRoot As String) As String
    MBP_LegacySettingKey = "LegacyRoot_" & _
                          MBH_HashText(LCase$(MBH_NormalizePath(managedRoot)))
End Function

Private Function MBP_ResolveLegacyRoot(ByVal selectedBase As String, _
                                      ByVal managedRoot As String) As String
    Dim remembered As String, candidate As String, fso As Object
    remembered = MBP_GetSettingText(MBP_LegacySettingKey(managedRoot), "")
    If Len(remembered) > 0 Then
        ' Return the saved mapping even if temporarily offline. The caller must
        ' fail closed/report it, not silently interpret every old mail as new.
        MBP_ResolveLegacyRoot = remembered
        Exit Function
    End If
    remembered = MBP_GetSettingText("LastLegacyRoot", "")
    If Len(remembered) > 0 Then
        ' Changing the new-output base must not forget a known old archive.
        ' Preserve this reference even if its drive is temporarily unavailable.
        ' Choose the existing archive with MB_ChooseLegacyFolder to change it.
        MBP_ResolveLegacyRoot = remembered
        Exit Function
    End If
    candidate = MBH_NormalizePath(selectedBase)
    managedRoot = MBH_NormalizePath(managedRoot)
    If StrComp(candidate, managedRoot, vbTextCompare) = 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If StrComp(CStr(fso.GetFileName(candidate)), MBP_MANAGED_LEAF, vbTextCompare) = 0 Or _
           StrComp(CStr(fso.GetFileName(candidate)), MBP_MANAGED_LEAF_LEGACY, vbTextCompare) = 0 Then
            candidate = CStr(fso.GetParentFolderName(candidate))
        End If
    End If
    MBP_ResolveLegacyRoot = candidate
End Function

Private Sub MBP_PersistLegacyRoot(ByVal managedRoot As String, ByVal legacyRoot As String)
    ' Caller verifies readability and its intended base/root mapping first.
    ' Normal backup runs call this only after a complete successful legacy scan.
    ' Keep both a managed-root-specific mapping and a cross-root fallback.
    legacyRoot = MBH_NormalizePath(legacyRoot)
    MBP_SaveSettingText MBP_LegacySettingKey(managedRoot), legacyRoot
    MBP_SaveSettingText "LastLegacyRoot", legacyRoot
End Sub

Private Sub MB_ChooseBackupFolder()
    Dim selectedBase As String, managed As String, held As Boolean
    Dim n As Long, detail As String, releaseIssue As String
    If gBusy Then
        MsgBox "A backup/verification is already running. Finish or cancel it first.", vbInformation
        Exit Sub
    End If
    On Error GoTo Failed
    gBusy = True
    MBP_Reset
    managed = MBP_ChooseBackupRoot(selectedBase)
    If Len(managed) = 0 Then GoTo CleanUp
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    MBH_EnsureDirectory managed
    MBH_AcquireRunLock MBH_JoinPath(managed, "backup.lock")
    held = True
    gRoot = managed: gBaseRoot = managed
    MB_CheckOwnership managed
    MBS_BindStore Application.Session.DefaultStore.StoreID
    MBP_PersistBackupRoot selectedBase
    GoTo CleanUp
Failed:
    n = Err.Number: detail = Err.Description
CleanUp:
    On Error Resume Next
    If held Then
        Err.Clear
        MBH_ReleaseRunLock
        If Err.Number <> 0 Then releaseIssue = " Lock release failed: " & Err.Description
    End If
    Set gFSO = Nothing
    gBusy = False
    On Error GoTo 0
    If n <> 0 Or Len(releaseIssue) > 0 Then
        MsgBox "Cannot finish changing the backup folder." & vbCrLf & _
               detail & releaseIssue, vbExclamation
    ElseIf Len(managed) > 0 Then
        MsgBox "Backup folder remembered: " & selectedBase & vbCrLf & _
               "Managed new exports: " & managed & vbCrLf & _
               "Future runs reuse this folder while it remains available.", vbInformation
    End If
End Sub


' ===================== Existing pre-SafeV2 MSG files ========================
' Adoption reads the original MSG in place; it never saves, copies or renames it.
' Filename, sent-second and bounded date-independent tail keys are candidates only.
' Ordinary by-value attachments are compared by exported byte hashes in owned
' temporary stages. Embedded, OLE and reference attachments require review.
' Unknown/renamed MSG names are indexed by a read-only SentOn lookup when possible.
' Unreadable or undated MSGs remain unknown and defer unmatched mail, rather than
' assuming the absence of a name candidate proves that a message is new.

Private Sub MBL_Reset()
    Set mblByName = Nothing: Set mblByStamp = Nothing: Set mblVisited = Nothing
    gLegacyFileCount = 0: gLegacyFolderCount = 0: gLegacyUnrecognizedNames = 0
    gLegacyNewestFileTime = 0
    gLegacyScanComplete = False
End Sub

Private Function MBL_UnderRoot(ByVal path As String, ByVal root As String) As Boolean
    Dim prefix As String
    If StrComp(path, root, vbTextCompare) = 0 Then
        MBL_UnderRoot = True
        Exit Function
    End If
    prefix = root
    If Right$(prefix, 1) <> "\" Then prefix = prefix & "\"
    MBL_UnderRoot = (StrComp(Left$(path, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

Private Function MBL_IsFileSystemRoot(ByVal path As String) As Boolean
    Dim parts As Variant
    If Len(path) = 3 And Mid$(path, 2, 2) = ":\" Then
        MBL_IsFileSystemRoot = True
    ElseIf Left$(path, 2) = "\\" Then
        parts = Split(Mid$(path, 3), "\")
        MBL_IsFileSystemRoot = (UBound(parts) = 1)
    End If
End Function

Private Function MBL_CanonicalLegacyPath(ByVal path As String, ByVal boundRoot As String, _
                                       Optional ByVal requireExisting As Boolean = False) As String
    Dim root As String, current As String, parent As String, f As Object
    root = MBH_NormalizePath(boundRoot)
    path = MBH_NormalizePath(path)
    If Not MBL_UnderRoot(path, root) Or StrComp(path, root, vbTextCompare) = 0 Then
        Err.Raise MB_ERR + 150, "MBL_CanonicalLegacyPath", "Legacy path is outside its recorded root."
    End If
    If LCase$(Right$(path, 4)) <> ".msg" Then
        Err.Raise MB_ERR + 151, "MBL_CanonicalLegacyPath", "Legacy reference must be an MSG file."
    End If
    ' The default validates structure only, retaining missing-file tombstones.
    If requireExisting Then
        Set f = gFSO.GetFile(path)
        If (CLng(f.Attributes) And MBL_REPARSE_POINT) <> 0 Then
            Err.Raise MB_ERR + 152, "MBL_CanonicalLegacyPath", "Legacy MSG is a link/reparse point."
        End If
        current = gFSO.GetParentFolderName(path)
        Do
            Set f = gFSO.GetFolder(current)
            If (CLng(f.Attributes) And MBL_REPARSE_POINT) <> 0 Then
                Err.Raise MB_ERR + 153, "MBL_CanonicalLegacyPath", "Legacy path traverses a reparse point."
            End If
            If MBL_IsFileSystemRoot(current) Then Exit Do
            parent = gFSO.GetParentFolderName(current)
            If Len(parent) = 0 Or StrComp(parent, current, vbTextCompare) = 0 Then Exit Do
            current = parent
        Loop
    End If
    MBL_CanonicalLegacyPath = path
End Function

Private Function MBL_HasOwner(ByVal folderPath As String) As Boolean
    Dim marker As String, stream As Object, value As String, n As Long, d As String
    On Error GoTo Failed
    If MBR_IsRecoveryDirectory(folderPath) Then
        MBL_HasOwner = True
        Exit Function
    End If
    marker = MBH_JoinPath(folderPath, "owner.txt")
    If Not gFSO.FileExists(marker) Then Exit Function
    Set stream = gFSO.OpenTextFile(marker, 1, False, -1)
    value = stream.ReadAll
    stream.Close
    MBL_HasOwner = (value = MB_OWNER & vbCrLf Or value = MB_OWNER_LEGACY & vbCrLf Or _
                   value = MB_OWNER_PREVIOUS & vbCrLf Or value = MB_OWNER_V34 & vbCrLf)
    Exit Function
Failed:
    n = Err.Number: d = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBL_HasOwner", d
End Function

Private Sub MBL_AddLookup(ByVal lookup As Object, ByVal key As String, ByVal path As String)
    Dim paths As Collection
    If lookup.Exists(key) Then
        Set paths = lookup(key)
    Else
        Set paths = New Collection
        lookup.Add key, paths
    End If
    paths.Add path
End Sub

Private Function MBL_StampKey(ByVal fileName As String) As String
    Dim key As String, value As Date, i As Long, ch As String
    On Error GoTo NotRecognized
    If Len(fileName) < 21 Then Exit Function
    If Mid$(fileName, 9, 1) <> "_" Or Mid$(fileName, 16, 1) <> "_" Then Exit Function
    key = Left$(fileName, 15)
    For i = 1 To 15
        If i <> 9 Then
            ch = Mid$(key, i, 1)
            If ch < "0" Or ch > "9" Then Exit Function
        End If
    Next i
    value = MBC_ParseFileStamp(key)
    If Not MB_ValidDate(value) Then Exit Function
    MBL_StampKey = key
NotRecognized:
End Function

Private Function MBL_TailKey(ByVal fileName As String) As String
    ' Ignore only a syntactically valid old timestamp prefix. Calendar and
    ' time-zone differences may change that prefix without changing the mail.
    ' The retained sender/subject text is still just a candidate, never proof.
    Dim i As Long, ch As String
    If Len(fileName) < 21 Then Exit Function
    If Mid$(fileName, 9, 1) <> "_" Or Mid$(fileName, 16, 1) <> "_" Then Exit Function
    If LCase$(Right$(fileName, 4)) <> ".msg" Then Exit Function
    For i = 1 To 15
        If i <> 9 Then
            ch = Mid$(fileName, i, 1)
            If ch < "0" Or ch > "9" Then Exit Function
        End If
    Next i
    MBL_TailKey = Mid$(fileName, 17)
End Function

Private Function MBL_TailBucket(ByVal fileName As String, ByVal accountRoot As String) As String
    Dim tail As String
    tail = MBL_TailKey(fileName)
    If Len(tail) > 0 Then MBL_TailBucket = "TAIL|" & LCase$(accountRoot) & "|" & tail
End Function

Private Function MBL_CatalogSentStamp(ByVal path As String, ByVal fileName As String, _
                                     ByRef reason As String) As String
    Dim stamp As String
    reason = ""
    ' A recovery scan has already reopened stable bytes. Its hints were read in
    ' THIS run; persisted local timestamps are never substituted across zones.
    If Not gRecoverySentHints Is Nothing Then
        If gRecoverySentHints.Exists(path) Then
            MBL_CatalogSentStamp = CStr(gRecoverySentHints(path))
            Exit Function
        End If
    End If
    If Not gRecoveryForceFull Then stamp = MBL_StampKey(fileName)
    If Len(stamp) = 0 Then stamp = MBL_ReadSentStamp(path, reason)
    MBL_CatalogSentStamp = stamp
End Function

Private Function MBL_ReadSentStamp(ByVal path As String, ByRef reason As String) As String
    Dim obj As Object, mail As Outlook.MailItem, stamp As Date, result As String
    reason = ""
    On Error GoTo Failed
    MBY_NoReparse path, True
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then
        reason = "Renamed MSG is not a MailItem."
        GoTo CleanUp
    End If
    Set mail = obj
    stamp = mail.SentOn
    If MB_ValidDate(stamp) Then
        result = MBC_FileStamp(stamp)
    Else
        reason = "Renamed MSG has no valid SentOn value for candidate lookup."
    End If
CleanUp:
    On Error Resume Next
    Err.Clear
    If Not obj Is Nothing Then obj.Close olDiscard
    If Err.Number <> 0 Then
        result = ""
        reason = reason & " MSG close failed: " & Err.Description
    End If
    Set mail = Nothing: Set obj = Nothing
    On Error GoTo 0
    MBL_ReadSentStamp = result
    Exit Function
Failed:
    result = ""
    reason = "Cannot read renamed MSG SentOn: " & CStr(Err.Number) & ": " & Err.Description
    Resume CleanUp
End Function

Private Sub MBL_BuildIndex(ByVal legacyRoot As String, ByVal managedRoot As String, _
                           Optional ByVal appendCatalog As Boolean = False)
    Dim pending As Collection, current As String, root As String, managed As String
    Dim folder As Object, child As Object, f As Object, path As String, stamp As String, tailBucket As String
    Dim n As Long, d As String, ancestor As String, parent As String, fileTime As Date, why As String
    On Error GoTo Failed
    If Not appendCatalog Then
        MBL_Reset
        Set mblByName = CreateObject("Scripting.Dictionary")
        mblByName.CompareMode = vbTextCompare
        Set mblByStamp = CreateObject("Scripting.Dictionary")
        mblByStamp.CompareMode = vbBinaryCompare
        Set mblVisited = CreateObject("Scripting.Dictionary")
        mblVisited.CompareMode = vbTextCompare
    End If
    gLegacyScanComplete = False
    If Len(legacyRoot) = 0 Then
        gLegacyScanComplete = True
        Exit Sub
    End If
    root = MBH_NormalizePath(legacyRoot)
    managed = MBH_NormalizePath(managedRoot)
    ancestor = root
    Do
        Set folder = gFSO.GetFolder(ancestor)
        If (CLng(folder.Attributes) And MBL_REPARSE_POINT) <> 0 Then
            Err.Raise MB_ERR + 153, "MBL_BuildIndex", "Legacy root traverses a reparse point."
        End If
        If MBL_IsFileSystemRoot(ancestor) Then Exit Do
        parent = gFSO.GetParentFolderName(ancestor)
        If Len(parent) = 0 Or StrComp(parent, ancestor, vbTextCompare) = 0 Then Exit Do
        ancestor = parent
    Loop
    Set pending = New Collection
    pending.Add root
    Do While pending.Count > 0
        If Not MB_ContinueWork() Then Exit Sub
        current = CStr(pending(pending.Count))
        pending.Remove pending.Count
        If Not MBL_UnderRoot(current, managed) Then
            If Not mblVisited.Exists(current) Then
                mblVisited.Add current, True
                Set folder = gFSO.GetFolder(current)
                If (CLng(folder.Attributes) And MBL_REPARSE_POINT) <> 0 Then
                    Err.Raise MB_ERR + 154, "MBL_BuildIndex", _
                              "Legacy scan cannot safely traverse link/reparse folder: " & current
                End If
                If Not MBL_HasOwner(current) Then
                    gLegacyFolderCount = gLegacyFolderCount + 1
                    For Each f In folder.Files
                        If Not MB_ContinueWork() Then Exit Sub
                        If LCase$(gFSO.GetExtensionName(CStr(f.Name))) = "msg" And _
                           Not MBX_IsManagedPath(CStr(f.Path)) Then
                            path = MBL_CanonicalLegacyPath(CStr(f.Path), root)
                            If (CLng(f.Attributes) And MBL_REPARSE_POINT) <> 0 Then
                                Err.Raise MB_ERR + 152, "MBL_BuildIndex", "Legacy MSG is a link/reparse point."
                            End If
                            MBL_AddLookup mblByName, CStr(f.Name), path
                            tailBucket = MBL_TailBucket(CStr(f.Name), root)
                            If Len(tailBucket) > 0 Then MBL_AddLookup mblByName, tailBucket, path
                            ' Existing ordinary catalogs stay metadata-only for
                            ' recognized names. MSG-only recovery uses fresh
                            ' SentOn hints for every file, including newer names.
                            stamp = MBL_CatalogSentStamp(path, CStr(f.Name), why)
                            If Len(stamp) > 0 Then
                                MBL_AddLookup mblByStamp, stamp, path
                            Else
                                gLegacyUnrecognizedNames = gLegacyUnrecognizedNames + 1
                                MB_LogEvent "WARNING", "LEGACY_LOOKUP_UNRESOLVED", "", path, 0, why
                            End If
                            gLegacyFileCount = gLegacyFileCount + 1
                            fileTime = CDate(f.DateLastModified)
                            If fileTime > gLegacyNewestFileTime Then gLegacyNewestFileTime = fileTime
                        End If
                    Next f
                    For Each child In folder.SubFolders
                        pending.Add MBH_NormalizePath(CStr(child.Path))
                    Next child
                End If
            End If
        End If
    Loop
    gLegacyScanComplete = True
    Exit Sub
Failed:
    n = Err.Number: d = Err.Description
    gLegacyScanComplete = False
    Err.Raise n, "MBL_BuildIndex", "Incomplete legacy scan; new exports are blocked. " & d
End Sub

Private Function MBL_OldSanitize(ByVal value As String) As String
    Dim bad As Variant, ch As Variant
    value = Trim$(value)
    bad = Array("/", "\", ":", "*", "?", "'", Chr$(34), "<", ">", "|")
    For Each ch In bad
        value = Replace(value, CStr(ch), "_")
    Next ch
    MBL_OldSanitize = value
End Function

Private Function MBL_OldLeaf(ByVal mail As Outlook.MailItem, ByVal senderLabel As String) As String
    MBL_OldLeaf = MBL_OldSanitize(MBC_FileStamp(mail.SentOn) & "_" & _
                  Left$(Split(senderLabel, "(")(0), 10) & "_" & Left$(mail.Subject, 40)) & ".msg"
End Function

Private Function MBL_AddressBookName(ByVal mail As Outlook.MailItem) As String
    Dim entry As Outlook.AddressEntry
    On Error GoTo Missing
    Set entry = mail.Sender
    If Not entry Is Nothing Then MBL_AddressBookName = entry.Name
Missing:
End Function

Private Sub MBL_AddCandidates(ByVal lookup As Object, ByVal key As String, ByVal candidates As Object, _
                              ByVal accountRoot As String)
    Dim paths As Collection, value As Variant
    If Not lookup.Exists(key) Then Exit Sub
    Set paths = lookup(key)
    ' A very large collision bucket is unresolved, never evidence of absence.
    ' Bound both memory and repeated content work; do not truncate to success.
    If paths.Count > 512 Then Err.Raise MB_ERR + 183, "MBL_AddCandidates", _
        "Legacy candidate bucket exceeds 512 files; all are preserved for review."
    For Each value In paths
        If MB_PathWithin(CStr(value), accountRoot) Then
            If Not candidates.Exists(CStr(value)) Then
                If candidates.Count >= 512 Then Err.Raise MB_ERR + 183, "MBL_AddCandidates", _
                    "Legacy candidate budget exceeded; unvisited files remain."
                candidates.Add CStr(value), True
            End If
        End If
    Next value
End Sub

Private Function MBL_InternetId(ByVal mail As Outlook.MailItem, ByRef present As Boolean) As String
    Dim values As Variant, first As Long
    present = False
    values = MBL_ReadProperties(mail, Array(MBL_ID_UNICODE, MBL_ID_ANSI), 2)
    first = LBound(values)
    MBL_InternetId = MBL_OptionalText(values(first), values(first + 1), "Internet Message-ID")
    present = (Len(MBL_InternetId) > 0)
End Function

Private Function MBL_ReadProperties(ByVal item As Object, ByVal names As Variant, _
                                    ByVal expectedCount As Long) As Variant
    Dim values As Variant
    ' GetProperties exposes each property's Error variant separately. A failed
    ' COM call or malformed result is an error, never proof that a property is absent.
    values = item.PropertyAccessor.GetProperties(names)
    If Not IsArray(values) Then Err.Raise MB_ERR + 177, , "MAPI property result is not an array."
    If UBound(values) - LBound(values) + 1 <> expectedCount Then _
        Err.Raise MB_ERR + 177, , "MAPI property result has the wrong number of values."
    MBL_ReadProperties = values
End Function

Private Function MBL_MissingProperty(ByVal value As Variant, ByVal label As String) As Boolean
    If Not IsError(value) Then Exit Function
    ' CStr(Error) preserves the error number; compare two Error variants after
    ' conversion so neither a localized description nor a Boolean Error is used.
    If CStr(value) = CStr(CVErr(MBL_MAPI_NOT_FOUND)) Then
        MBL_MissingProperty = True
    Else
        Err.Raise MB_ERR + 178, "MBL_MissingProperty", _
                  "Cannot verify " & label & ": " & CStr(value)
    End If
End Function

Private Function MBL_OptionalText(ByVal unicodeValue As Variant, ByVal ansiValue As Variant, _
                                  ByVal label As String) As String
    Dim value As Variant
    value = unicodeValue
    If MBL_MissingProperty(value, label) Then value = ansiValue
    If MBL_MissingProperty(value, label) Then Exit Function
    If VarType(value) <> vbString Then
        Err.Raise MB_ERR + 179, "MBL_OptionalText", "Unexpected type for " & label & "."
    End If
    MBL_OptionalText = CStr(value)
End Function

Private Function MBL_AttachmentProperties(ByVal attachment As Outlook.Attachment) As String
    Dim values As Variant, first As Long, contentId As String, location As String
    Dim mime As String, hidden As Boolean
    values = MBL_ReadProperties(attachment, Array(MBL_CID_UNICODE, MBL_CID_ANSI, _
                               MBL_LOCATION_UNICODE, MBL_LOCATION_ANSI, _
                               MBL_MIME_UNICODE, MBL_MIME_ANSI, MBL_HIDDEN), 7)
    first = LBound(values)
    contentId = MBL_OptionalText(values(first), values(first + 1), "attachment Content-ID")
    location = MBL_OptionalText(values(first + 2), values(first + 3), "attachment Content-Location")
    mime = MBL_OptionalText(values(first + 4), values(first + 5), "attachment MIME type")
    If Not MBL_MissingProperty(values(first + 6), "attachment hidden flag") Then
        If VarType(values(first + 6)) <> vbBoolean Then
            Err.Raise MB_ERR + 180, , "Unexpected type for attachment hidden flag."
        End If
        hidden = CBool(values(first + 6))
    End If
    ' Missing/empty text both mean no identifier; a missing hidden flag is False.
    ' These semantic defaults do not treat access failure as an absent property.
    MBL_AttachmentProperties = MB_EncodeFields(Array(contentId, location, mime, CStr(hidden)))
End Function

Private Sub MBL_CompareAttachmentProperties(ByVal source As Outlook.Attachment, _
                                            ByVal copy As Outlook.Attachment)
    If StrComp(source.DisplayName, copy.DisplayName, vbBinaryCompare) <> 0 Then _
        Err.Raise MB_ERR + 181, , "Attachment display name mismatch."
    If MBL_AttachmentProperties(source) <> MBL_AttachmentProperties(copy) Then _
        Err.Raise MB_ERR + 182, , "Attachment inline reference, MIME type or hidden flag mismatch."
End Sub

Private Function MBL_PropertySelfTest() As Boolean
    Dim absent As Variant, denied As Variant
    absent = CVErr(MBL_MAPI_NOT_FOUND)
    denied = CVErr(&H80070005)
    If Not MBL_MissingProperty(absent, "test") Then Exit Function
    If MBL_MissingProperty("", "test") Then Exit Function
    If MBL_OptionalText("unicode", "ansi", "test") <> "unicode" Then Exit Function
    If MBL_OptionalText(absent, "ansi", "test") <> "ansi" Then Exit Function
    If MBL_OptionalText(absent, absent, "test") <> "" Then Exit Function
    If MBL_OptionalText("", absent, "test") <> "" Then Exit Function
    If Not MBL_TestRejectProperty(denied, "ansi") Then Exit Function
    If Not MBL_TestRejectProperty(absent, denied) Then Exit Function
    If Not MBL_TestRejectProperty(Null, "ansi") Then Exit Function
    If Not MBL_TestRejectProperty(Empty, "ansi") Then Exit Function
    If Not MBL_TestRejectProperty(0, "ansi") Then Exit Function
    If Not MBL_TestRejectProperty(absent, False) Then Exit Function
    MBL_PropertySelfTest = True
End Function

Private Function MBL_TestRejectProperty(ByVal unicodeValue As Variant, ByVal ansiValue As Variant) As Boolean
    Dim ignored As String
    On Error GoTo Rejected
    ignored = MBL_OptionalText(unicodeValue, ansiValue, "test")
    Exit Function
Rejected:
    MBL_TestRejectProperty = True
End Function

Private Function MBL_AttachmentHash(ByVal attachment As Outlook.Attachment) As String
    Dim parent As String, temporary As String, n As Long, d As String, cleanupText As String
    On Error GoTo Failed
    If attachment.Type <> olByValue Then
        Err.Raise MB_ERR + 155, "MBL_AttachmentHash", _
                  "Embedded/OLE/reference attachment requires manual legacy review."
    End If
    If Not MB_ContinueWork() Then Err.Raise MB_ERR + 156, , "Legacy comparison cancelled."
    parent = MBH_JoinPath(gRunDir, "legacy_compare")
    MB_EnsureDir parent
    temporary = MBH_NewStagePath(MBH_JoinPath(parent, "attachment.bin"))
    attachment.SaveAsFile temporary
    MBL_AttachmentHash = MBH_HashFile(temporary)
    MBH_DeleteStage temporary
    temporary = ""
    Exit Function
Failed:
    n = Err.Number: d = Err.Description
    On Error Resume Next
    If Len(temporary) > 0 Then
        Err.Clear
        MBH_DeleteStage temporary
        If Err.Number <> 0 Then cleanupText = " Temporary cleanup: " & Err.Description
    End If
    On Error GoTo 0
    If Len(cleanupText) > 0 Then gCleanupIssue = gCleanupIssue & cleanupText
    Err.Raise n, "MBL_AttachmentHash", d & cleanupText
End Function

Private Function MBL_ByteArraysEqual(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim firstLow As Long, secondLow As Long, last As Long, i As Long
    If VarType(first) <> (vbArray Or vbByte) Or VarType(second) <> (vbArray Or vbByte) Then
        Err.Raise MB_ERR + 175, "MBL_ByteArraysEqual", "RTF body is not a Byte array."
    End If
    firstLow = LBound(first): secondLow = LBound(second): last = UBound(first)
    If last - firstLow <> UBound(second) - secondLow Then Exit Function
    For i = firstLow To last
        If CByte(first(i)) <> CByte(second(secondLow + i - firstLow)) Then Exit Function
        If (i - firstLow) Mod 16384 = 0 Then
            If Not MB_ContinueWork() Then Exit Function
        End If
    Next i
    MBL_ByteArraysEqual = True
End Function

Private Sub MBL_CompareExisting(ByVal source As Outlook.MailItem, ByVal path As String)
    Dim obj As Object, copy As Outlook.MailItem, a As Outlook.Attachment, b As Outlook.Attachment
    Dim ra As Outlook.Recipient, rb As Outlook.Recipient, i As Long
    Dim idSource As String, idCopy As String, hasSource As Boolean, hasCopy As Boolean
    Dim n As Long, d As String, digest As String
    On Error GoTo Failed
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 157, , "Legacy file is not MailItem."
    Set copy = obj
    If StrComp(source.MessageClass, copy.MessageClass, vbBinaryCompare) <> 0 Then _
        Err.Raise MB_ERR + 158, , "Legacy message class mismatch."
    If StrComp(source.Subject, copy.Subject, vbBinaryCompare) <> 0 Then _
        Err.Raise MB_ERR + 159, , "Legacy subject mismatch."
    If source.Sent <> copy.Sent Then Err.Raise MB_ERR + 160, , "Legacy sent-state mismatch."
    If MB_Stamp(source.SentOn) <> MB_Stamp(copy.SentOn) Or _
       MB_Stamp(source.ReceivedTime) <> MB_Stamp(copy.ReceivedTime) Then _
        Err.Raise MB_ERR + 161, , "Legacy sent/received timestamp mismatch."
    If StrComp(source.SenderEmailType, copy.SenderEmailType, vbTextCompare) <> 0 Or _
       StrComp(source.SenderEmailAddress, copy.SenderEmailAddress, vbTextCompare) <> 0 Or _
       StrComp(source.SenderName, copy.SenderName, vbBinaryCompare) <> 0 Then _
        Err.Raise MB_ERR + 162, , "Legacy sender mismatch."
    idSource = MBL_InternetId(source, hasSource)
    idCopy = MBL_InternetId(copy, hasCopy)
    If hasSource <> hasCopy Then Err.Raise MB_ERR + 163, , "Legacy Internet Message-ID presence mismatch."
    If hasSource Then
        If StrComp(idSource, idCopy, vbBinaryCompare) <> 0 Then _
            Err.Raise MB_ERR + 164, , "Legacy Internet Message-ID mismatch."
    End If
    If source.BodyFormat <> copy.BodyFormat Then Err.Raise MB_ERR + 165, , "Legacy body format mismatch."
    If StrComp(source.Body, copy.Body, vbBinaryCompare) <> 0 Then _
        Err.Raise MB_ERR + 166, , "Legacy body mismatch."
    If source.BodyFormat = olFormatHTML Then
        If StrComp(source.HTMLBody, copy.HTMLBody, vbBinaryCompare) <> 0 Then _
            Err.Raise MB_ERR + 167, , "Legacy HTML mismatch."
    ElseIf source.BodyFormat = olFormatRichText Then
        If Not MBL_ByteArraysEqual(source.RTFBody, copy.RTFBody) Then _
            Err.Raise MB_ERR + 176, , "Legacy RTF bytes mismatch."
    End If
    If source.Recipients.Count <> copy.Recipients.Count Then _
        Err.Raise MB_ERR + 168, , "Legacy recipient count mismatch."
    For i = 1 To source.Recipients.Count
        Set ra = source.Recipients.Item(i): Set rb = copy.Recipients.Item(i)
        If ra.Type <> rb.Type Or StrComp(ra.Address, rb.Address, vbBinaryCompare) <> 0 Then _
            Err.Raise MB_ERR + 169, , "Legacy recipient mismatch at index " & CStr(i)
    Next i
    Set ra = Nothing: Set rb = Nothing
    If source.Attachments.Count <> copy.Attachments.Count Then _
        Err.Raise MB_ERR + 170, , "Legacy attachment count mismatch."
    For i = 1 To source.Attachments.Count
        Set a = source.Attachments.Item(i): Set b = copy.Attachments.Item(i)
        If a.Type <> b.Type Or a.Size <> b.Size Or _
           StrComp(a.FileName, b.FileName, vbBinaryCompare) <> 0 Then _
            Err.Raise MB_ERR + 171, , "Legacy attachment metadata mismatch at index " & CStr(i)
        MBL_CompareAttachmentProperties a, b
        digest = MBL_AttachmentHash(a)
        If digest <> MBL_AttachmentHash(b) Then _
            Err.Raise MB_ERR + 172, , "Legacy attachment bytes mismatch at index " & CStr(i)
    Next i
    Set a = Nothing: Set b = Nothing
    MBE_RequireMetadataComplete source, copy
    copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    Exit Sub
Failed:
    n = Err.Number: d = Err.Description
    On Error Resume Next
    Set a = Nothing: Set b = Nothing: Set ra = Nothing: Set rb = Nothing
    If Not copy Is Nothing Then copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    On Error GoTo 0
    Err.Raise n, "MBL_CompareExisting", d
End Sub

Private Function MBL_VerifyCandidate(ByVal mail As Outlook.MailItem, ByVal candidate As String, _
                                    ByVal boundRoot As String, ByRef digest As String, _
                                    ByRef bytes As Double, ByRef reason As String) As Boolean
    Dim beforeHash As String, beforeBytes As Double, beforeDate As Date, f As Object, path As String
    Dim baseline As Variant
    On Error GoTo Rejected
    digest = "": bytes = 0: reason = ""
    path = MBL_CanonicalLegacyPath(candidate, boundRoot, True)
    Set f = gFSO.GetFile(path)
    beforeBytes = CDbl(f.Size): beforeDate = f.DateLastModified
    If beforeBytes = 0 Then Err.Raise MB_ERR + 173, , "Legacy MSG is empty."
    beforeHash = MBH_HashFile(path)
    ' A previously recovered file already has an expected hash. Adoption must
    ' not silently replace that baseline merely because a live mail matches.
    If Not gRecoveryFiles Is Nothing Then
        If gRecoveryFiles.Exists(path) Then
            baseline = MBM_Unpack(CStr(gRecoveryFiles(path)), 5)
            If StrComp(CStr(baseline(0)), path, vbTextCompare) <> 0 Or _
               MB_ByteCount(CStr(baseline(1))) <> beforeBytes Or CStr(baseline(2)) <> beforeHash Then
                Err.Raise MB_ERR + 184, , "Recovered MSG differs from its recorded baseline."
            End If
        End If
    End If
    MBL_CompareExisting mail, path
    digest = MBH_HashFile(path)
    Set f = gFSO.GetFile(path)
    bytes = CDbl(f.Size)
    If bytes <> beforeBytes Or f.DateLastModified <> beforeDate Or digest <> beforeHash Then _
        Err.Raise MB_ERR + 174, , "Legacy file changed during verification."
    MBL_VerifyCandidate = True
    Exit Function
Rejected:
    reason = "Error " & CStr(Err.Number) & ": " & Err.Description
    digest = "": bytes = 0
End Function

Private Function MBL_TryMatch(ByVal mail As Outlook.MailItem, ByRef path As String, _
                             ByRef digest As String, ByRef bytes As Double, _
                             ByRef reason As String, ByVal accountRoot As String) As Long
    Dim candidates As Object, attempted As Object, name As String, stamp As String, candidate As Variant
    Dim candidateHash As String, candidateBytes As Double, why As String, matchCount As Long
    Dim differentBytes As Boolean, failureCount As Long, pass As Long
    On Error GoTo Failed
    path = "": digest = "": bytes = 0: reason = ""
    If Not MB_ContinueWork() Then
        MBL_TryMatch = 3
        Exit Function
    End If
    If Not gLegacyScanComplete Or mblByName Is Nothing Then
        reason = "Legacy scan is incomplete; no new export is permitted."
        MBL_TryMatch = 2
        Exit Function
    End If
    If gLegacyFileCount = 0 Then Exit Function
    Set candidates = CreateObject("Scripting.Dictionary")
    candidates.CompareMode = vbTextCompare
    Set attempted = CreateObject("Scripting.Dictionary")
    attempted.CompareMode = vbTextCompare
    name = MBL_AddressBookName(mail)
    If Len(name) > 0 Then MBL_AddCandidates mblByName, MBL_OldLeaf(mail, name), candidates, accountRoot
    MBL_AddCandidates mblByName, MBL_OldLeaf(mail, mail.SenderName), candidates, accountRoot
    stamp = MBC_FileStamp(mail.SentOn)
    For pass = 1 To 3
        If pass = 2 Then
            ' A complete exact-name match avoids opening unrelated mails sent
            ' in the same second. Every colliding exact-name path was checked.
            If matchCount > 0 Then Exit For
            MBL_AddCandidates mblByStamp, stamp, candidates, accountRoot
        ElseIf pass = 3 Then
            If matchCount > 0 Then Exit For
            ' A timezone/calendar change before first registration may shift
            ' every timestamp lookup. Sender+subject tails provide bounded
            ' candidates; full MSG/content/hash checks remain mandatory.
            If Len(name) > 0 Then MBL_AddCandidates mblByName, _
                MBL_TailBucket(MBL_OldLeaf(mail, name), accountRoot), candidates, accountRoot
            MBL_AddCandidates mblByName, MBL_TailBucket(MBL_OldLeaf(mail, mail.SenderName), accountRoot), _
                              candidates, accountRoot
        End If
        For Each candidate In candidates.Keys
            If Not attempted.Exists(CStr(candidate)) Then
                If Not MB_ContinueWork() Then
                    path = "": digest = "": bytes = 0
                    MBL_TryMatch = 3
                    Exit Function
                End If
                attempted.Add CStr(candidate), True
                why = ""
                If MBL_VerifyCandidate(mail, CStr(candidate), accountRoot, _
                                       candidateHash, candidateBytes, why) Then
                    matchCount = matchCount + 1
                    If Len(path) = 0 Then
                        path = CStr(candidate): digest = candidateHash: bytes = candidateBytes
                    ElseIf candidateHash <> digest Then
                        differentBytes = True
                    ElseIf StrComp(CStr(candidate), path, vbTextCompare) < 0 Then
                        path = CStr(candidate)
                    End If
                Else
                    failureCount = failureCount + 1
                    If Len(reason) < 1200 Then reason = reason & CStr(candidate) & " => " & why & " | "
                End If
            End If
        Next candidate
    Next pass
    If attempted.Count = 0 Then Exit Function
    If Not MB_ContinueWork() Then
        path = "": digest = "": bytes = 0
        MBL_TryMatch = 3
        Exit Function
    End If
    If matchCount = 0 Or differentBytes Then
        If differentBytes Then reason = "Multiple matching legacy files have different bytes. " & reason
        If matchCount = 0 Then reason = "Filename/sent-second/tail candidates could not be verified. " & reason
        path = "": digest = "": bytes = 0
        MBL_TryMatch = 2
    Else
        reason = "Verified in place; matching files=" & CStr(matchCount) & _
                 "; rejected name/sent-second/tail candidates=" & CStr(failureCount)
        MBL_TryMatch = 1
    End If
    Exit Function
Failed:
    reason = "Cannot establish legacy correspondence: " & CStr(Err.Number) & ": " & Err.Description
    path = "": digest = "": bytes = 0
    If gCancel Then
        MBL_TryMatch = 3
    Else
        MBL_TryMatch = 2
    End If
End Function

Private Sub MBL_BuildAccountIndexes(ByVal roots As Collection, ByVal managedRoot As String)
    Dim value As Variant, first As Boolean, totalFiles As Long, totalFolders As Long
    Dim totalUnknown As Long, newest As Date
    first = True
    gLegacyAccountStats.RemoveAll
    For Each value In roots
        If Not MB_ContinueWork() Then Exit Sub
        gLegacyFileCount = 0: gLegacyFolderCount = 0: gLegacyUnrecognizedNames = 0
        gLegacyNewestFileTime = 0
        MBL_BuildIndex CStr(value), managedRoot, Not first
        first = False
        If Not gLegacyScanComplete Then Exit Sub
        gLegacyAccountStats(CStr(value)) = Array(gLegacyFileCount, gLegacyUnrecognizedNames, gLegacyNewestFileTime)
        totalFiles = totalFiles + gLegacyFileCount
        totalFolders = totalFolders + gLegacyFolderCount
        totalUnknown = totalUnknown + gLegacyUnrecognizedNames
        If gLegacyNewestFileTime > newest Then newest = gLegacyNewestFileTime
    Next value
    If first Then MBL_BuildIndex "", managedRoot
    gLegacyFileCount = totalFiles: gLegacyFolderCount = totalFolders
    gLegacyUnrecognizedNames = totalUnknown: gLegacyNewestFileTime = newest
    gLegacyScanComplete = True
End Sub


' =================== Existing-export references and policy ==================
Private Sub MB_RegisterLegacyAll()
    ' No MailItem.SaveAs or persistent separate-attachment export in this mode.
    ' Only verified references / logs are written; temporary attachment bytes
    ' may be written and removed to compare ordinary attachment payloads.
    MB_Run "REGISTER_LEGACY", False, True, False
End Sub

Private Sub MB_ChooseLegacyFolder()
    Dim selectedBase As String, managed As String, legacy As String, held As Boolean
    Dim n As Long, detail As String, cleanup As String, folder As Object
    If gBusy Then
        MsgBox "A backup/verification is already running. Finish or cancel it first.", vbInformation
        Exit Sub
    End If
    On Error GoTo Failed
    gBusy = True
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    managed = MBP_ResolveBackupRoot(selectedBase)
    If Len(managed) = 0 Then GoTo CleanUp
    legacy = MBP_PickExistingFolder("Choose the folder containing the OLD exported MSG files.", _
                                   MBP_ResolveLegacyRoot(selectedBase, managed))
    If Len(legacy) = 0 Then GoTo CleanUp
    legacy = MBH_NormalizePath(legacy)
    If Not gFSO.FolderExists(legacy) Then Err.Raise MB_ERR + 260, , "The legacy folder is unavailable."
    If MB_PathWithin(legacy, managed) Then
        Err.Raise MB_ERR + 261, , "Select the original archive, outside the managed backup tree."
    End If
    Set folder = gFSO.GetFolder(legacy)
    If (CLng(folder.Attributes) And 1024) <> 0 Then
        Err.Raise MB_ERR + 262, , "A reparse-point legacy root cannot be selected."
    End If
    MBH_EnsureDirectory managed
    MBH_AcquireRunLock MBH_JoinPath(managed, "backup.lock")
    held = True
    MB_CheckOwnership managed
    MBP_PersistBackupRoot selectedBase
    MBP_PersistLegacyRoot managed, legacy
    GoTo CleanUp
Failed:
    n = Err.Number: detail = Err.Description
CleanUp:
    On Error Resume Next
    If held Then
        Err.Clear
        MBH_ReleaseRunLock
        If Err.Number <> 0 Then cleanup = " Lock release: " & Err.Description
    End If
    Err.Clear
    MBH_ReleaseCrypto
    If Err.Number <> 0 Then cleanup = cleanup & " Crypto release: " & Err.Description
    Set folder = Nothing: Set gFSO = Nothing
    gBusy = False
    On Error GoTo 0
    If n <> 0 Or Len(cleanup) > 0 Then
        MsgBox "Cannot finish choosing the old archive." & vbCrLf & detail & cleanup, vbExclamation
    ElseIf Len(legacy) > 0 Then
        MsgBox "Old archive remembered: " & legacy & vbCrLf & _
               "Run MB_RegisterLegacyAll to link verified files without exporting mail again.", vbInformation
    End If
End Sub

Private Function MB_PathWithin(ByVal path As String, ByVal root As String) As Boolean
    Dim prefix As String
    path = MBH_NormalizePath(path): root = MBH_NormalizePath(root)
    If StrComp(path, root, vbTextCompare) = 0 Then MB_PathWithin = True: Exit Function
    prefix = root
    If Right$(prefix, 1) <> "\" Then prefix = prefix & "\"
    MB_PathWithin = (StrComp(Left$(path, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

Private Function MB_LegacyBoundPath(ByVal path As String, ByVal root As String) As String
    ' Structural validation only: missing/offline files remain indexed tombstones.
    ' Runtime verification separately checks existence, reparse points and bytes.
    path = MBH_NormalizePath(path): root = MBH_NormalizePath(root)
    If StrComp(path, root, vbTextCompare) = 0 Or Not MB_PathWithin(path, root) Then
        Err.Raise MB_ERR + 263, , "Legacy path is outside its recorded archive root."
    End If
    If LCase$(Right$(path, 4)) <> ".msg" Then
        Err.Raise MB_ERR + 264, , "A legacy reference must point to a MSG file."
    End If
    MB_LegacyBoundPath = path
End Function

Private Function MB_LegacyOwnerKey(ByVal path As String, ByVal digest As String) As String
    If Not MB_IsHexKey(digest) Then Err.Raise MB_ERR + 271, , "Invalid legacy owner file hash."
    ' A path can be reused after a verified move. Bind ownership to its recorded
    ' bytes as well, so a later different file at that path is not the old file.
    MB_LegacyOwnerKey = LCase$(MBH_NormalizePath(path)) & "|" & digest
End Function

Private Sub MB_CacheLegacy(ByVal key As String, ByVal path As String, ByVal bytes As String, _
                           ByVal digest As String, ByVal root As String, ByVal sourceKey As String, _
                           ByVal validation As String)
    Dim canonical As String, owner As String
    canonical = MB_LegacyBoundPath(path, root)
    If validation <> MB_LEGACY_VALIDATION And validation <> MB_LEGACY_VALIDATION_V2 And _
       validation <> MB_LEGACY_VALIDATION_V1 Then _
        Err.Raise MB_ERR + 265, , "Unknown legacy validation policy."
    If MB_ByteCount(bytes) <= 0 Then Err.Raise MB_ERR + 266, , "Empty legacy MSG reference."
    owner = MB_LegacyOwnerKey(canonical, digest)
    If gLegacyOwners.Exists(owner) Then
        If CStr(gLegacyOwners(owner)) <> sourceKey Then
            Err.Raise MB_ERR + 267, , "One legacy file is already assigned to a different source item."
        End If
    End If
    gLegacyOwners(owner) = sourceKey
    gLegacyBindings(key) = MB_EncodeFields(Array(MBH_NormalizePath(root), sourceKey))
    gLegacySources(sourceKey) = key
    gIndex(key) = MB_EncodeFields(Array(canonical, bytes, digest, validation))
End Sub

Private Sub MB_LoadLegacyRow(ByVal f As Variant)
    If CStr(f(13)) <> "0" Then Err.Raise MB_ERR + 268, , "Legacy reference has an attachment ordinal."
    If MB_SourceKey(CStr(f(7)), CStr(f(8))) <> CStr(f(6)) Then
        Err.Raise MB_ERR + 269, , "Legacy source identity checksum does not match."
    End If
    MB_CacheLegacy "L|" & CStr(f(2)), CStr(f(3)), CStr(f(4)), CStr(f(5)), _
                   CStr(f(14)), CStr(f(6)), CStr(f(15))
End Sub

Private Function MB_LegacyBinding(ByVal key As String) As Variant
    Dim fields As Variant
    fields = Split(CStr(gLegacyBindings(key)), vbTab)
    If UBound(fields) <> 1 Then Err.Raise MB_ERR + 270, , "Invalid legacy root binding."
    fields(0) = MB_Unescape(CStr(fields(0)))
    fields(1) = MB_Unescape(CStr(fields(1)))
    MB_LegacyBinding = fields
End Function

Private Function MB_ReuseLegacy(ByVal mail As Outlook.MailItem, ByVal q As Variant, _
                                 ByVal revision As String, ByRef reason As String) As Long
    ' 0=no candidate; 1=reused/adopted; 2=review, never SaveAs; 3=cancel.
    Dim key As String, path As String, root As String, digest As String, bytes As Double
    Dim fields As Variant, binding As Variant, result As Long, n As Long, detail As String
    Dim ownerKey As String, expectedDigest As String, expectedBytes As Double
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    reason = ""
    If gLegacySources.Exists(CStr(q(7))) Then
        key = CStr(gLegacySources(CStr(q(7))))
        fields = MB_IndexFields(key): binding = MB_LegacyBinding(key)
        root = MBM_ResolvedRoot(CStr(fields(0)), CStr(fields(2)), CStr(binding(0)))
        path = MBX_IndexedPath(key, fields, True)
        If Not gFSO.FileExists(path) Then
            reason = "A previously linked old MSG is unavailable. Restore it or review the legacy reference."
            GoTo NeedsReview
        End If
        bytes = CDbl(gFSO.GetFile(path).Size)
        If bytes <> MB_ByteCount(CStr(fields(1))) Then
            reason = "A previously linked old MSG has changed size.": GoTo NeedsReview
        End If
        digest = MBH_HashFile(path)
        If digest <> CStr(fields(2)) Then
            reason = "A previously linked old MSG has changed SHA-256.": GoTo NeedsReview
        End If
        If key = "L|" & revision And CStr(fields(3)) = MB_LEGACY_VALIDATION And _
           Not MBW_RequireContentCheck(CStr(q(7))) Then
            MB_ReuseLegacy = 1
            Exit Function
        End If
        ' Changed revisions and V1 records flow through the same guarded final
        ' check as a new adoption. Preserve the original expected bytes/hash;
        ' MBL_VerifyCandidate overwrites its ByRef result arguments.
    Else
        root = MBY_AccountRoot(CStr(q(1)))
        If Len(root) = 0 Then MB_ReuseLegacy = 3: Exit Function
        result = MBL_TryMatch(mail, path, digest, bytes, reason, root)
        If result = 0 Then
            MB_ReuseLegacy = 0
            Exit Function
        ElseIf result = 3 Then
            MB_ReuseLegacy = 3
            Exit Function
        ElseIf result <> 1 Then
            GoTo NeedsReview
        End If
    End If
    path = MB_LegacyBoundPath(path, root)
    expectedDigest = digest: expectedBytes = bytes
    ownerKey = MB_LegacyOwnerKey(path, expectedDigest)
    If gLegacyOwners.Exists(ownerKey) Then
        If CStr(gLegacyOwners(ownerKey)) <> CStr(q(7)) Then
            reason = "This old MSG is already linked to a different mailbox item."
            GoTo NeedsReview
        End If
    End If
    ' Earlier candidates may have yielded to Outlook or external processes.
    ' Recheck the chosen file while denying write/delete, and keep that guard
    ' until its L reference has actually been committed.
    guard = MBH_GuardFile(path)
    If Not MBL_VerifyCandidate(mail, path, root, digest, bytes, reason) Then GoTo NeedsReview
    If digest <> expectedDigest Or bytes <> expectedBytes Then
        reason = "Chosen legacy file changed since it was selected or last recorded.": GoTo NeedsReview
    End If
    If Not MB_SourceStable(mail, q, revision) Then
        reason = "The source changed during legacy verification.": GoTo NeedsReview
    End If
    MB_RecordCommit "L", revision, path, bytes, digest, q, mail.Subject, 0, root, MB_LEGACY_VALIDATION
    MBH_CloseGuard guard
    MB_LogEvent "INFO", "LEGACY_LINKED_NO_SAVEAS", q(7), path, 0, _
                "Original MSG kept in place; current source verified and indexed."
    MB_ReuseLegacy = 1
    Exit Function
NeedsReview:
    MBH_CloseGuard guard
    If gCancel Then
        MB_ReuseLegacy = 3
    Else
        MB_ReuseLegacy = 2
    End If
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    If gFatal Then Err.Raise n, "MB_ReuseLegacy", detail
    reason = "Legacy verification failed: " & detail
    If gCancel Then
        MB_ReuseLegacy = 3
    Else
        MB_ReuseLegacy = 2
    End If
End Function

Private Function MB_NewExportAllowed(ByVal mail As Outlook.MailItem, ByRef reason As String, _
                                     Optional ByVal verifiedKnownChange As Boolean = False) As Boolean
    Dim folder As Outlook.Folder, root As String, stats As Variant
    If gMode = "BENCHMARK" Then MB_NewExportAllowed = True: Exit Function
    If gJournalIssues > 0 Then
        reason = "The journal has invalid rows; new SaveAs is blocked until the history is reviewed."
        Exit Function
    End If
    If Not gLegacyScanComplete Then
        reason = "The old archive scan did not complete; an absent candidate is not trustworthy."
        Exit Function
    End If
    Set folder = mail.Parent
    root = MBY_GetMappedRoot(folder.StoreID)
    If Len(root) = 0 Or Not gLegacyAccountStats.Exists(root) Then
        reason = "No complete legacy catalog for this account."
        Exit Function
    End If
    stats = gLegacyAccountStats(root)
    ' The old-archive cutoff protects UNKNOWN historical mail from duplicates.
    ' A stable same-source revision with an unchanged prior hash and explicitly
    ' different bounded content is a new version, even at an old received date.
    ' Global index/catalog completion still applies above this narrow bypass.
    If verifiedKnownChange Then MB_NewExportAllowed = True: Exit Function
    If CLng(stats(1)) > 0 Then
        reason = "Some old MSGs in this account have no readable filename/SentOn lookup; unmatched mail needs review."
        Exit Function
    End If
    If CLng(stats(0)) > 0 Then
        If MB_EffectiveDate(mail) <= CDate(stats(2)) Then
            reason = "No verified old-file match; this message predates the old archive's latest write. Automatic re-export is blocked."
            Exit Function
        End If
    End If
    MB_NewExportAllowed = True
End Function


' ===================== Account / year / month layout ========================
' The folder date is SentOn, then ReceivedTime, then CreationTime. Recent-mail
' filtering separately retains its ReceivedTime-first policy. Old original
' YYYYMMDD_HHNNSS prefixes preserve their original SentOn grouping.
' Account selection is remembered in checksummed state; new message files and
' old files organized by the move engine share the actual selected account root.

Private Sub MB_ChooseAccountFolder()
    MB_Run "CHOOSE_ACCOUNT", False, False, False
End Sub

Private Sub MBY_Reset()
    Set mbyRoots = Nothing
    Set mbyValidated = Nothing
    mbyLoadedState = ""
End Sub

Private Function MBY_StateRoot() As String
    If Len(gBaseRoot) = 0 Then Err.Raise MB_ERR + 300, , "Managed state root is not initialized."
    MBY_StateRoot = MBH_NormalizePath(gBaseRoot)
End Function

Private Function MBY_MapPath() As String
    MBY_MapPath = MB_StatePath(MBY_MAP_FILE)
End Function

Private Sub MBY_LoadRoots()
    Dim state As String, path As String, stream As Object, line As String, f As Variant
    Dim key As Variant, other As Variant, n As Long, detail As String
    On Error GoTo Failed
    state = MBY_StateRoot()
    If Not mbyRoots Is Nothing Then
        If StrComp(state, mbyLoadedState, vbTextCompare) = 0 Then Exit Sub
    End If
    Set mbyRoots = CreateObject("Scripting.Dictionary")
    mbyRoots.CompareMode = vbTextCompare
    Set mbyValidated = CreateObject("Scripting.Dictionary")
    mbyValidated.CompareMode = vbTextCompare
    mbyLoadedState = ""
    path = MBY_MapPath()
    If gFSO.FileExists(path) Then
        If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then
            Err.Raise MB_ERR + 301, , "Account mapping has an incomplete UTF-16 byte."
        End If
        Set stream = gFSO.OpenTextFile(path, 1, False, -1)
        Do Until stream.AtEndOfStream
            line = stream.ReadLine
            If Len(line) > 0 Then
                f = MB_ReadCheckedLine(line, 4)
                If f(0) <> "MBAC3" Or Len(CStr(f(1))) = 0 Then
                    Err.Raise MB_ERR + 302, , "Invalid account mapping schema/StoreID."
                End If
                If StrComp(MBH_NormalizePath(CStr(f(3))), state, vbTextCompare) <> 0 Then
                    Err.Raise MB_ERR + 303, , "Account mapping belongs to a different managed state root."
                End If
                mbyRoots(LCase$(CStr(f(1)))) = MBH_NormalizePath(CStr(f(2)))
            End If
        Loop
        stream.Close
        Set stream = Nothing
    End If
    For Each key In mbyRoots.Keys
        For Each other In mbyRoots.Keys
            If StrComp(CStr(key), CStr(other), vbTextCompare) <> 0 Then
                If MB_PathWithin(CStr(mbyRoots(key)), CStr(mbyRoots(other))) Or _
                   MB_PathWithin(CStr(mbyRoots(other)), CStr(mbyRoots(key))) Then
                    Err.Raise MB_ERR + 304, , "Different account roots overlap in the saved mapping."
                End If
            End If
        Next other
    Next key
    mbyLoadedState = state
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    MBY_Reset
    On Error GoTo 0
    Err.Raise n, "MBY_LoadRoots", detail
End Sub

Private Sub MBY_CheckMarker(ByVal root As String, ByVal storeId As String, ByVal requireMarker As Boolean)
    Dim path As String, stream As Object, line As String, f As Variant, n As Long, detail As String
    On Error GoTo Failed
    path = MBH_JoinPath(root, MBY_ACCOUNT_MARKER)
    If Not gFSO.FileExists(path) Then
        If requireMarker Then Err.Raise MB_ERR + 305, , "Account folder marker is missing: " & path
        Exit Sub
    End If
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then
        Err.Raise MB_ERR + 306, , "Account folder marker is truncated."
    End If
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    If stream.AtEndOfStream Then Err.Raise MB_ERR + 307, , "Account folder marker is empty."
    line = stream.ReadLine
    f = MB_ReadCheckedLine(line, 4)
    If Not stream.AtEndOfStream Then Err.Raise MB_ERR + 308, , "Unexpected extra account marker data."
    stream.Close
    Set stream = Nothing
    If f(0) <> "MBAC3" Then Err.Raise MB_ERR + 309, , "Unknown account folder marker."
    If StrComp(CStr(f(1)), storeId, vbTextCompare) <> 0 Then
        Err.Raise MB_ERR + 310, , "This folder is already bound to a different Outlook store."
    End If
    If StrComp(MBH_NormalizePath(CStr(f(2))), root, vbTextCompare) <> 0 Then
        Err.Raise MB_ERR + 311, , "This account folder was moved; its binding needs review."
    End If
    If StrComp(MBH_NormalizePath(CStr(f(3))), MBY_StateRoot(), vbTextCompare) <> 0 Then
        Err.Raise MB_ERR + 312, , "This folder belongs to a different managed backup state."
    End If
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBY_CheckMarker", detail
End Sub

Private Sub MBY_NoReparse(ByVal path As String, Optional ByVal requireExists As Boolean = False)
    Dim current As String, parent As String, f As Object, exists As Boolean
    path = MBH_NormalizePath(path)
    current = path
    Do
        exists = False
        If gFSO.FileExists(current) Then
            Set f = gFSO.GetFile(current): exists = True
        ElseIf gFSO.FolderExists(current) Then
            Set f = gFSO.GetFolder(current): exists = True
        ElseIf requireExists And StrComp(current, path, vbTextCompare) = 0 Then
            Err.Raise MB_ERR + 313, , "Required account path is unavailable: " & path
        End If
        If exists Then
            If (CLng(f.Attributes) And 1024) <> 0 Then
                Err.Raise MB_ERR + 314, , "Account path traverses a link/reparse point: " & current
            End If
        End If
        If MBL_IsFileSystemRoot(current) Then Exit Do
        parent = gFSO.GetParentFolderName(current)
        If Len(parent) = 0 Or StrComp(parent, current, vbTextCompare) = 0 Then Exit Do
        current = parent
    Loop
End Sub

Private Sub MBY_CheckNestedMarkers(ByVal root As String)
    Dim pending As Collection, folder As Object, child As Object, current As String, parent As String
    ' Ancestor markers would make this account a child of another account.
    current = gFSO.GetParentFolderName(root)
    Do While Len(current) > 0
        If gFSO.FileExists(MBH_JoinPath(current, MBY_ACCOUNT_MARKER)) Then
            Err.Raise MB_ERR + 315, , "An account root cannot be nested below another account root."
        End If
        If MBL_IsFileSystemRoot(current) Then Exit Do
        parent = gFSO.GetParentFolderName(current)
        If StrComp(parent, current, vbTextCompare) = 0 Then Exit Do
        current = parent
    Loop
    Set pending = New Collection
    pending.Add root
    Do While pending.Count > 0
        If Not MB_ContinueWork() Then Exit Sub
        current = CStr(pending(pending.Count)): pending.Remove pending.Count
        Set folder = gFSO.GetFolder(current)
        For Each child In folder.SubFolders
            current = MBH_NormalizePath(CStr(child.Path))
            If Not MB_PathWithin(current, MBY_StateRoot()) Then
                If (CLng(child.Attributes) And 1024) <> 0 Then
                    Err.Raise MB_ERR + 316, , "Account tree contains a link/reparse folder: " & current
                End If
                If gFSO.FileExists(MBH_JoinPath(current, MBY_ACCOUNT_MARKER)) Then
                    Err.Raise MB_ERR + 317, , "Selected account folder contains another account root: " & current
                End If
                pending.Add current
            End If
        Next child
    Loop
End Sub

Private Function MBY_ValidateAccountRoot(ByVal path As String, ByVal storeId As String, _
                                        ByVal createMarker As Boolean) As String
    Dim key As Variant, marker As String, stage As String, stream As Object
    Dim n As Long, detail As String
    On Error GoTo Failed
    MBY_LoadRoots
    path = MBH_NormalizePath(path)
    If Len(storeId) = 0 Then Err.Raise MB_ERR + 318, , "An Outlook StoreID is required."
    If Len(path) > 80 Then Err.Raise MB_ERR + 319, , "Account root must be at most 80 UTF-16 characters."
    If Not gFSO.FolderExists(path) Then Err.Raise MB_ERR + 320, , "The account folder is unavailable."
    If MBL_IsFileSystemRoot(path) Then Err.Raise MB_ERR + 321, , "Select an account folder, not a drive/share root."
    If MB_PathWithin(path, MBY_StateRoot()) Then
        Err.Raise MB_ERR + 322, , "Account folders must be outside the managed backup state tree."
    End If
    MBY_NoReparse path, True
    For Each key In mbyRoots.Keys
        If StrComp(CStr(key), storeId, vbTextCompare) <> 0 Then
            If MB_PathWithin(path, CStr(mbyRoots(key))) Or MB_PathWithin(CStr(mbyRoots(key)), path) Then
                Err.Raise MB_ERR + 323, , "This folder overlaps another Outlook account's folder."
            End If
        End If
    Next key
    MBY_CheckMarker path, storeId, Not createMarker
    MBY_CheckRecoveryAccountBinding path, storeId
    If createMarker Then
        MBY_CheckNestedMarkers path
        If gCancel Or gFatal Then Exit Function
        MBP_ProbeWritable path, gFSO
        marker = MBH_JoinPath(path, MBY_ACCOUNT_MARKER)
        If Not gFSO.FileExists(marker) Then
            stage = MBH_NewStagePath(marker)
            Set stream = gFSO.CreateTextFile(stage, True, True)
            MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBAC3", LCase$(storeId), path, MBY_StateRoot()))
            stream.Close
            Set stream = Nothing
            MBH_CommitStage stage, marker
            stage = ""
            MBM_FlushPath marker
        End If
    End If
    MBY_ValidateAccountRoot = path
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(stage) > 0 Then MBH_DeleteStage stage
    On Error GoTo 0
    Err.Raise n, "MBY_ValidateAccountRoot", detail
End Function

Private Sub MBY_SaveRoot(ByVal storeId As String, ByVal path As String)
    Dim stream As Object, stateDir As String, n As Long, detail As String
    On Error GoTo Failed
    MBY_LoadRoots
    If mbyRoots.Exists(storeId) Then
        If StrComp(CStr(mbyRoots(storeId)), path, vbTextCompare) = 0 Then Exit Sub
    End If
    stateDir = MB_StatePath()
    MBH_EnsureDirectory stateDir
    Set stream = gFSO.OpenTextFile(MBY_MapPath(), 8, True, -1)
    stream.WriteLine ""
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBAC3", LCase$(storeId), path, MBY_StateRoot()))
    stream.Close
    Set stream = Nothing
    MBM_FlushPath MBY_MapPath()
    mbyRoots(LCase$(storeId)) = path
    MBS_RememberAccount storeId, path
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBY_SaveRoot", "Cannot remember the account folder: " & detail
End Sub

Private Sub MBY_CheckRecoveryAccountBinding(ByVal account As String, ByVal storeId As String)
    Dim path As String, row As Variant, recoveryRoot As String
    recoveryRoot = MBR_ExistingRecoveryPath(account, gFSO)
    If Len(recoveryRoot) = 0 Then Exit Sub
    path = MBH_JoinPath(recoveryRoot, MBR_OWNER_FILE)
    If Not gFSO.FileExists(path) Then Exit Sub
    row = MB_ReadCheckedLine(MBP_ReadOneLine(path, gFSO), 5)
    If CStr(row(0)) <> "MBRECOVERY1" Or CStr(row(4)) <> "1" Then _
        Err.Raise MB_ERR + 313, , "Unknown account recovery owner."
    If StrComp(MBH_NormalizePath(CStr(row(1))), MBY_StateRoot(), vbTextCompare) <> 0 Or _
       StrComp(MBH_NormalizePath(CStr(row(2))), account, vbTextCompare) <> 0 Or _
       StrComp(CStr(row(3)), storeId, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 313, , "This account recovery folder is bound to a different root, account, or Outlook store."
End Sub

Private Function MBY_AccountBindingUsable(ByVal path As String, ByVal storeId As String, _
                                          ByRef exactBinding As Boolean) As Boolean
    Dim recoveryOwner As String, recoveryRoot As String
    On Error GoTo Unusable
    exactBinding = False
    path = MBH_NormalizePath(path)
    If Not gFSO.FolderExists(path) Then Exit Function
    If Len(path) > 80 Or MBL_IsFileSystemRoot(path) Or MB_PathWithin(path, MBY_StateRoot()) Then Exit Function
    MBY_NoReparse path, True
    MBY_CheckMarker path, storeId, False
    MBY_CheckRecoveryAccountBinding path, storeId
    recoveryRoot = MBR_ExistingRecoveryPath(path, gFSO)
    exactBinding = gFSO.FileExists(MBH_JoinPath(path, MBY_ACCOUNT_MARKER))
    If Len(recoveryRoot) > 0 Then
        recoveryOwner = MBH_JoinPath(recoveryRoot, MBR_OWNER_FILE)
        If gFSO.FileExists(recoveryOwner) Then exactBinding = True
    End If
    MBY_AccountBindingUsable = True
Unusable:
End Function

Private Sub MBY_AddBoundAccount(ByVal path As String, ByVal storeId As String, ByVal candidates As Object)
    Dim exactBinding As Boolean
    If MBY_AccountBindingUsable(path, storeId, exactBinding) Then
        If exactBinding Then candidates(MBH_NormalizePath(path)) = True
    End If
End Sub

Private Sub MBY_DiscoverBoundAccounts(ByVal storeId As String, ByVal candidates As Object)
    Dim bases As Object, value As Variant, base As String, folder As Object
    Set bases = CreateObject("Scripting.Dictionary")
    bases.CompareMode = vbTextCompare
    ' Deliberately bounded discovery: known bases and their direct account
    ' children. Exact StoreID/root markers, never folder names, authorize reuse.
    For Each value In Array(gLegacyRoot, gSelectedBase, CStr(gFSO.GetParentFolderName(gBaseRoot)))
        If Len(CStr(value)) > 0 Then bases(MBH_NormalizePath(CStr(value))) = True
    Next value
    For Each value In bases.Keys
        base = CStr(value)
        If gFSO.FolderExists(base) Then
            MBY_NoReparse base, True
            MBY_AddBoundAccount base, storeId, candidates
            For Each folder In gFSO.GetFolder(base).SubFolders
                If Not MB_ContinueWork() Then Exit Sub
                If (CLng(folder.Attributes) And 1024) = 0 Then MBY_AddBoundAccount CStr(folder.Path), storeId, candidates
            Next folder
        End If
    Next value
End Sub

Private Function MBY_InferAccountRoot(ByVal storeId As String) As String
    Dim st As Outlook.Store, top As Outlook.Folder, names As Variant, leaf As Variant
    Dim candidates As Object, base As String, candidate As String, value As Variant, folderName As String
    Dim exactBinding As Boolean
    Set candidates = CreateObject("Scripting.Dictionary")
    candidates.CompareMode = vbTextCompare
    MBY_DiscoverBoundAccounts storeId, candidates
    If gCancel Or gFatal Then Exit Function
    If candidates.Count > 1 Then Exit Function
    If candidates.Count = 1 Then GoTo UniqueCandidate
    Set st = Application.Session.GetStoreFromID(storeId)
    Set top = st.GetRootFolder
    folderName = MBL_OldSanitize(top.Name)
    names = Array(folderName & "_RAW", folderName & "_byFolder", folderName & "_byTopic", folderName)
    If Len(gLegacyRoot) = 0 Then Exit Function
    base = MBH_NormalizePath(gLegacyRoot)
    For Each leaf In names
        If StrComp(CStr(gFSO.GetFileName(base)), CStr(leaf), vbTextCompare) = 0 Then
            If MBY_AccountBindingUsable(base, storeId, exactBinding) Then candidates(base) = True
        End If
        candidate = MBY_OptionalChild(base, CStr(leaf))
        If Len(candidate) > 0 Then
            If MBY_AccountBindingUsable(candidate, storeId, exactBinding) Then candidates(candidate) = True
        End If
    Next leaf
UniqueCandidate:
    If candidates.Count = 1 Then
        For Each value In candidates.Keys
            MBY_InferAccountRoot = CStr(value)
        Next value
    End If
End Function

Private Function MBY_OptionalChild(ByVal base As String, ByVal leaf As String) As String
    ' An unusable derived account display name must lead to the folder picker.
    ' The chosen real path still receives full validation before it is saved.
    On Error GoTo UnusableName
    MBY_OptionalChild = MBH_JoinPath(base, leaf)
UnusableName:
End Function

Private Function MBY_ResolveAccountRoot(ByVal storeId As String, ByVal forcePicker As Boolean) As String
    Dim candidate As String, reason As String, valid As String, st As Outlook.Store
    Dim n As Long, detail As String
    On Error GoTo Failed
    MBY_LoadRoots
    If Not forcePicker Then
        If mbyRoots.Exists(storeId) And mbyValidated.Exists(storeId) Then
            MBY_ResolveAccountRoot = CStr(mbyRoots(storeId))
            Exit Function
        End If
    End If
    Set st = Application.Session.GetStoreFromID(storeId)
    If Not forcePicker Then
        If mbyRoots.Exists(storeId) Then
            candidate = CStr(mbyRoots(storeId))
            If mbyValidated.Exists(storeId) Then
                MBY_ResolveAccountRoot = candidate
                Exit Function
            End If
            On Error Resume Next
            valid = MBY_ValidateAccountRoot(candidate, storeId, False)
            n = Err.Number: reason = Err.Description
            Err.Clear
            On Error GoTo Failed
            If n = 0 Then
                MBM_AcquireAccountLock valid
                mbyValidated(storeId) = True
                MBY_ResolveAccountRoot = valid
                Exit Function
            End If
            MsgBox "The remembered account folder cannot be used." & vbCrLf & candidate & _
                   vbCrLf & reason, vbExclamation, "Account folder"
        Else
            candidate = MBY_InferAccountRoot(storeId)
            If Len(candidate) > 0 Then
                On Error Resume Next
                Err.Clear
                valid = MBY_ValidateAccountRoot(candidate, storeId, True)
                n = Err.Number: reason = Err.Description
                Err.Clear
                On Error GoTo Failed
                If gCancel Or gFatal Then Exit Function
                If n = 0 Then
                    MBM_AcquireAccountLock valid
                    MBY_SaveRoot storeId, valid
                    mbyValidated(storeId) = True
                    MBY_ResolveAccountRoot = valid
                    Exit Function
                End If
                ' An inferred display-name match is only a candidate. Invalid
                ' markers, path length or nested roots must still allow selection.
                MsgBox "The automatically detected account folder cannot be used." & vbCrLf & _
                       candidate & vbCrLf & reason, vbExclamation, "Account folder"
            End If
        End If
    ElseIf mbyRoots.Exists(storeId) Then
        candidate = CStr(mbyRoots(storeId))
    End If
    Do
        candidate = MBP_PickExistingFolder("Choose the EXISTING account folder for: " & st.DisplayName & _
                    vbCrLf & "New exports use yyyy\mm inside this folder.", candidate)
        If Len(candidate) = 0 Then
            gCancel = True
            Exit Function
        End If
        On Error Resume Next
        Err.Clear
        valid = MBY_ValidateAccountRoot(candidate, storeId, True)
        n = Err.Number: reason = Err.Description
        On Error GoTo Failed
        If gCancel Or gFatal Then Exit Function
        If n = 0 Then
            MBM_AcquireAccountLock valid
            MBY_SaveRoot storeId, valid
            mbyValidated(storeId) = True
            MBY_ResolveAccountRoot = valid
            Exit Function
        End If
        MsgBox reason & vbCrLf & "Select another account folder or Cancel.", vbExclamation
    Loop
Failed:
    n = Err.Number: detail = Err.Description
    gFatal = True
    gFatalText = "Account folder setup failed: " & detail
    Err.Raise n, "MBY_ResolveAccountRoot", gFatalText
End Function

Private Function MBY_AccountRoot(ByVal storeId As String) As String
    MBY_AccountRoot = MBY_ResolveAccountRoot(storeId, False)
End Function

Private Function MBY_ChooseAccountRoot(ByVal storeId As String) As String
    MBY_ChooseAccountRoot = MBY_ResolveAccountRoot(storeId, True)
End Function

Private Function MBY_GetMappedRoot(ByVal storeId As String) As String
    ' No folder picker or filesystem mutation. Missing mappings return empty.
    MBY_LoadRoots
    If mbyRoots.Exists(storeId) Then MBY_GetMappedRoot = CStr(mbyRoots(storeId))
End Function

Private Function MBY_GetKnownRootsAsCollection() As Collection
    Dim roots As Collection, key As Variant
    MBY_LoadRoots
    Set roots = New Collection
    For Each key In mbyRoots.Keys
        roots.Add CStr(mbyRoots(key))
    Next key
    Set MBY_GetKnownRootsAsCollection = roots
End Function

Private Function MBY_CanonicalFilePath(ByVal path As String, ByVal boundRoot As String, _
                                       Optional ByVal requireExisting As Boolean = False) As String
    path = MBH_NormalizePath(path): boundRoot = MBH_NormalizePath(boundRoot)
    If StrComp(path, boundRoot, vbTextCompare) = 0 Or Not MB_PathWithin(path, boundRoot) Then
        Err.Raise MB_ERR + 326, , "File path is outside its recorded account root."
    End If
    If requireExisting Then
        If Not gFSO.FileExists(path) Then Err.Raise MB_ERR + 327, , "Referenced file is unavailable: " & path
        MBY_NoReparse path, True
    End If
    MBY_CanonicalFilePath = path
End Function

Private Function MBY_MessageDate(ByVal mail As Outlook.MailItem) As Date
    Dim value As Date
    value = mail.SentOn
    If MB_ValidDate(value) Then MBY_MessageDate = value: Exit Function
    value = mail.ReceivedTime
    If MB_ValidDate(value) Then MBY_MessageDate = value: Exit Function
    value = mail.CreationTime
    If MB_ValidDate(value) Then MBY_MessageDate = value: Exit Function
    Err.Raise MB_ERR + 324, , "The mail has no usable date for year/month filing."
End Function

Private Function MBY_MonthDirectory(ByVal accountRoot As String, ByVal stamp As Date) As String
    Dim result As String, yearMonth As String
    If Not MB_ValidDate(stamp) Then Err.Raise MB_ERR + 325, , "Invalid year/month filing date."
    yearMonth = MBC_YearMonth(stamp)
    result = MBH_JoinPath(MBH_JoinPath(accountRoot, Left$(yearMonth, 4)), Right$(yearMonth, 2))
    MBY_NoReparse result
    MBY_MonthDirectory = result
End Function

Private Function MBY_FileMonth(ByVal path As String, ByRef stamp As Date, ByRef reason As String) As Boolean
    Dim leaf As String, ext As String, key As String, remaining As String, separator As Long
    Dim obj As Object, mail As Outlook.MailItem, n As Long, detail As String
    On Error GoTo Failed
    stamp = 0: reason = ""
    leaf = gFSO.GetFileName(path)
    ext = LCase$(gFSO.GetExtensionName(path))
    key = MBL_StampKey(leaf)
    If Len(key) > 0 Then
        stamp = MBC_ParseFileStamp(key)
        If Not MB_ValidDate(stamp) Then stamp = 0
    End If
    If ext = "msg" Then
        If stamp <> 0 Then
            reason = "Original SentOn filename prefix"
            MBY_FileMonth = True
            Exit Function
        End If
        MBY_NoReparse path, True
        Set obj = Application.Session.OpenSharedItem(path)
        If Not TypeOf obj Is Outlook.MailItem Then
            reason = "MSG is not an Outlook MailItem."
            GoTo CleanUp
        End If
        Set mail = obj
        stamp = MBY_MessageDate(mail)
        reason = "MSG SentOn/ReceivedTime/CreationTime"
        MBY_FileMonth = True
    ElseIf Len(key) > 0 And stamp <> 0 And Len(ext) > 0 Then
        If InStr(1, MBY_ATTACHMENT_EXTENSIONS, ";" & ext & ";", vbBinaryCompare) > 0 Then
            remaining = Mid$(leaf, 17)
            separator = InStr(remaining, "_")
            If separator >= 1 And separator <= 11 And Len(Mid$(remaining, separator + 1)) > Len(ext) + 1 Then
                reason = "Original extracted-attachment filename pattern"
                MBY_FileMonth = True
            End If
        End If
    End If
CleanUp:
    If Not mail Is Nothing Then
        mail.Close olDiscard
    ElseIf Not obj Is Nothing Then
        obj.Close olDiscard
    End If
    Set mail = Nothing: Set obj = Nothing
    If Not MBY_FileMonth Then
        stamp = 0
        If Len(reason) = 0 Then reason = "Unrecognized/non-mail file; left unchanged."
    End If
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not mail Is Nothing Then
        mail.Close olDiscard
    ElseIf Not obj Is Nothing Then
        obj.Close olDiscard
    End If
    Set mail = Nothing: Set obj = Nothing
    On Error GoTo 0
    stamp = 0: reason = "Cannot determine a filing date: " & CStr(n) & ": " & detail
    MBY_FileMonth = False
End Function



' =================== Year/month organization and recovery ==================
' Callers hold gRoot's run lock and initialize gFSO/gDirs/gEvents/gRunDir.
' A preview is complete only when its checked END row has been closed/flushed.
' Apply revalidates the entire plan before the first file move.
' MoveFileExW uses WRITE_THROUGH only: no REPLACE_EXISTING, no COPY_ALLOWED.
' The account tree and already committed managed files are the only sources.
' No old MSG is saved again. Identical duplicates are retained as separate files.

Private Sub MBM_Reset()
    Set mbmOps = Nothing: Set mbmRedirects = Nothing
    mbmLoaded = False
    gMovePlanned = 0: gMoveKept = 0: gMoveSkipped = 0
    gMoveApplied = 0: gMoveAlready = 0: gMoveRestored = 0
    gMoveRecovered = 0: gMoveReview = 0: gMovePlanReview = 0: gMovePlanPath = ""
End Sub

Private Sub MBM_AcquireAccountLock(ByVal accountRoot As String)
    Dim path As String, winError As Long, n As Long, detail As String, registered As Boolean
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    On Error GoTo Failed
    accountRoot = MBH_NormalizePath(accountRoot)
    If mbmAccountLocks Is Nothing Then
        Set mbmAccountLocks = CreateObject("Scripting.Dictionary")
        mbmAccountLocks.CompareMode = vbTextCompare
    End If
    If mbmAccountLocks.Exists(accountRoot) Then Exit Sub
    MBM_CheckAncestors accountRoot
    path = MBH_JoinPath(accountRoot, "_mb_account.lock")
    If gFSO.FileExists(path) Then
        If (CLng(gFSO.GetFile(path).Attributes) And 1024) <> 0 Then
            Err.Raise MB_ERR + 404, , "Account lock cannot be a link/reparse point."
        End If
    End If
    handle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_READ Or MBH_GENERIC_WRITE, 0, 0, _
                            MBH_OPEN_ALWAYS, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If handle = -1 Then
        winError = Err.LastDllError
        Err.Raise MB_ERR + 400, "MBM_AcquireAccountLock", _
                  "Account archive is busy or unavailable. Win32=" & CStr(winError) & "; " & accountRoot
    End If
    mbmAccountLocks.Add accountRoot, handle
    registered = True
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    If Not registered And handle <> 0 And handle <> -1 Then MBH_CloseHandle handle
    Err.Raise n, "MBM_AcquireAccountLock", detail
End Sub

Private Sub MBM_ReleaseAccountLocks()
    Dim key As Variant, failed As String, winError As Long
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    If mbmAccountLocks Is Nothing Then Exit Sub
    For Each key In mbmAccountLocks.Keys
#If VBA7 Then
        handle = CLngPtr(mbmAccountLocks(key))
#Else
        handle = CLng(mbmAccountLocks(key))
#End If
        If MBH_CloseHandle(handle) = 0 Then
            winError = Err.LastDllError
            failed = failed & " " & CStr(key) & " (" & CStr(winError) & ")"
        Else
            mbmAccountLocks.Remove key
        End If
    Next key
    If mbmAccountLocks.Count = 0 Then Set mbmAccountLocks = Nothing
    If Len(failed) > 0 Then Err.Raise MB_ERR + 401, "MBM_ReleaseAccountLocks", "Cannot release account locks:" & failed
End Sub

Private Sub MBM_CheckAncestors(ByVal directory As String)
    Dim current As String, parent As String, folder As Object
    current = MBH_NormalizePath(directory)
    Do
        Set folder = gFSO.GetFolder(current)
        If (CLng(folder.Attributes) And 1024) <> 0 Then
            Err.Raise MB_ERR + 402, "MBM_CheckAncestors", "Link/reparse directories are not moved: " & current
        End If
        If MBL_IsFileSystemRoot(current) Then Exit Do
        parent = CStr(gFSO.GetParentFolderName(current))
        If Len(parent) = 0 Or StrComp(parent, current, vbTextCompare) = 0 Then Exit Do
        current = parent
    Loop
End Sub

Private Function MBM_File(ByVal path As String, ByVal boundRoot As String, _
                          Optional ByVal requireExisting As Boolean = True) As String
    Dim f As Object
    path = MBH_NormalizePath(path): boundRoot = MBH_NormalizePath(boundRoot)
    If Not MB_PathWithin(path, boundRoot) Or StrComp(path, boundRoot, vbTextCompare) = 0 Then
        Err.Raise MB_ERR + 403, "MBM_File", "A move path is outside its recorded root."
    End If
    If requireExisting Then
        MBM_CheckAncestors CStr(gFSO.GetParentFolderName(path))
        Set f = gFSO.GetFile(path)
        If (CLng(f.Attributes) And 1024) <> 0 Then
            Err.Raise MB_ERR + 404, "MBM_File", "Link/reparse files are not moved: " & path
        End If
    End If
    MBM_File = path
End Function

Private Function MBM_JournalPath() As String
    MBM_JournalPath = MB_StatePath("moves.tsv")
End Function

Private Sub MBM_FlushPath(ByVal path As String)
    Dim winError As Long, closeError As Long
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    handle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_WRITE, 3, 0, _
                            MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If handle = -1 Then
        winError = Err.LastDllError
        Err.Raise MB_ERR + 405, "MBM_FlushPath", "Cannot open durable move state. Win32=" & CStr(winError)
    End If
    If MBM_FlushFileBuffers(handle) = 0 Then winError = Err.LastDllError
    If MBH_CloseHandle(handle) = 0 Then closeError = Err.LastDllError
    If winError <> 0 Or closeError <> 0 Then
        Err.Raise MB_ERR + 406, "MBM_FlushPath", _
                  "Cannot flush/close move state. Win32=" & CStr(winError) & "/" & CStr(closeError)
    End If
End Sub

Private Function MBM_Unpack(ByVal value As String, ByVal count As Long) As Variant
    Dim fields As Variant, i As Long
    fields = Split(value, vbTab)
    If UBound(fields) <> count - 1 Then Err.Raise MB_ERR + 407, , "Invalid move state field count."
    For i = 0 To UBound(fields)
        fields(i) = MB_Unescape(CStr(fields(i)))
    Next i
    MBM_Unpack = fields
End Function

Private Function MBM_RedirectKey(ByVal path As String, ByVal digest As String) As String
    ' One old path can be reused by different files after its original moved.
    ' Keep every content version's route; a new route must not erase the old one.
    MBM_RedirectKey = LCase$(MBH_NormalizePath(path)) & "|" & LCase$(digest)
End Function

Private Sub MBM_SetRedirect(ByVal source As String, ByVal target As String, _
                             ByVal digest As String, ByVal targetRoot As String)
    Dim targetKey As String
    ' A reverse move supersedes only this hash's old forward edge.
    ' Other files that previously occupied the same pathname retain their route.
    targetKey = MBM_RedirectKey(target, digest)
    If mbmRedirects.Exists(targetKey) Then mbmRedirects.Remove targetKey
    mbmRedirects(MBM_RedirectKey(source, digest)) = MB_EncodeFields(Array(target, digest, targetRoot))
End Sub

Private Function MBM_ResolvePath(ByVal original As String, ByVal expectedHash As String) As String
    Dim path As String, fields As Variant, seen As Object, key As String
    path = MBH_NormalizePath(original)
    If Not mbmLoaded Then Err.Raise MB_ERR + 408, , "Move journal must be loaded before resolving files."
    key = MBM_RedirectKey(path, expectedHash)
    If Not mbmRedirects.Exists(key) Then
        MBM_ResolvePath = path
        Exit Function
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Do While mbmRedirects.Exists(key)
        If seen.Exists(path) Then Err.Raise MB_ERR + 409, , "Cycle in move history; no export permitted."
        seen.Add path, True
        fields = MBM_Unpack(CStr(mbmRedirects(key)), 3)
        If CStr(fields(1)) <> expectedHash Then Err.Raise MB_ERR + 455, , "Move redirect hash mismatch."
        path = CStr(fields(0))
        key = MBM_RedirectKey(path, expectedHash)
    Loop
    MBM_ResolvePath = path
End Function

Private Function MBM_ResolvedRoot(ByVal original As String, ByVal expectedHash As String, _
                                   ByVal fallbackRoot As String) As String
    Dim path As String, fields As Variant, root As String, seen As Object, key As String
    path = MBH_NormalizePath(original): root = MBH_NormalizePath(fallbackRoot)
    If Not mbmLoaded Then Err.Raise MB_ERR + 408, , "Move journal must be loaded before resolving files."
    key = MBM_RedirectKey(path, expectedHash)
    If Not mbmRedirects.Exists(key) Then
        MBM_ResolvedRoot = root
        Exit Function
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Do While mbmRedirects.Exists(key)
        If seen.Exists(path) Then Err.Raise MB_ERR + 409, , "Cycle in move history; no export permitted."
        seen.Add path, True
        fields = MBM_Unpack(CStr(mbmRedirects(key)), 3)
        If CStr(fields(1)) <> expectedHash Then Err.Raise MB_ERR + 455, , "Move redirect hash mismatch."
        path = CStr(fields(0)): root = CStr(fields(2))
        key = MBM_RedirectKey(path, expectedHash)
    Loop
    MBM_ResolvedRoot = root
End Function

Private Function MBM_OperationId(ByVal planId As String, ByVal source As String, _
                                  ByVal target As String, ByVal digest As String) As String
    MBM_OperationId = MBH_HashText(MB_EncodeFields(Array("MBMOVE3", planId, _
                           LCase$(source), LCase$(target), digest)))
End Function

Private Sub MBM_ValidateMoveRow(ByVal f As Variant)
    Dim p As String, root As String, sourceRoot As String
    If f(0) <> "MBMOVE3" Then Err.Raise MB_ERR + 410, , "Wrong move journal schema."
    If Not MB_IsHexKey(CStr(f(1))) Or Not MB_IsHexKey(CStr(f(2))) Or _
       Not MB_IsHexKey(CStr(f(9))) Then Err.Raise MB_ERR + 411, , "Invalid move ID/hash."
    If f(3) <> "PREPARE" And f(3) <> "COMMIT" And f(3) <> "ABORT" And _
       f(3) <> "UNDO_PREPARE" And f(3) <> "UNDO_COMMIT" And f(3) <> "UNDO_ABORT" Then
        Err.Raise MB_ERR + 412, , "Unknown move phase."
    End If
    root = MBH_NormalizePath(CStr(f(4)))
    sourceRoot = MBH_NormalizePath(CStr(f(5)))
    If StrComp(sourceRoot, root, vbTextCompare) <> 0 And _
       StrComp(sourceRoot, gRoot, vbTextCompare) <> 0 Then
        Err.Raise MB_ERR + 444, , "Move source root is outside this archive."
    End If
    p = MBM_File(CStr(f(6)), sourceRoot, False)
    p = MBM_File(CStr(f(7)), root, False)
    If StrComp(CStr(f(6)), CStr(f(7)), vbTextCompare) = 0 Then
        Err.Raise MB_ERR + 413, , "Move source and destination are equal."
    End If
    If MB_ByteCount(CStr(f(8))) < 0 Then Err.Raise MB_ERR + 414, , "Invalid move byte count."
    If CStr(f(1)) <> MBM_OperationId(CStr(f(2)), CStr(f(6)), CStr(f(7)), CStr(f(9))) Then
        Err.Raise MB_ERR + 415, , "Move operation identity mismatch."
    End If
End Sub

Private Sub MBM_CacheMove(ByVal f As Variant)
    Dim prior As Variant, phase As String, i As Long
    phase = CStr(f(3))
    If mbmOps.Exists(CStr(f(1))) Then
        prior = MBM_Unpack(CStr(mbmOps(CStr(f(1)))), 11)
        For i = 0 To 9
            If i <> 3 Then
                If CStr(prior(i)) <> CStr(f(i)) Then Err.Raise MB_ERR + 416, , "Move operation fields changed."
            End If
        Next i
    ElseIf phase <> "PREPARE" Then
        Err.Raise MB_ERR + 417, , "Move record has no prior PREPARE."
    End If
    If phase = "COMMIT" Then
        If Not mbmOps.Exists(CStr(f(1))) Then Err.Raise MB_ERR + 417, , "COMMIT has no PREPARE."
        If prior(3) <> "PREPARE" Then Err.Raise MB_ERR + 418, , "Unexpected forward COMMIT."
        MBM_SetRedirect CStr(f(6)), CStr(f(7)), CStr(f(9)), CStr(f(4))
    ElseIf phase = "UNDO_COMMIT" Then
        If prior(3) <> "UNDO_PREPARE" Then Err.Raise MB_ERR + 419, , "Unexpected undo COMMIT."
        MBM_SetRedirect CStr(f(7)), CStr(f(6)), CStr(f(9)), CStr(f(5))
    ElseIf phase = "UNDO_PREPARE" Then
        If prior(3) <> "COMMIT" And prior(3) <> "UNDO_ABORT" Then
            Err.Raise MB_ERR + 420, , "Only a completed move can be undone."
        End If
    ElseIf phase = "PREPARE" And mbmOps.Exists(CStr(f(1))) Then
        If prior(3) <> "ABORT" And prior(3) <> "UNDO_COMMIT" Then
            Err.Raise MB_ERR + 421, , "Unexpected repeated PREPARE."
        End If
    ElseIf phase = "ABORT" Then
        If prior(3) <> "PREPARE" Then Err.Raise MB_ERR + 422, , "Unexpected move ABORT."
    ElseIf phase = "UNDO_ABORT" Then
        If prior(3) <> "UNDO_PREPARE" Then Err.Raise MB_ERR + 423, , "Unexpected undo ABORT."
    End If
    mbmOps(CStr(f(1))) = MB_EncodeFields(f)
End Sub

Private Sub MBM_AppendMove(ByVal fields As Variant, ByVal phase As String)
    Dim stream As Object, payload As String, path As String, n As Long, detail As String
    On Error GoTo Failed
    fields(3) = phase: fields(10) = MB_Stamp(Now)
    MBM_ValidateMoveRow fields
    payload = MB_EncodeFields(fields)
    path = MBM_JournalPath()
    Set stream = gFSO.OpenTextFile(path, 8, True, -1)
    MB_WriteCheckedLine stream, payload
    stream.Close
    Set stream = Nothing
    MBM_FlushPath path
    MBM_CacheMove fields
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    gFatal = True
    gFatalText = "Move journal cannot be persisted: " & detail
    Err.Raise n, "MBM_AppendMove", gFatalText
End Sub

Private Sub MBM_VerifyFile(ByVal path As String, ByVal root As String, _
                            ByVal expectedBytes As String, ByVal expectedHash As String)
    path = MBM_File(path, root, True)
    If CDbl(gFSO.GetFile(path).Size) <> MB_ByteCount(expectedBytes) Then
        Err.Raise MB_ERR + 424, , "Move source/destination size changed: " & path
    End If
    If MBH_HashFile(path) <> expectedHash Then
        Err.Raise MB_ERR + 425, , "Move source/destination hash changed: " & path
    End If
End Sub

Private Sub MBM_Recover(ByVal f As Variant)
    Dim source As String, target As String, sourceRoot As String, targetRoot As String
    Dim reverse As Boolean, sourceExists As Boolean, targetExists As Boolean
    reverse = (f(3) = "UNDO_PREPARE")
    source = CStr(f(6)): target = CStr(f(7))
    sourceRoot = CStr(f(5)): targetRoot = CStr(f(4))
    If reverse Then
        source = CStr(f(7)): target = CStr(f(6))
        sourceRoot = CStr(f(4)): targetRoot = CStr(f(5))
    End If
    MBM_AcquireAccountLock CStr(f(4))
    sourceExists = gFSO.FileExists(source): targetExists = gFSO.FileExists(target)
    If sourceExists And Not targetExists Then
        MBM_VerifyFile source, sourceRoot, CStr(f(8)), CStr(f(9))
        If reverse Then MBM_AppendMove f, "UNDO_ABORT" Else MBM_AppendMove f, "ABORT"
    ElseIf targetExists And Not sourceExists Then
        MBM_VerifyFile target, targetRoot, CStr(f(8)), CStr(f(9))
        If reverse Then MBM_AppendMove f, "UNDO_COMMIT" Else MBM_AppendMove f, "COMMIT"
        gMoveRecovered = gMoveRecovered + 1
    Else
        gMoveReview = gMoveReview + 1
        Err.Raise MB_ERR + 426, "MBM_Recover", _
                  "Unresolved interrupted move (both paths exist or both are missing): " & source & " => " & target
    End If
End Sub

Private Sub MBM_LoadMoves()
    Dim stream As Object, line As String, f As Variant, path As String, key As Variant
    Dim n As Long, detail As String
    On Error GoTo Failed
    mbmLoaded = False
    Set mbmOps = CreateObject("Scripting.Dictionary")
    Set mbmRedirects = CreateObject("Scripting.Dictionary")
    mbmOps.CompareMode = vbBinaryCompare: mbmRedirects.CompareMode = vbTextCompare
    path = MBM_JournalPath()
    If gFSO.FileExists(path) Then
        If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 427, , "Move journal is truncated (odd UTF-16 byte count)."
        Set stream = gFSO.OpenTextFile(path, 1, False, -1)
        Do Until stream.AtEndOfStream
            If Not MB_ContinueWork() Then GoTo Cancelled
            line = stream.ReadLine
            If Len(line) > 0 Then
                f = MB_ReadCheckedLine(line, 11)
                MBM_ValidateMoveRow f
                MBM_CacheMove f
            End If
        Loop
        stream.Close
        Set stream = Nothing
    End If
    For Each key In mbmOps.Keys
        If Not MB_ContinueWork() Then GoTo Cancelled
        f = MBM_Unpack(CStr(mbmOps(key)), 11)
        If f(3) = "PREPARE" Or f(3) = "UNDO_PREPARE" Then MBM_Recover f
    Next key
    mbmLoaded = True
    Exit Sub
Cancelled:
    If Not stream Is Nothing Then stream.Close
    mbmLoaded = False
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    gFatal = True
    gFatalText = "Move history requires recovery/review; new exports are blocked. " & detail
    Err.Raise n, "MBM_LoadMoves", gFatalText
End Sub

Private Function MBM_Destination(ByVal source As String, ByVal parent As String, _
                                  ByVal reserved As Object) As String
    Dim leaf As String, stem As String, extension As String, suffix As String
    Dim candidate As String, budget As Long, n As Long, identity As String
    leaf = CStr(gFSO.GetFileName(source))
    extension = CStr(gFSO.GetExtensionName(source))
    If Len(extension) > 0 Then extension = "." & extension
    stem = CStr(gFSO.GetBaseName(source))
    identity = Left$(MBH_HashText(LCase$(source)), 16)
    budget = MBH_MAX_PATH_UNITS - Len(parent) - 1
    If Len(leaf) > budget Then
        suffix = "__" & identity
        leaf = MBH_SafeComponent(stem, budget - Len(extension) - Len(suffix), "file") & suffix & extension
    End If
    candidate = MBH_JoinPath(parent, leaf)
    If StrComp(candidate, source, vbTextCompare) = 0 Then
        MBM_Destination = source
        Exit Function
    End If
    If Not gFSO.FileExists(candidate) And Not gFSO.FolderExists(candidate) And Not reserved.Exists(candidate) Then
        reserved.Add candidate, True
        MBM_Destination = candidate
        Exit Function
    End If
    For n = 0 To 99999
        suffix = "__" & identity
        If n > 0 Then suffix = suffix & "_" & Format$(n, "00000")
        leaf = MBH_SafeComponent(stem, budget - Len(extension) - Len(suffix), "file") & suffix & extension
        candidate = MBH_JoinPath(parent, leaf)
        If Not gFSO.FileExists(candidate) And Not gFSO.FolderExists(candidate) And Not reserved.Exists(candidate) Then
            reserved.Add candidate, True
            MBM_Destination = candidate
            Exit Function
        End If
    Next n
    Err.Raise MB_ERR + 428, , "Cannot reserve a distinct year/month destination."
End Function

Private Sub MBM_PlanItem(ByVal output As Object, ByVal planId As String, ByVal accountRoot As String, _
                          ByVal sourceRoot As String, ByVal source As String, ByVal reserved As Object, _
                          ByVal seen As Object, Optional ByVal knownDate As Date = 0, _
                          Optional ByVal expectedHash As String = "")
    Dim stamp As Date, reason As String, parent As String, target As String, f As Object
    Dim kind As String, bytes As String, digest As String, beforeDate As Date
    Dim n As Long, detail As String
    On Error GoTo Review
    source = MBM_File(source, sourceRoot, True)
    If seen.Exists(source) Then Exit Sub
    seen.Add source, True
    stamp = knownDate
    If stamp = 0 Then
        If Not MBY_FileMonth(source, stamp, reason) Then
            kind = "SKIP"
            If LCase$(Right$(source, 4)) = ".msg" Or Len(expectedHash) > 0 Then reason = "REVIEW: " & reason
            GoTo Record
        End If
    End If
    If Not MB_ValidDate(stamp) Then Err.Raise MB_ERR + 429, , "No valid archive date."
    parent = MBY_MonthDirectory(accountRoot, stamp)
    target = MBM_Destination(source, parent, reserved)
    Set f = gFSO.GetFile(source)
    bytes = Format$(CDbl(f.Size), "0")
    If StrComp(target, source, vbTextCompare) = 0 Then
        If Len(expectedHash) > 0 And Not gAutoMaintenance Then
            If MBH_HashFile(source) <> expectedHash Then
                Err.Raise MB_ERR + 430, , "Previously indexed file hash does not match."
            End If
        End If
        kind = "KEEP": reason = "Already in the correct year/month folder."
        GoTo Record
    End If
    beforeDate = CDate(f.DateLastModified)
    digest = MBH_HashFile(source)
    If Len(expectedHash) > 0 Then
        If digest <> expectedHash Then Err.Raise MB_ERR + 430, , "Previously indexed file hash does not match."
    End If
    Set f = gFSO.GetFile(source)
    If Format$(CDbl(f.Size), "0") <> bytes Or CDate(f.DateLastModified) <> beforeDate Then
        Err.Raise MB_ERR + 431, , "File changed while the organization plan was being built."
    End If
    kind = "MOVE": reason = "Same-volume move only; preserve original bytes."
    GoTo Record
Review:
    n = Err.Number: detail = Err.Description
    If gFatal Then Err.Raise n, "MBM_PlanItem", detail
    kind = "SKIP": target = "": digest = "": reason = "REVIEW: " & detail
Record:
    If Len(bytes) = 0 Then bytes = "0"
    MB_WriteCheckedLine output, MB_EncodeFields(Array("MBPLAN3", planId, kind, accountRoot, _
                         sourceRoot, source, target, bytes, digest, MB_Stamp(stamp), reason))
    Select Case kind
        Case "MOVE": gMovePlanned = gMovePlanned + 1
        Case "KEEP": gMoveKept = gMoveKept + 1
        Case Else
            gMoveSkipped = gMoveSkipped + 1
            If Left$(reason, 7) = "REVIEW:" Then gMovePlanReview = gMovePlanReview + 1
    End Select
End Sub

Private Function MBM_BuildPlan(ByVal accountRoot As String, ByVal planPath As String, _
                                Optional ByVal extraFiles As Variant) As Long
    Dim output As Object, pending As Collection, folder As Object, child As Object, f As Object
    Dim reserved As Object, seen As Object, folders As Object, planId As String, current As String
    Dim extra As Variant, path As String, n As Long, detail As String, ordinal As Long
    On Error GoTo Failed
    If Not mbmLoaded Then MBM_LoadMoves
    If gCancel Then Exit Function
    If Not mbmLoaded Then Err.Raise MB_ERR + 408, , "Move history is incomplete."
    accountRoot = MBH_NormalizePath(accountRoot)
    If MB_PathWithin(accountRoot, gRoot) Then Err.Raise MB_ERR + 432, , "Choose the account folder outside the managed state tree."
    MBM_AcquireAccountLock accountRoot
    gMovePlanned = 0: gMoveKept = 0: gMoveSkipped = 0: gMovePlanReview = 0
    planPath = MBH_NormalizePath(planPath)
    If Not MB_PathWithin(planPath, gRunDir) Then Err.Raise MB_ERR + 433, , "Plan must be in this run directory."
    planId = MBH_HashText("MBPLAN3|" & gRunId & "|" & accountRoot & "|" & planPath)
    Set reserved = CreateObject("Scripting.Dictionary"): reserved.CompareMode = vbTextCompare
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    Set folders = CreateObject("Scripting.Dictionary"): folders.CompareMode = vbTextCompare
    Set pending = New Collection: pending.Add accountRoot
    Set output = gFSO.CreateTextFile(planPath, False, True)
    MB_WriteCheckedLine output, MB_EncodeFields(Array("MBPLAN3", planId, "HEADER", accountRoot, _
                         "", "", "", "0", "", "", "SentOn month; original filename date prefix preferred."))
    ' Known indexed dates/hashes take precedence over filename heuristics.
    If Not IsMissing(extraFiles) Then
        If IsObject(extraFiles) Then
            For Each extra In extraFiles
                If Not MB_ContinueWork() Then GoTo Cancelled
                path = MBH_NormalizePath(CStr(extra(0)))
                If MB_PathWithin(path, accountRoot) Then
                    MBM_PlanItem output, planId, accountRoot, accountRoot, path, reserved, seen, CDate(extra(1)), CStr(extra(2))
                Else
                    If Not MB_PathWithin(path, MBH_JoinPath(gRoot, "messages")) And _
                       Not MB_PathWithin(path, MBH_JoinPath(gRoot, "attachments")) Then
                        Err.Raise MB_ERR + 434, , "Extra organization source is outside committed managed namespaces."
                    End If
                    MBM_PlanItem output, planId, accountRoot, gRoot, path, reserved, seen, CDate(extra(1)), CStr(extra(2))
                End If
            Next extra
        End If
    End If
    Do While pending.Count > 0
        If Not MB_ContinueWork() Then GoTo Cancelled
        current = CStr(pending(pending.Count)): pending.Remove pending.Count
        If Not folders.Exists(current) And Not MB_PathWithin(current, gRoot) Then
            folders.Add current, True
            MBM_CheckAncestors current
            If Not MBL_HasOwner(current) Then
                Set folder = gFSO.GetFolder(current)
                For Each f In folder.Files
                    If Not MB_ContinueWork() Then GoTo Cancelled
                    path = CStr(f.Path)
                    If Left$(LCase$(CStr(f.Name)), 4) <> "_mb_" Then
                        MBM_PlanItem output, planId, accountRoot, accountRoot, path, reserved, seen
                    End If
                Next f
                For Each child In folder.SubFolders
                    pending.Add MBH_NormalizePath(CStr(child.Path))
                Next child
            End If
        End If
    Loop
    ordinal = gMovePlanned + gMoveKept + gMoveSkipped
    MB_WriteCheckedLine output, MB_EncodeFields(Array("MBPLAN3", planId, "END", accountRoot, _
                         "", "", "", CStr(ordinal), "", "", "Plan complete; no files have moved."))
    output.Close: Set output = Nothing
    MBM_FlushPath planPath
    gMovePlanPath = planPath
    MBM_BuildPlan = gMovePlanned
    Exit Function
Cancelled:
    output.Close: Set output = Nothing
    gMovePlanPath = ""
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close
    On Error GoTo 0
    gMovePlanPath = ""
    Err.Raise n, "MBM_BuildPlan", detail
End Function

Private Sub MBM_ValidatePlan(ByVal planPath As String, ByRef planId As String, ByRef accountRoot As String)
    Dim readStream As Object, line As String, fields As Variant, header As Boolean, footer As Boolean
    Dim count As Long, sourceRoot As String, p As String, targets As Object, sources As Object
    Dim n As Long, detail As String
    On Error GoTo Failed
    gMovePlanned = 0: gMoveKept = 0: gMoveSkipped = 0: gMovePlanReview = 0
    planPath = MBH_NormalizePath(planPath)
    If Not MBS_PlanPathAllowed(planPath) Then Err.Raise MB_ERR + 435, , "Plan is not in this archive's run history."
    If CDbl(gFSO.GetFile(planPath).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 436, , "Organization plan is truncated."
    Set targets = CreateObject("Scripting.Dictionary"): targets.CompareMode = vbTextCompare
    Set sources = CreateObject("Scripting.Dictionary"): sources.CompareMode = vbTextCompare
    Set readStream = gFSO.OpenTextFile(planPath, 1, False, -1)
    Do Until readStream.AtEndOfStream
        If Not MB_ContinueWork() Then GoTo Cancelled
        line = readStream.ReadLine
        If Len(line) > 0 Then
            If footer Then Err.Raise MB_ERR + 437, , "Unexpected records after plan END."
            fields = MB_ReadCheckedLine(line, 11)
            If fields(0) <> "MBPLAN3" Then Err.Raise MB_ERR + 438, , "Unknown organization plan schema."
            If Not header Then
                If fields(2) <> "HEADER" Or Not MB_IsHexKey(CStr(fields(1))) Then Err.Raise MB_ERR + 439, , "Invalid plan header."
                planId = CStr(fields(1)): accountRoot = MBH_NormalizePath(CStr(fields(3)))
                If MB_PathWithin(accountRoot, gRoot) Then Err.Raise MB_ERR + 432, , "Invalid account root in plan."
                header = True
            Else
                If fields(1) <> planId Or StrComp(CStr(fields(3)), accountRoot, vbTextCompare) <> 0 Then
                    Err.Raise MB_ERR + 440, , "Plan identity/root changed between rows."
                End If
                If fields(2) = "END" Then
                    If MB_ByteCount(CStr(fields(7))) <> CDbl(count) Then Err.Raise MB_ERR + 441, , "Plan END count mismatch."
                    footer = True
                Else
                    count = count + 1
                    If fields(2) <> "MOVE" And fields(2) <> "KEEP" And fields(2) <> "SKIP" Then
                        Err.Raise MB_ERR + 442, , "Unknown plan action."
                    End If
                    Select Case CStr(fields(2))
                        Case "MOVE": gMovePlanned = gMovePlanned + 1
                        Case "KEEP": gMoveKept = gMoveKept + 1
                        Case "SKIP"
                            gMoveSkipped = gMoveSkipped + 1
                            If Left$(CStr(fields(10)), 7) = "REVIEW:" Then gMovePlanReview = gMovePlanReview + 1
                    End Select
                    If fields(2) <> "SKIP" Then
                        sourceRoot = MBH_NormalizePath(CStr(fields(4)))
                        p = MBM_File(CStr(fields(5)), sourceRoot, False)
                        If sources.Exists(p) Then Err.Raise MB_ERR + 443, , "Duplicate source in organization plan."
                        sources.Add p, True
                        If StrComp(sourceRoot, accountRoot, vbTextCompare) <> 0 Then
                            If StrComp(sourceRoot, gRoot, vbTextCompare) <> 0 Then Err.Raise MB_ERR + 444, , "Unexpected source root."
                            If Not MB_PathWithin(p, MBH_JoinPath(gRoot, "messages")) And _
                               Not MB_PathWithin(p, MBH_JoinPath(gRoot, "attachments")) Then
                                Err.Raise MB_ERR + 434, , "Unexpected managed source namespace."
                            End If
                        End If
                        p = MBM_File(CStr(fields(6)), accountRoot, False)
                        If targets.Exists(p) Then Err.Raise MB_ERR + 445, , "Duplicate target in organization plan."
                        targets.Add p, True
                        If fields(2) = "MOVE" Then
                            If Not MB_IsHexKey(CStr(fields(8))) Then Err.Raise MB_ERR + 446, , "Invalid planned file hash."
                            If StrComp(CStr(fields(5)), CStr(fields(6)), vbTextCompare) = 0 Then Err.Raise MB_ERR + 413, , "MOVE paths are equal."
                        Else
                            If StrComp(CStr(fields(5)), CStr(fields(6)), vbTextCompare) <> 0 Then Err.Raise MB_ERR + 447, , "KEEP paths differ."
                        End If
                        If StrComp(CStr(gFSO.GetParentFolderName(CStr(fields(6)))), _
                           MBY_MonthDirectory(accountRoot, MB_ParseStamp(CStr(fields(9)))), vbTextCompare) <> 0 Then
                            Err.Raise MB_ERR + 448, , "Target does not match its recorded year/month."
                        End If
                        If MB_ByteCount(CStr(fields(7))) < 0 Then Err.Raise MB_ERR + 414, , "Invalid planned file size."
                    End If
                End If
            End If
        End If
    Loop
    readStream.Close: Set readStream = Nothing
    If Not header Or Not footer Then Err.Raise MB_ERR + 449, , "Organization plan is incomplete; no files may move."
    Exit Sub
Cancelled:
    If Not readStream Is Nothing Then readStream.Close
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not readStream Is Nothing Then readStream.Close
    On Error GoTo 0
    Err.Raise n, "MBM_ValidatePlan", detail
End Sub

Private Sub MBM_DoMove(ByVal fields As Variant, ByVal reverse As Boolean)
    Dim source As String, target As String, sourceRoot As String, targetRoot As String
    Dim winError As Long, preparePhase As String, commitPhase As String
    source = CStr(fields(6)): target = CStr(fields(7))
    sourceRoot = CStr(fields(5)): targetRoot = CStr(fields(4))
    preparePhase = "PREPARE": commitPhase = "COMMIT"
    If reverse Then
        source = CStr(fields(7)): target = CStr(fields(6))
        sourceRoot = CStr(fields(4)): targetRoot = CStr(fields(5))
        preparePhase = "UNDO_PREPARE": commitPhase = "UNDO_COMMIT"
    End If
    MBM_VerifyFile source, sourceRoot, CStr(fields(8)), CStr(fields(9))
    If gFSO.FileExists(target) Or gFSO.FolderExists(target) Then
        Err.Raise MB_ERR + 450, , "Destination is occupied; both files are preserved: " & target
    End If
    MB_EnsureDir CStr(gFSO.GetParentFolderName(target))
    MBM_CheckAncestors CStr(gFSO.GetParentFolderName(target))
    MBM_AppendMove fields, preparePhase
    If MBM_MoveFileExW(StrPtr(source), StrPtr(target), 8) = 0 Then
        winError = Err.LastDllError
        ' Keep PREPARE as pending: the next load resolves actual filesystem state.
        Err.Raise MB_ERR + 451, "MBM_DoMove", _
                  "Same-volume no-overwrite move failed. Win32=" & CStr(winError) & "; " & source & " => " & target
    End If
    MBM_VerifyFile target, targetRoot, CStr(fields(8)), CStr(fields(9))
    If gFSO.FileExists(source) Then Err.Raise MB_ERR + 452, , "Source still exists after move; recovery required."
    MBM_AppendMove fields, commitPhase
End Sub

Private Sub MBM_ApplyPlan(ByVal planPath As String)
    Dim planId As String, accountRoot As String, readStream As Object, line As String
    Dim p As Variant, fields As Variant, prior As Variant, key As String
    Dim n As Long, detail As String
    On Error GoTo Failed
    gMoveApplied = 0: gMoveAlready = 0
    If Not mbmLoaded Then MBM_LoadMoves
    If gCancel Then Exit Sub
    If Not mbmLoaded Then Err.Raise MB_ERR + 408, , "Move history is incomplete."
    MBM_ValidatePlan planPath, planId, accountRoot
    If gCancel Then Exit Sub
    MBM_AcquireAccountLock accountRoot
    Set readStream = gFSO.OpenTextFile(planPath, 1, False, -1)
    Do Until readStream.AtEndOfStream
        If Not MB_ContinueWork() Then Exit Do
        line = readStream.ReadLine
        If Len(line) > 0 Then
            p = MB_ReadCheckedLine(line, 11)
            If p(2) = "MOVE" Then
                key = MBM_OperationId(planId, CStr(p(5)), CStr(p(6)), CStr(p(8)))
                fields = Array("MBMOVE3", key, planId, "PREPARE", accountRoot, CStr(p(4)), _
                               CStr(p(5)), CStr(p(6)), CStr(p(7)), CStr(p(8)), MB_Stamp(Now))
                If mbmOps.Exists(key) Then
                    prior = MBM_Unpack(CStr(mbmOps(key)), 11)
                    If prior(3) = "COMMIT" Or prior(3) = "UNDO_ABORT" Then
                        MBM_VerifyFile CStr(p(6)), accountRoot, CStr(p(7)), CStr(p(8))
                        gMoveAlready = gMoveAlready + 1
                        GoTo NextRow
                    End If
                End If
                MBM_DoMove fields, False
                gMoveApplied = gMoveApplied + 1
            End If
        End If
NextRow:
    Loop
    readStream.Close: Set readStream = Nothing
    gMovePlanPath = planPath
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not readStream Is Nothing Then readStream.Close
    On Error GoTo 0
    gMoveReview = gMoveReview + 1
    gFatal = True
    gFatalText = "Organization stopped; preserve the plan and move journal. " & detail
    Err.Raise n, "MBM_ApplyPlan", gFatalText
End Sub

Private Sub MBM_UndoPlan(ByVal planPath As String)
    Dim planId As String, accountRoot As String, readStream As Object, line As String
    Dim p As Variant, fields As Variant, key As String, n As Long, detail As String
    On Error GoTo Failed
    gMoveRestored = 0: gMoveAlready = 0
    If Not mbmLoaded Then MBM_LoadMoves
    If gCancel Then Exit Sub
    If Not mbmLoaded Then Err.Raise MB_ERR + 408, , "Move history is incomplete."
    MBM_ValidatePlan planPath, planId, accountRoot
    If gCancel Then Exit Sub
    MBM_AcquireAccountLock accountRoot
    ' Every planned target was unoccupied and reserved; no target is a planned
    ' source. Therefore forward plan order is also safe for restoring the files.
    Set readStream = gFSO.OpenTextFile(planPath, 1, False, -1)
    Do Until readStream.AtEndOfStream
        If Not MB_ContinueWork() Then Exit Do
        line = readStream.ReadLine
        If Len(line) > 0 Then
            p = MB_ReadCheckedLine(line, 11)
            If p(2) = "MOVE" Then
                key = MBM_OperationId(planId, CStr(p(5)), CStr(p(6)), CStr(p(8)))
                If mbmOps.Exists(key) Then
                    fields = MBM_Unpack(CStr(mbmOps(key)), 11)
                    If fields(3) = "COMMIT" Or fields(3) = "UNDO_ABORT" Then
                        MBM_DoMove fields, True
                        gMoveRestored = gMoveRestored + 1
                    ElseIf fields(3) = "UNDO_COMMIT" Then
                        MBM_VerifyFile CStr(fields(6)), CStr(fields(5)), CStr(fields(8)), CStr(fields(9))
                        gMoveAlready = gMoveAlready + 1
                    End If
                End If
            End If
        End If
    Loop
    readStream.Close: Set readStream = Nothing
    gMovePlanPath = planPath
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not readStream Is Nothing Then readStream.Close
    On Error GoTo 0
    gMoveReview = gMoveReview + 1
    gFatal = True
    gFatalText = "Undo stopped; occupied/changed originals are preserved. " & detail
    Err.Raise n, "MBM_UndoPlan", gFatalText
End Sub


Private Function MBM_PlanAccountRoot(ByVal planPath As String) As String
    Dim planId As String, root As String
    MBM_ValidatePlan planPath, planId, root
    If Not gCancel Then MBM_PlanAccountRoot = root
End Function

Private Sub MBM_RequirePlanAccount(ByVal planPath As String, ByVal expectedRoot As String)
    Dim root As String
    root = MBM_PlanAccountRoot(planPath)
    If gCancel Then Exit Sub
    If StrComp(root, MBH_NormalizePath(expectedRoot), vbTextCompare) <> 0 Then
        Err.Raise MB_ERR + 453, , "The saved plan belongs to a different account folder. Create a new preview."
    End If
End Sub

Private Function MBM_Summary() As String
    MBM_Summary = "Year/month files: planned moves=" & CStr(gMovePlanned) & _
                  "; already organized=" & CStr(gMoveKept) & _
                  "; unchanged=" & CStr(gMoveSkipped) & "; plan review=" & CStr(gMovePlanReview) & vbCrLf & _
                  "Moved=" & CStr(gMoveApplied) & "; prior moves verified=" & CStr(gMoveAlready) & _
                  "; restored=" & CStr(gMoveRestored) & "; recovered=" & CStr(gMoveRecovered) & _
                  "; move errors=" & CStr(gMoveReview)
    If Len(gMovePlanPath) > 0 Then MBM_Summary = MBM_Summary & vbCrLf & "Plan: " & gMovePlanPath
End Function


' ========================= Year/month integration ==========================
Private Sub MB_PreviewYearMonth()
    MB_Run "PREVIEW_YEAR_MONTH", False, False, False
End Sub

Private Sub MB_OrganizeYearMonth()
    MB_Run "ORGANIZE_YEAR_MONTH", False, False, False
End Sub

Private Sub MB_ApplyYearMonthPlan()
    MB_Run "APPLY_YEAR_MONTH", False, False, False
End Sub

Private Sub MB_UndoYearMonth()
    MB_Run "UNDO_YEAR_MONTH", False, False, False
End Sub

Private Function MBX_IsFileMode(ByVal mode As String) As Boolean
    MBX_IsFileMode = (mode = "PREVIEW_YEAR_MONTH" Or mode = "ORGANIZE_YEAR_MONTH" Or _
                     mode = "APPLY_YEAR_MONTH" Or mode = "UNDO_YEAR_MONTH" Or mode = "CHOOSE_ACCOUNT")
End Function

Private Function MBX_IndexKey(ByVal kind As String, ByVal key As String) As String
    If kind = "N" Or kind = "R" Then kind = "M"
    If kind = "T" Then kind = "A"
    MBX_IndexKey = kind & "|" & key
End Function

Private Function MBX_StoreToken(ByVal storeId As String) As String
    If Not gStoreTokens.Exists(storeId) Then gStoreTokens.Add storeId, CStr(gStoreTokens.Count + 1)
    MBX_StoreToken = CStr(gStoreTokens(storeId))
End Function

Private Sub MBX_CacheRecord(ByVal f As Variant)
    Dim key As String, kind As String, p As String, bytes As Double
    Dim reference As String, referenceFields As Variant, referenceMeta As Variant
    f = MBX_ExpandIndexFields(f)
    kind = CStr(f(1)): key = MBX_IndexKey(kind, CStr(f(2)))
    bytes = MB_ByteCount(CStr(f(4)))
    If (kind = "M" Or kind = "N" Or kind = "L" Or kind = "R") And bytes <= 0 Then
        Err.Raise MB_ERR + 600, , "Empty MSG record."
    End If
    If MB_SourceKey(CStr(f(7)), CStr(f(8))) <> CStr(f(6)) Then
        Err.Raise MB_ERR + 601, , "Index source identity mismatch."
    End If
    ' Every reader/writer shares the immutable expected-byte contract. Normal
    ' journal replay must not silently replace proof that recovery would reject.
    MBR_RejectIndexConflict f
    Select Case kind
        Case "N", "T", "R"
            If UBound(f) <> 16 Or CStr(f(0)) <> "MBS3" Then Err.Raise MB_ERR + 602, , "Missing account root binding."
            p = MBY_CanonicalFilePath(CStr(f(3)), CStr(f(16)))
            If (kind = "N" Or kind = "R") And LCase$(Right$(p, 4)) <> ".msg" Then Err.Raise MB_ERR + 603, , "Message path is not MSG."
            If kind = "R" Then
                If CStr(f(13)) <> "0" Or (CStr(f(15)) <> MB_NEW_CONTENT_VALIDATION And _
                   CStr(f(15)) <> MB_NEW_CONTENT_VALIDATION_V2) Then _
                    Err.Raise MB_ERR + 606, , "Invalid restored-identity validation."
                If (Left$(CStr(f(14)), 2) <> "M|" And Left$(CStr(f(14)), 2) <> "L|") Or _
                   Not MB_IsHexKey(Mid$(CStr(f(14)), 3)) Then _
                    Err.Raise MB_ERR + 607, , "Invalid restored-identity source reference."
                reference = CStr(f(14))
                If reference = key Or Not gIndex.Exists(reference) Or Not gRecordMeta.Exists(reference) Then _
                    Err.Raise MB_ERR + 608, , "Restored-identity reference has no prior committed proof."
                referenceFields = MB_IndexFields(reference): referenceMeta = MBX_Meta(reference)
                If CStr(referenceMeta(0)) <> MBX_StoreToken(CStr(f(7))) Or _
                   CStr(referenceFields(2)) <> CStr(f(5)) Or MB_ByteCount(CStr(referenceFields(1))) <> bytes Then _
                    Err.Raise MB_ERR + 609, , "Restored-identity reference account or bytes disagree."
            End If
            gExternalRoots(key) = MBH_NormalizePath(CStr(f(16)))
            gIndex(key) = MB_EncodeFields(Array(p, f(4), f(5), f(15)))
        Case "L"
            MB_LoadLegacyRow f
        Case "M", "A"
            p = MB_RelPath(CStr(f(3)))
            If gExternalRoots.Exists(key) Then gExternalRoots.Remove key
            gIndex(key) = MB_EncodeFields(Array(f(3), f(4), f(5), f(15)))
        Case Else
            Err.Raise MB_ERR + 604, , "Unknown index file kind."
    End Select
    ' Store the account once, not a repeated long StoreID in every memory entry.
    gRecordMeta(key) = MB_EncodeFields(Array(MBX_StoreToken(CStr(f(7))), CStr(f(11)), CStr(f(6))))
    MBV_CacheRecord f
    MBI_CacheRecord f
    MBW_CacheRecord f
End Sub

Private Function MBX_Meta(ByVal key As String) As Variant
    Dim f As Variant, i As Long
    f = Split(CStr(gRecordMeta(key)), vbTab)
    If UBound(f) <> 2 Then Err.Raise MB_ERR + 605, , "Invalid index metadata."
    For i = 0 To 2
        f(i) = MB_Unescape(CStr(f(i)))
    Next i
    MBX_Meta = f
End Function

Private Function MBX_IndexedRoot(ByVal key As String) As String
    Dim binding As Variant
    If Left$(key, 2) = "L|" Then
        binding = MB_LegacyBinding(key)
        MBX_IndexedRoot = CStr(binding(0))
    ElseIf gExternalRoots.Exists(key) Then
        MBX_IndexedRoot = CStr(gExternalRoots(key))
    Else
        MBX_IndexedRoot = gRoot
    End If
End Function

Private Function MBX_IndexedOriginal(ByVal key As String, ByVal fields As Variant) As String
    If Left$(key, 2) = "L|" Or gExternalRoots.Exists(key) Then
        MBX_IndexedOriginal = MBY_CanonicalFilePath(CStr(fields(0)), MBX_IndexedRoot(key))
    Else
        MBX_IndexedOriginal = MB_RelPath(CStr(fields(0)))
    End If
End Function

Private Function MBX_IndexedPath(ByVal key As String, ByVal fields As Variant, _
                                 Optional ByVal requireExisting As Boolean = False) As String
    Dim original As String, resolved As String, bound As String
    original = MBX_IndexedOriginal(key, fields)
    bound = MBM_ResolvedRoot(original, CStr(fields(2)), MBX_IndexedRoot(key))
    resolved = MBM_ResolvePath(original, CStr(fields(2)))
    MBX_IndexedPath = MBY_CanonicalFilePath(resolved, bound, requireExisting)
End Function

Private Sub MBX_BuildManagedPaths()
    Dim key As Variant, f As Variant, p As String, binding As Variant, owner As String, ownerKey As String
    gManagedFilePaths.RemoveAll
    For Each key In gIndex.Keys
        If Not MB_ContinueWork() Then Exit Sub
        f = MB_IndexFields(CStr(key))
        p = MBX_IndexedPath(CStr(key), f)
        If Left$(CStr(key), 2) = "M|" Or Left$(CStr(key), 2) = "A|" Then
            gManagedFilePaths(p) = True
        ElseIf Left$(CStr(key), 2) = "L|" Then
            binding = MB_LegacyBinding(CStr(key)): owner = CStr(binding(1))
            ownerKey = MB_LegacyOwnerKey(p, CStr(f(2)))
            If gLegacyOwners.Exists(ownerKey) Then
                If CStr(gLegacyOwners(ownerKey)) <> owner Then
                    Err.Raise MB_ERR + 610, , "Moved legacy file is assigned to conflicting source identities."
                End If
            End If
            gLegacyOwners(ownerKey) = owner
        End If
    Next key
End Sub

Private Function MBX_IsManagedPath(ByVal path As String) As Boolean
    MBX_IsManagedPath = gManagedFilePaths.Exists(path)
End Function

Private Sub MBX_AddAccount(ByVal storeId As String)
    Dim root As String, value As Variant
    root = MBY_AccountRoot(storeId)
    If Len(root) = 0 Then gCancel = True: Exit Sub
    For Each value In gArchiveRoots
        If StrComp(CStr(value), root, vbTextCompare) = 0 Then Exit Sub
    Next value
    gArchiveRoots.Add root
End Sub

Private Sub MBX_ResolveAccounts(ByVal roots As Collection, ByVal selection As Outlook.Selection)
    Dim folder As Outlook.Folder, obj As Object, i As Long
    Set gArchiveRoots = New Collection
    If gMode = "SELECTED" Then
        For i = 1 To selection.Count
            If Not MB_ContinueWork() Then Exit Sub
            Set obj = selection.Item(i)
            If TypeOf obj Is Outlook.MailItem Then
                Set folder = obj.Parent
                MBX_AddAccount folder.StoreID
                If gCancel Then Exit Sub
            End If
        Next i
    Else
        For Each folder In roots
            MBX_AddAccount folder.StoreID
            If gCancel Then Exit Sub
        Next folder
    End If
End Sub

Private Function MBX_ManagedFiles(ByVal storeId As String, ByVal accountRoot As String) As Collection
    Dim output As New Collection, dates As Object, token As String, key As Variant
    Dim meta As Variant, fields As Variant, path As String, dt As Date, reason As String
    Dim parentRevision As String, originalPath As String
    Set dates = CreateObject("Scripting.Dictionary")
    token = MBX_StoreToken(storeId)
    For Each key In gIndex.Keys
        If Not MB_ContinueWork() Then Exit For
        If Left$(CStr(key), 2) = "M|" Or Left$(CStr(key), 2) = "L|" Then
            meta = MBX_Meta(CStr(key))
            If CStr(meta(0)) = token Then
                fields = MB_IndexFields(CStr(key)): path = MBX_IndexedPath(CStr(key), fields)
                If MBY_FileMonth(path, dt, reason) Then
                    dates(Mid$(CStr(key), 3)) = dt
                    If MB_PathWithin(path, accountRoot) Or MB_PathWithin(path, gRoot) Then
                        output.Add Array(path, dt, CStr(fields(2)))
                    End If
                Else
                    gFileExtraReview = gFileExtraReview + 1
                    MB_LogEvent "WARNING", "MANAGED_MONTH_REVIEW", CStr(key), path, 0, reason
                End If
            End If
        End If
    Next key
    For Each key In gIndex.Keys
        If Not MB_ContinueWork() Then Exit For
        If Left$(CStr(key), 2) = "A|" Then
            meta = MBX_Meta(CStr(key))
            If CStr(meta(0)) = token Then
                fields = MB_IndexFields(CStr(key)): path = MBX_IndexedPath(CStr(key), fields)
                If MB_PathWithin(path, accountRoot) Or MB_PathWithin(path, gRoot) Then
                    originalPath = MBX_IndexedOriginal(CStr(key), fields)
                    parentRevision = CStr(gFSO.GetFileName(gFSO.GetParentFolderName(originalPath)))
                    If MB_IsHexKey(parentRevision) And dates.Exists(parentRevision) Then
                        output.Add Array(path, CDate(dates(parentRevision)), CStr(fields(2)))
                    Else
                        gFileExtraReview = gFileExtraReview + 1
                        MB_LogEvent "WARNING", "ATTACHMENT_MONTH_REVIEW", CStr(key), path, 0, _
                                    "No readable parent MSG establishes this attachment's sent month."
                    End If
                End If
            End If
        End If
    Next key
    Set MBX_ManagedFiles = output
End Function

Private Function MBX_PlanKey(ByVal storeId As String, Optional ByVal applied As Boolean = False) As String
    Dim prefix As String
    prefix = "YearMonthPreview_"
    If applied Then prefix = "YearMonthApplied_"
    MBX_PlanKey = prefix & MBH_HashText(LCase$(gBaseRoot) & "|" & LCase$(storeId))
End Function

Private Sub MBX_RunFileMode()
    Dim storeId As String, extra As Collection, count As Long
    storeId = Application.Session.DefaultStore.StoreID
    If gMode = "CHOOSE_ACCOUNT" Then
        gFileAccount = MBY_ChooseAccountRoot(storeId)
        gFileDetail = "Account folder remembered. Future messages use yyyy\mm below this folder."
        Exit Sub
    End If
    gFileAccount = MBY_AccountRoot(storeId)
    If Len(gFileAccount) = 0 Then gCancel = True: Exit Sub
    If gJournalIssues > 0 Then Err.Raise MB_ERR + 606, , "Repair invalid backup index rows before organizing files."
    gScope = gFileAccount
    If gMode = "PREVIEW_YEAR_MONTH" Or gMode = "ORGANIZE_YEAR_MONTH" Then
        gPhase = "Planning year/month paths; original files are unchanged"
        Set extra = MBX_ManagedFiles(storeId, gFileAccount)
        If gCancel Then Exit Sub
        gFilePlan = MBH_JoinPath(gRunDir, "organize_plan.tsv")
        count = MBM_BuildPlan(gFileAccount, gFilePlan, extra)
        If gCancel Then Exit Sub
        MBP_SaveSettingText MBX_PlanKey(storeId), gFilePlan
        gFileDetail = "Completed checked plan; eligible records=" & CStr(count)
    Else
        gFilePlan = MBP_GetSettingText(MBX_PlanKey(storeId, gMode = "UNDO_YEAR_MONTH"), "")
        If Len(gFilePlan) = 0 Then Err.Raise MB_ERR + 607, , "No completed year/month plan. Run MB_PreviewYearMonth first."
    End If
    If gMode = "ORGANIZE_YEAR_MONTH" Or gMode = "APPLY_YEAR_MONTH" Then
        gPhase = "Moving existing files into yyyy\mm; no mail SaveAs"
        MBM_RequirePlanAccount gFilePlan, gFileAccount
        If gCancel Then Exit Sub
        MBP_SaveSettingText MBX_PlanKey(storeId, True), gFilePlan
        MBM_ApplyPlan gFilePlan
        gFileDetail = "Year/month move plan processed. No mail was re-exported."
    ElseIf gMode = "UNDO_YEAR_MONTH" Then
        gPhase = "Restoring the previous file locations without overwriting"
        MBM_RequirePlanAccount gFilePlan, gFileAccount
        If gCancel Then Exit Sub
        MBM_UndoPlan gFilePlan
        gFileDetail = "Undo plan processed; files with conflicts or changed bytes require review."
    End If
End Sub

Private Function MBX_FileSummary() As String
    Dim status As String
    status = "FILE OPERATION FINISHED"
    If gMode = "PREVIEW_YEAR_MONTH" Then status = "PREVIEW FINISHED; NO ORIGINAL FILES MOVED"
    If gMoveReview > 0 Or gMovePlanReview > 0 Or gFileExtraReview > 0 Then status = status & "; ITEMS REQUIRE REVIEW"
    If gCancel Then status = "CANCELLED; KEEP THE PLAN TO RESUME"
    If gFatal Then status = "FAILED: " & gFatalText
    MBX_FileSummary = status & vbCrLf & "Version: " & MB_VERSION & "; mode: " & gMode & vbCrLf & _
                     "Account folder: " & gFileAccount & vbCrLf & gFileDetail & vbCrLf & _
                     "Plan: " & gFilePlan & vbCrLf & "Indexed items needing date review: " & _
                     CStr(gFileExtraReview) & vbCrLf & "Results: " & gRunDir & vbCrLf & _
                     MBM_Summary()
End Function



' Account-relative on-disk file references. Runtime paths remain absolute.
' MBS4 retains one exact account-root binding and stores only the file's path
' below it. It is NOT permission to relocate an archive or rewrite old proof.
' Older MBS2/MBS3 rows are read unchanged; only new commits use this format.
Private Function MBX_AccountPath(ByVal relative As String, ByVal accountRoot As String) As String
    Dim value As String, part As Variant
    If Len(relative) = 0 Or Left$(relative, 1) = "\" Or _
       InStr(relative, "/") > 0 Or InStr(relative, ":") > 0 Then _
        Err.Raise MB_ERR + 1180, , "An account-relative file path is required."
    value = MBH_NormalizePath(accountRoot)
    If MBL_IsFileSystemRoot(value) Then Err.Raise MB_ERR + 1181, , "Account binding cannot be a drive or share root."
    For Each part In Split(relative, "\")
        ' JoinPath rejects empty/dot/dot-dot, device names, ADS and unsafe text.
        value = MBH_JoinPath(value, CStr(part))
    Next part
    MBX_AccountPath = MBY_CanonicalFilePath(value, accountRoot)
End Function

Private Function MBX_RelativeRejected(ByVal relative As String, ByVal accountRoot As String) As Boolean
    Dim ignored As String
    On Error GoTo Rejected
    ignored = MBX_AccountPath(relative, accountRoot)
    Exit Function
Rejected:
    MBX_RelativeRejected = True
End Function

Private Function MBX_RelativeSelfTest(ByVal boundAccountRoot As String, ByRef detail As String) As Boolean
    ' Pure codec/path checks under the caller's owned test directory. No MSG,
    ' directory or metadata file is created, opened, moved or rewritten here.
    Dim root As String, relative As String, absolute As String, payload As String, line As String
    Dim row As Variant, expanded As Variant, kind As Variant, unsafe As Variant
    On Error GoTo Failed
    detail = ""
    root = MBH_NormalizePath(boundAccountRoot)
    relative = "2026\11\" & ChrW$(&HD55C) & ChrW$(&HAE00) & "_%09_" & _
               ChrW$(&HD83D) & ChrW$(&HDE00) & ".msg"
    absolute = MBX_AccountPath(relative, root)
    If MBX_AccountRelative(absolute, root) <> relative Then _
        Err.Raise MB_ERR + 1184, , "Unicode/percent relative path did not round-trip."
    For Each kind In Array("N", "T", "R", "L")
        row = Array("MBS4", CStr(kind), String$(64, "a"), relative, "100", String$(64, "b"), _
                    String$(64, "c"), "synthetic-store", "synthetic-entry", "synthetic-folder", _
                    "20261111_010000", "20261110_235959", "subject%09" & vbTab & "text", "0", "", "policy", root)
        If CStr(kind) = "R" Then row(14) = "M|" & String$(64, "d")
        payload = MB_EncodeFields(row)
        line = payload & vbTab & MBH_HashText(payload) & vbTab & "END"
        row = MBR_ParseIndexFields(line)
        If LBound(row) <> 0 Or UBound(row) <> 16 Then Err.Raise MB_ERR + 1184, , "Relative wire field count changed."
        expanded = MBX_ExpandIndexFields(row)
        If CStr(expanded(3)) <> absolute Then Err.Raise MB_ERR + 1184, , "Expanded path differs from the account binding."
        If CStr(kind) = "L" Then
            If UBound(expanded) <> 15 Or CStr(expanded(0)) <> "MBS2" Or CStr(expanded(14)) <> root Then _
                Err.Raise MB_ERR + 1184, , "Relative legacy expansion changed its 16-field contract."
        Else
            If UBound(expanded) <> 16 Or CStr(expanded(0)) <> "MBS3" Or CStr(expanded(16)) <> root Then _
                Err.Raise MB_ERR + 1184, , "Relative N/T/R expansion changed its 17-field contract."
        End If
        If MB_EncodeFields(row) <> payload Then Err.Raise MB_ERR + 1184, , "Expansion modified the checked wire array."
        If MB_EncodeFields(MBR_ParseIndexFields(line)) <> payload Then _
            Err.Raise MB_ERR + 1184, , "Relative checked line did not remain readable."
    Next kind
    For Each unsafe In Array("", ".", "..", "..\x.msg", "2026\..\x.msg", "2026\.\x.msg", _
                             "2026/11/x.msg", "2026\\x.msg", "file.msg:stream", root & "\x.msg")
        If Not MBX_RelativeRejected(CStr(unsafe), root) Then _
            Err.Raise MB_ERR + 1184, , "Unsafe relative path was accepted: " & CStr(unsafe)
    Next unsafe
    detail = "PASS: relative Unicode/percent paths, checked N/T/R/L wire rows, immutable expansion and unsafe-path rejection."
    MBX_RelativeSelfTest = True
    Exit Function
Failed:
    detail = "FAIL: " & CStr(Err.Number) & " - " & Err.Description
End Function

Private Function MBX_AccountRelative(ByVal path As String, ByVal accountRoot As String) As String
    Dim canonical As String, root As String, relative As String
    root = MBH_NormalizePath(accountRoot)
    canonical = MBY_CanonicalFilePath(path, root)
    If StrComp(canonical, root, vbTextCompare) = 0 Or Not MB_PathWithin(canonical, root) Then _
        Err.Raise MB_ERR + 1181, , "File must be below the bound account root."
    relative = Mid$(canonical, Len(root) + 2)
    If StrComp(MBX_AccountPath(relative, root), canonical, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 1182, , "Account-relative file path did not round-trip."
    MBX_AccountRelative = relative
End Function

Private Function MBX_ExpandIndexFields(ByVal row As Variant) As Variant
    Dim fields() As Variant, i As Long, kind As String, root As String
    If CStr(row(0)) <> "MBS4" Then
        MBX_ExpandIndexFields = row
        Exit Function
    End If
    If LBound(row) <> 0 Or UBound(row) <> 16 Then _
        Err.Raise MB_ERR + 1183, , "Relative journal row requires exactly 17 fields."
    kind = CStr(row(1))
    If kind <> "N" And kind <> "T" And kind <> "R" And kind <> "L" Then _
        Err.Raise MB_ERR + 1183, , "Unknown relative journal record kind."
    root = MBH_NormalizePath(CStr(row(16)))
    ' Copy into a separate array: the checked serialized row remains immutable.
    If kind = "L" Then ReDim fields(0 To 15) Else ReDim fields(0 To 16)
    For i = 0 To UBound(fields)
        fields(i) = row(i)
    Next i
    fields(3) = MBX_AccountPath(CStr(row(3)), root)
    If kind = "L" Then
        If Len(CStr(row(14))) <> 0 Then Err.Raise MB_ERR + 1183, , "Relative legacy row has a duplicate root binding."
        fields(0) = "MBS2": fields(14) = root
    Else
        fields(0) = "MBS3": fields(16) = root
    End If
    MBX_ExpandIndexFields = fields
End Function


' This disposable native Windows test is NOT executed while authoring on Linux.
' It uses only a newly created Temp directory, synthetic TXT files and no settings.
Private Sub MB_YearMonthSelfTest()
    Dim resultDetail As String, passed As Boolean
    If gBusy Then
        MsgBox "Wait until the current job is finished.", vbInformation
        Exit Sub
    End If
    gBusy = True
    passed = MBT_YearMonthSelfTest(resultDetail)
    gBusy = False
    MsgBox resultDetail, IIf(passed, vbInformation, vbExclamation), "Year/month self-test"
End Sub

Private Function MBT_YearMonthSelfTest(ByRef resultDetail As String) As Boolean
    ' The acceptance runner owns gBusy across this test and its sibling tests.
    Dim testBase As String, account As String, topic As String, month As String
    Dim a As String, b As String, occupied As String, destination As String, plan As String
    Dim hashA As String, hashB As String, hashOccupied As String, fields As Variant
    Dim c As String, targetC As String, d As String, targetD As String, planId As String
    Dim reusedTarget As String, originalHash As String, reusedHash As String
    Dim badPlan As String, stream As Object, contents As String, ignoredId As String, ignoredRoot As String
    Dim owned As Boolean, rejected As Boolean, n As Long, detail As String, cleanup As String
    On Error GoTo Failed
    resultDetail = ""
    MB_Reset
    gMode = "SELFTEST_YEAR_MONTH": gPhase = "Testing organization with disposable files"
    gClockStart = MB_Clock(): gRunStart = Now: gUntil = gRunStart
    Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set gDirs = CreateObject("Scripting.Dictionary"): gDirs.CompareMode = vbTextCompare
    testBase = MBH_JoinPath(gFSO.GetSpecialFolder(2).Path, "MBYM_" & gFSO.GetTempName)
    gFSO.CreateFolder testBase
    owned = True
    gRoot = MBH_JoinPath(testBase, MBP_MANAGED_LEAF): gBaseRoot = gRoot
    account = MBH_JoinPath(testBase, "account")
    MB_EnsureDir gRoot: MB_EnsureDir account
    MB_EnsureDir MBH_JoinPath(gRoot, "_state")
    gRunId = "selftest"
    gRunDir = MBH_JoinPath(MBH_JoinPath(gRoot, "runs"), gRunId)
    MB_EnsureDir gRunDir
    Set gEvents = gFSO.CreateTextFile(MBH_JoinPath(gRunDir, "events.tsv"), False, True)
    gEvents.WriteLine "Disposable year/month self-test"
    MBM_LoadMoves
    topic = MBH_JoinPath(account, "partly_organized")
    month = MBY_MonthDirectory(account, DateSerial(2024, 1, 15))
    MB_EnsureDir topic: MB_EnsureDir month
    a = MBH_JoinPath(account, "20240115_120000_sender_A.txt")
    occupied = MBH_JoinPath(month, "20240115_120000_sender_A.txt")
    b = MBH_JoinPath(topic, "20231231_235959_sender_B.txt")
    MBM_TestWrite a, "first source bytes"
    MBM_TestWrite occupied, "different existing destination bytes"
    MBM_TestWrite b, "nested source bytes"
    MBM_TestWrite MBH_JoinPath(account, "notes.bin"), "unrelated file"
    hashA = MBH_HashFile(a): hashB = MBH_HashFile(b): hashOccupied = MBH_HashFile(occupied)
    plan = MBH_JoinPath(gRunDir, "organize_plan.tsv")
    MBM_Assert MBM_BuildPlan(account, plan) = 2, "Preview did not find exactly two moves."
    MBM_Assert gMoveKept = 1 And gMoveSkipped = 1, "Preview KEEP/SKIP totals are wrong."
    MBM_Assert gFSO.FileExists(a) And gFSO.FileExists(b), "Preview moved a source file."
    MBM_ApplyPlan plan
    MBM_Assert gMoveApplied = 2, "Apply did not move exactly two files."
    destination = MBM_ResolvePath(a, hashA)
    MBM_Assert StrComp(destination, occupied, vbTextCompare) <> 0, "Collision did not preserve both names."
    MBM_Assert MBH_HashFile(destination) = hashA, "Moved source bytes changed."
    MBM_Assert MBH_HashFile(occupied) = hashOccupied, "Occupied destination was overwritten."
    MBM_Assert Not gFSO.FileExists(a) And Not gFSO.FileExists(b), "Original paths remained after successful moves."
    MBM_ApplyPlan plan
    MBM_Assert gMoveApplied = 0 And gMoveAlready = 2, "Applying the same plan was not idempotent."
    MBM_UndoPlan plan
    MBM_Assert gMoveRestored = 2, "Undo did not restore exactly two files."
    MBM_Assert MBH_HashFile(a) = hashA And MBH_HashFile(b) = hashB, "Undo bytes do not match originals."
    MBM_Assert MBM_ResolvePath(a, hashA) = a, "Undo retained a forward redirect/cycle."
    MBM_Assert MBM_ResolvePath(destination, hashA) = a, "Undo did not redirect newer references to the restored file."
    MBM_Assert MBM_MoveFileExW(StrPtr(a), StrPtr(occupied), 8) = 0, "Native move overwrote an occupied destination."
    MBM_Assert MBH_HashFile(a) = hashA And MBH_HashFile(occupied) = hashOccupied, "No-overwrite check changed bytes."

    ' Simulate process interruption after a successful rename and before COMMIT.
    c = MBH_JoinPath(account, "20240201_120000_sender_C.txt")
    targetC = MBH_JoinPath(MBY_MonthDirectory(account, DateSerial(2024, 2, 1)), CStr(gFSO.GetFileName(c)))
    MBM_TestWrite c, "recover renamed file"
    MB_EnsureDir CStr(gFSO.GetParentFolderName(targetC))
    planId = MBH_HashText("disposable-recovery-test")
    hashB = MBH_HashFile(c)
    fields = Array("MBMOVE3", MBM_OperationId(planId, c, targetC, hashB), planId, "PREPARE", _
                   account, account, c, targetC, Format$(CDbl(gFSO.GetFile(c).Size), "0"), hashB, MB_Stamp(Now))
    MBM_AppendMove fields, "PREPARE"
    MBM_Assert MBM_MoveFileExW(StrPtr(c), StrPtr(targetC), 8) <> 0, "Could not simulate interrupted rename."
    MBM_LoadMoves
    MBM_Assert MBM_ResolvePath(c, hashB) = targetC, "Destination-only recovery did not commit the redirect."
    MBM_Assert MBH_HashFile(targetC) = hashB, "Recovered destination bytes changed."

    ' A different file can later occupy a formerly moved source pathname.
    ' Its movement must retain both files' hash-bound historical references.
    originalHash = hashB
    MBM_TestWrite c, "different bytes at the reused original pathname"
    reusedHash = MBH_HashFile(c)
    reusedTarget = MBH_JoinPath(gFSO.GetParentFolderName(targetC), "20240201_120000_sender_C_reused.txt")
    fields = Array("MBMOVE3", MBM_OperationId(planId, c, reusedTarget, reusedHash), planId, "PREPARE", _
                   account, account, c, reusedTarget, Format$(CDbl(gFSO.GetFile(c).Size), "0"), reusedHash, MB_Stamp(Now))
    MBM_DoMove fields, False
    MBM_Assert MBM_ResolvePath(c, originalHash) = targetC, "Reused pathname erased the original file's redirect."
    MBM_Assert MBM_ResolvePath(c, reusedHash) = reusedTarget, "Reused pathname did not retain the newer file's redirect."
    MBM_LoadMoves
    MBM_Assert MBM_ResolvePath(c, originalHash) = targetC, "Reload lost the older content version's redirect."
    MBM_Assert MBM_ResolvePath(c, reusedHash) = reusedTarget, "Reload lost the newer content version's redirect."
    MBM_DoMove fields, True
    MBM_Assert MBM_ResolvePath(c, originalHash) = targetC, "Undo erased a different content version's redirect."
    MBM_Assert MBM_ResolvePath(c, reusedHash) = c, "Undo failed to restore the reused-path file reference."
    MBM_Assert MBH_HashFile(targetC) = originalHash And MBH_HashFile(c) = reusedHash, _
                "Reused-path move/undo changed either file's bytes."

    ' Simulate interruption after PREPARE while the original file is still present.
    d = MBH_JoinPath(account, "20240301_120000_sender_D.txt")
    targetD = MBH_JoinPath(MBY_MonthDirectory(account, DateSerial(2024, 3, 1)), CStr(gFSO.GetFileName(d)))
    MBM_TestWrite d, "recover unmoved file"
    hashB = MBH_HashFile(d)
    fields = Array("MBMOVE3", MBM_OperationId(planId, d, targetD, hashB), planId, "PREPARE", _
                   account, account, d, targetD, Format$(CDbl(gFSO.GetFile(d).Size), "0"), hashB, MB_Stamp(Now))
    MBM_AppendMove fields, "PREPARE"
    MBM_LoadMoves
    MBM_Assert MBM_ResolvePath(d, hashB) = d And gFSO.FileExists(d), "Source-only recovery moved the file."

    badPlan = MBH_JoinPath(gRunDir, "incomplete_plan.tsv")
    Set stream = gFSO.OpenTextFile(plan, 1, False, -1): contents = stream.ReadAll: stream.Close
    Set stream = gFSO.CreateTextFile(badPlan, False, True)
    stream.Write Left$(contents, Len(contents) - 12): stream.Close: Set stream = Nothing
    On Error Resume Next
    Err.Clear
    MBM_ValidatePlan badPlan, ignoredId, ignoredRoot
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    MBM_Assert rejected, "A truncated organization plan was accepted."
    GoTo CleanUp
Failed:
    n = Err.Number: detail = Err.Description
CleanUp:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Not gEvents Is Nothing Then gEvents.Close
    Set gEvents = Nothing
    Err.Clear
    MBM_ReleaseAccountLocks
    If Err.Number <> 0 Then cleanup = " Account lock cleanup: " & Err.Description
    Err.Clear
    MBH_ReleaseCrypto
    If Err.Number <> 0 Then cleanup = cleanup & " Crypto cleanup: " & Err.Description
    If owned Then
        Err.Clear
        gFSO.DeleteFolder testBase, True
        If Err.Number <> 0 Then cleanup = cleanup & " Temp directory preserved: " & testBase & ": " & Err.Description
    End If
    MBM_Reset
    Set gDirs = Nothing: Set gFSO = Nothing
    gRoot = "": gBaseRoot = "": gRunDir = "": gRunId = ""
    On Error GoTo 0
    If n <> 0 Or Len(cleanup) > 0 Then
        resultDetail = "Year/month self-test FAILED: " & detail & cleanup
    Else
        resultDetail = "Year/month native file self-test PASSED." & vbCrLf & _
                       "Preview, collisions, apply/reapply, undo, interruption recovery, reused source paths, " & _
                       "truncated plan and no-overwrite were checked." & vbCrLf & _
                       "Synthetic TXT files only; Outlook provider behavior was not tested."
        MBT_YearMonthSelfTest = True
    End If
End Function

Private Sub MBM_TestWrite(ByVal path As String, ByVal contents As String)
    Dim stream As Object
    Set stream = gFSO.CreateTextFile(path, False, False)
    stream.Write contents
    stream.Close
End Sub

Private Sub MBM_Assert(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise MB_ERR + 454, "MB_YearMonthSelfTest", detail
End Sub



' Queue discovery. GetTable is dynamic: this fixes a work list, not an atomic
' mailbox snapshot. The export stage must recheck each item's current version.
Private Function MBP_BuildQueue(ByVal roots As Collection, ByVal sinceDate As Date, _
    ByVal untilDate As Date, ByVal includeChildren As Boolean, ByVal queuePath As String) As Long

    Dim fso As Object, output As Object, seenFolders As Object, seenItems As Object
    Dim badFolders As Object, pending As Collection, folder As Object, root As Variant
    Dim storeId As String, entryId As String, folderPath As String, folderKey As String
    Dim reason As String, errorNumber As Long, queued As Long, allowChildren As Boolean
    Dim savedNumber As Long, savedDescription As String, savedSource As String
    Dim closeNumber As Long, closeDescription As String

    On Error GoTo Fatal
    If roots Is Nothing Then Err.Raise vbObjectError + 2801, "MBP_BuildQueue", "No source roots."
    If sinceDate <> 0 And sinceDate > untilDate Then _
        Err.Raise vbObjectError + 2802, "MBP_BuildQueue", "The date interval is reversed."

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set output = fso.CreateTextFile(queuePath, False, True)
    Set seenFolders = CreateObject("Scripting.Dictionary")
    Set seenItems = CreateObject("Scripting.Dictionary")
    Set badFolders = CreateObject("Scripting.Dictionary")
    seenFolders.CompareMode = vbBinaryCompare
    seenItems.CompareMode = vbBinaryCompare
    badFolders.CompareMode = vbBinaryCompare
    Set pending = New Collection
    For Each root In roots
        pending.Add root
    Next root

    Do While pending.Count > 0
        If Not MB_ContinueWork() Then Exit Do
        Set folder = pending(pending.Count)
        pending.Remove pending.Count
        allowChildren = True
        storeId = vbNullString
        entryId = vbNullString
        folderPath = "<unavailable folder>"
        folderKey = vbNullString

        If MBQ_FolderInfo(folder, storeId, entryId, folderPath, errorNumber, reason) Then
            folderKey = storeId & "|" & entryId
            If seenFolders.Exists(folderKey) Then
                allowChildren = False
            Else
                seenFolders.Add folderKey, True
                gFoldersScanned = gFoldersScanned + 1
                If Not MBQ_TablePass(folder, storeId, folderPath, sinceDate, untilDate, _
                    output, seenItems, queued, errorNumber, reason) Then
                    MB_LogEvent "WARN", "SCAN_TABLE_FALLBACK", vbNullString, folderPath, _
                        errorNumber, reason & "; using Items enumeration with the same exact date policy."
                    If MB_ContinueWork() Then
                        MBQ_ItemsPass folder, storeId, folderPath, folderKey, sinceDate, untilDate, _
                            output, seenItems, badFolders, queued
                    End If
                End If
            End If
        Else
            gFoldersScanned = gFoldersScanned + 1
            MBQ_FolderFailed badFolders, folderPath, folderPath, "SCAN_FOLDER_IDENTITY", errorNumber, reason
        End If

        If includeChildren And allowChildren Then
            If MB_ContinueWork() Then MBQ_AddChildren folder, pending, badFolders, folderKey, folderPath
        End If
        Set folder = Nothing
    Loop

    output.Close
    Set output = Nothing
    MBP_BuildQueue = queued
    Exit Function

Fatal:
    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    ' Cleanup is deliberately narrow. Preserve and propagate every write/log error.
    On Error Resume Next
    Err.Clear
    If Not output Is Nothing Then output.Close
    closeNumber = Err.Number
    closeDescription = Err.Description
    Set output = Nothing
    On Error GoTo 0
    If closeNumber <> 0 Then savedDescription = savedDescription & _
        "; queue close also failed: " & CStr(closeNumber) & " " & closeDescription
    Err.Raise savedNumber, "MBP_BuildQueue/" & savedSource, savedDescription
End Function

Private Function MBQ_FolderInfo(ByVal folder As Object, ByRef storeId As String, _
    ByRef entryId As String, ByRef folderPath As String, ByRef errorNumber As Long, _
    ByRef reason As String) As Boolean
    On Error GoTo Failed
    If folder Is Nothing Then Err.Raise vbObjectError + 2810, , "A source folder is Nothing."
    folderPath = CStr(folder.FolderPath)
    storeId = CStr(folder.StoreID)
    entryId = CStr(folder.EntryID)
    If Len(storeId) = 0 Or Len(entryId) = 0 Then _
        Err.Raise vbObjectError + 2811, , "A folder has no StoreID or EntryID."
    errorNumber = 0
    reason = vbNullString
    MBQ_FolderInfo = True
    Exit Function
Failed:
    errorNumber = Err.Number
    reason = Err.Description
End Function

Private Sub MBQ_AddChildren(ByVal folder As Object, ByVal pending As Collection, _
    ByVal badFolders As Object, ByVal folderKey As String, ByVal folderPath As String)
    Dim children As Object, child As Object
    Dim stage As String, n As Long, detail As String
    On Error GoTo Failed
    stage = "SOURCE"
    Set children = folder.Folders
    For Each child In children
        stage = "CONTROL"
        If Not MB_ContinueWork() Then Exit Sub
        pending.Add child
        stage = "SOURCE"
    Next child
    Exit Sub
Failed:
    n = Err.Number
    detail = Err.Description
    If stage <> "SOURCE" Then Err.Raise n, "MBQ_AddChildren", detail
    MBQ_FolderFailed badFolders, folderKey, folderPath, "SCAN_CHILD_FOLDERS", n, detail
End Sub

Private Sub MBQ_FolderFailed(ByVal badFolders As Object, ByVal folderKey As String, _
    ByVal folderPath As String, ByVal stage As String, ByVal errorNumber As Long, ByVal detail As String)
    Dim key As String
    key = folderKey
    If Len(key) = 0 Then key = folderPath
    If Not badFolders.Exists(key) Then
        badFolders.Add key, True
        gFolderFailures = gFolderFailures + 1
    End If
    MB_LogEvent "ERROR", stage, vbNullString, folderPath, errorNumber, detail
End Sub

Private Function MBQ_DateFilter(ByVal sinceDate As Date, ByVal untilDate As Date) As String
    Dim fromMinute As Date, untilMinute As Date, low As String, high As String
    If sinceDate = 0 Or gMode = "AUTO" Then Exit Function
    fromMinute = DateSerial(Year(sinceDate), Month(sinceDate), Day(sinceDate)) _
        + TimeSerial(Hour(sinceDate), Minute(sinceDate), 0)
    untilMinute = DateAdd("n", 1, DateSerial(Year(untilDate), Month(untilDate), Day(untilDate)) _
        + TimeSerial(Hour(untilDate), Minute(untilDate), 0))
    low = Format$(fromMinute, "ddddd h:nn AMPM")
    high = Format$(untilMinute, "ddddd h:nn AMPM")
    If gModifiedWindow Then
        MBQ_DateFilter = "([LastModificationTime] >= '" & low & _
            "' AND [LastModificationTime] < '" & high & "')"
        Exit Function
    End If
    ' Jet uses local time. Deliberately broaden to whole minutes, then compare
    ' exact VBA Date values below. CreationTime is only a candidate-search aid.
    MBQ_DateFilter = "(([ReceivedTime] >= '" & low & "' AND [ReceivedTime] < '" & high & "')" _
        & " OR ([SentOn] >= '" & low & "' AND [SentOn] < '" & high & "')" _
        & " OR ([CreationTime] >= '" & low & "' AND [CreationTime] < '" & high & "'))"
End Function

Private Function MBQ_TablePass(ByVal folder As Object, ByVal storeId As String, _
    ByVal folderPath As String, ByVal sinceDate As Date, ByVal untilDate As Date, _
    ByVal output As Object, ByVal seenItems As Object, ByRef queued As Long, _
    ByRef errorNumber As Long, ByRef reason As String) As Boolean

    Dim tbl As Object, values As Variant, rowIndex As Long, rowCount As Long
    Dim countBefore As Long, countAfter As Long, rowsRead As Long
    Dim stage As String, enumerating As Boolean, needsFallback As Boolean
    Dim n As Long, detail As String, filter As String
    On Error GoTo Failed
    stage = "CONTROL"
    filter = MBQ_DateFilter(sinceDate, untilDate)
    stage = "SOURCE"
    Set tbl = folder.GetTable(filter, 0) ' olUserItems: normal, non-hidden items.
    tbl.Columns.RemoveAll
    tbl.Columns.Add "EntryID"
    tbl.Columns.Add "MessageClass"
    tbl.Columns.Add "LastModificationTime"
    tbl.Columns.Add "Size"
    tbl.Columns.Add "ReceivedTime"
    tbl.Columns.Add "SentOn"
    tbl.Columns.Add "CreationTime"
    countBefore = tbl.GetRowCount
    enumerating = True

    Do
        stage = "CONTROL"
        If Not MB_ContinueWork() Then
            MBQ_TablePass = True
            Exit Function
        End If
        stage = "SOURCE"
        If tbl.EndOfTable Then Exit Do
        values = tbl.GetArray(256)
        rowCount = MBQ_ArrayRowCount(values)
        If rowCount = 0 Then
            If tbl.EndOfTable Then Exit Do
            Err.Raise vbObjectError + 2820, , "GetArray returned no rows before EndOfTable."
        End If
        rowsRead = rowsRead + rowCount
        For rowIndex = 0 To rowCount - 1
            stage = "CONTROL"
            If Not MB_ContinueWork() Then
                MBQ_TablePass = True
                Exit Function
            End If
            stage = "RECORD"
            MBQ_TableRow values, rowIndex, storeId, folderPath, sinceDate, untilDate, _
                output, seenItems, queued, needsFallback
        Next rowIndex
    Loop

    stage = "SOURCE"
    countAfter = tbl.GetRowCount
    stage = "CONTROL"
    If countBefore <> countAfter Or rowsRead <> countAfter Then
        MBQ_Changed vbNullString, folderPath, "SCAN_TABLE_COUNT_CHANGED", 0, _
            "Table counts differ: before=" & CStr(countBefore) & "; read=" & CStr(rowsRead) _
            & "; after=" & CStr(countAfter) & ". This run is not an atomic snapshot."
        needsFallback = True
    End If

    If needsFallback Then
        errorNumber = 0
        reason = "At least one table row could not be resolved; retrying the folder through Items."
    Else
        MBQ_TablePass = True
    End If
    Exit Function
Failed:
    n = Err.Number
    detail = Err.Description
    If stage <> "SOURCE" Then Err.Raise n, "MBQ_TablePass", detail
    If enumerating Then MBQ_Changed vbNullString, folderPath, "SCAN_TABLE_UNSTABLE", n, detail
    errorNumber = n
    reason = detail
End Function

Private Function MBQ_ArrayRowCount(ByRef values As Variant) As Long
    On Error GoTo EmptyArray
    If Not IsArray(values) Then Exit Function
    MBQ_ArrayRowCount = UBound(values, 2) - LBound(values, 2) + 1
    Exit Function
EmptyArray:
    MBQ_ArrayRowCount = 0
End Function

Private Function MBQ_IsMailClass(ByVal messageClass As String) As Boolean
    Dim value As String
    value = LCase$(messageClass)
    MBQ_IsMailClass = (value = "ipm.note" Or Left$(value, 9) = "ipm.note.")
End Function

Private Function MBQ_Text(ByVal value As Variant, ByRef result As String) As Boolean
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    result = CStr(value)
    MBQ_Text = (Len(result) > 0)
    Exit Function
InvalidValue:
    MBQ_Text = False
End Function

Private Function MBQ_Date(ByVal value As Variant, ByRef result As Date) As Boolean
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) <> vbDate Then Exit Function
    result = CDate(value)
    ' Reject empty VBA dates and Outlook's far-future no-date sentinel.
    MBQ_Date = MB_ValidDate(result)
    Exit Function
InvalidValue:
    MBQ_Date = False
End Function

Private Function MBQ_Size(ByVal value As Variant, ByRef result As Long) As Boolean
    Dim number As Double
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    number = CDbl(value)
    If number < 0 Or number > 2147483647# Or number <> Fix(number) Then Exit Function
    result = CLng(number)
    MBQ_Size = True
    Exit Function
InvalidValue:
    MBQ_Size = False
End Function

Private Sub MBQ_TableRow(ByRef values As Variant, ByVal rowIndex As Long, ByVal storeId As String, _
    ByVal folderPath As String, ByVal sinceDate As Date, ByVal untilDate As Date, _
    ByVal output As Object, ByVal seenItems As Object, ByRef queued As Long, ByRef needsFallback As Boolean)

    Dim entryId As String, messageClass As String, key As String, reason As String
    Dim observedMod As Date, effectiveDate As Date, sourceSize As Long
    Dim parseStatus As Long, item As Object, n As Long, detail As String

    parseStatus = MBQ_ParseTableRow(values, rowIndex, entryId, messageClass, observedMod, _
        sourceSize, effectiveDate, reason)
    If Len(entryId) = 0 Then
        MBQ_Changed vbNullString, folderPath, "SCAN_METADATA", 0, reason
        needsFallback = True
        Exit Sub
    End If
    key = MB_SourceKey(storeId, entryId)
    If seenItems.Exists(key) Then
        If CStr(seenItems(key)) <> "PENDING" Then Exit Sub
    End If

    If parseStatus = 0 Then
        If seenItems.Exists(key) Then
            seenItems(key) = "EXCLUDED"
        Else
            seenItems.Add key, "EXCLUDED"
        End If
        gPolicySkipped = gPolicySkipped + 1
        Exit Sub
    End If

    If MBQ_IsMailClass(messageClass) Then
        If Not seenItems.Exists(key) Then
            seenItems.Add key, "PENDING"
            gCandidates = gCandidates + 1
        End If
    End If

    If parseStatus < 0 Then
        MBQ_Changed key, folderPath, "SCAN_METADATA", 0, reason & "; trying the full item."
        If MBQ_GetItem(entryId, storeId, item, n, detail) Then
            If Not MBQ_ItemRow(item, storeId, folderPath, sinceDate, untilDate, output, seenItems, queued, key) Then _
                needsFallback = True
        Else
            MBQ_Changed key, folderPath, "SCAN_ITEM_RESOLVE", n, detail
            needsFallback = True
        End If
        Set item = Nothing
        Exit Sub
    End If

    MBQ_QueueCandidate storeId, entryId, folderPath, observedMod, sourceSize, effectiveDate, _
        key, sinceDate, untilDate, output, seenItems, queued
End Sub

Private Function MBQ_ParseTableRow(ByRef values As Variant, ByVal rowIndex As Long, _
    ByRef entryId As String, ByRef messageClass As String, ByRef observedMod As Date, _
    ByRef sourceSize As Long, ByRef effectiveDate As Date, ByRef reason As String) As Long
    On Error GoTo InvalidRow
    MBQ_ParseTableRow = -1
    reason = "Table row has no usable EntryID."
    If Not MBQ_Text(values(0, rowIndex), entryId) Then Exit Function
    reason = "Table row has no usable MessageClass."
    If Not MBQ_Text(values(1, rowIndex), messageClass) Then Exit Function
    If Not MBQ_IsMailClass(messageClass) Then
        MBQ_ParseTableRow = 0
        Exit Function
    End If
    reason = "Table row has no usable LastModificationTime."
    If Not MBQ_Date(values(2, rowIndex), observedMod) Then Exit Function
    reason = "Table row has no usable Size."
    If Not MBQ_Size(values(3, rowIndex), sourceSize) Then Exit Function
    ' GetArray indexes are (column, row), both zero-based. Built-in date
    ' column names return local time. ReceivedTime has priority over SentOn.
    If Not MBQ_Date(values(4, rowIndex), effectiveDate) Then
        If Not MBQ_Date(values(5, rowIndex), effectiveDate) Then
            reason = "Table row has no usable ReceivedTime, SentOn, or CreationTime."
            If Not MBQ_Date(values(6, rowIndex), effectiveDate) Then Exit Function
        End If
    End If
    reason = vbNullString
    MBQ_ParseTableRow = 1
    Exit Function
InvalidRow:
    reason = "Cannot read table row: " & CStr(Err.Number) & " " & Err.Description
    MBQ_ParseTableRow = -1
End Function

Private Function MBQ_GetItem(ByVal entryId As String, ByVal storeId As String, ByRef item As Object, _
    ByRef errorNumber As Long, ByRef reason As String) As Boolean
    On Error GoTo Failed
    Set item = Nothing
    Set item = Application.Session.GetItemFromID(entryId, storeId)
    If item Is Nothing Then Err.Raise vbObjectError + 2830, , "GetItemFromID returned Nothing."
    MBQ_GetItem = True
    Exit Function
Failed:
    errorNumber = Err.Number
    reason = Err.Description
End Function

Private Sub MBQ_ItemsPass(ByVal folder As Object, ByVal storeId As String, ByVal folderPath As String, _
    ByVal folderKey As String, ByVal sinceDate As Date, ByVal untilDate As Date, _
    ByVal output As Object, ByVal seenItems As Object, ByVal badFolders As Object, ByRef queued As Long)

    Dim items As Object, item As Object, countBefore As Long, countAfter As Long
    Dim stage As String, n As Long, detail As String
    On Error GoTo Failed
    stage = "SOURCE"
    Set items = folder.Items
    countBefore = items.Count
    For Each item In items
        stage = "CONTROL"
        If Not MB_ContinueWork() Then Exit Sub
        stage = "RECORD"
        Call MBQ_ItemRow(item, storeId, folderPath, sinceDate, untilDate, output, seenItems, queued)
        stage = "SOURCE"
    Next item
    countAfter = items.Count
    stage = "CONTROL"
    If countBefore <> countAfter Then MBQ_Changed vbNullString, folderPath, "SCAN_ITEMS_CHANGED", 0, _
        "Folder item count changed during fallback scan: " & CStr(countBefore) & " -> " & CStr(countAfter)
    Exit Sub
Failed:
    n = Err.Number
    detail = Err.Description
    If stage <> "SOURCE" Then Err.Raise n, "MBQ_ItemsPass", detail
    MBQ_FolderFailed badFolders, folderKey, folderPath, "SCAN_ITEMS", n, detail
End Sub

Private Function MBQ_ItemRow(ByVal item As Object, ByVal storeId As String, ByVal folderPath As String, _
    ByVal sinceDate As Date, ByVal untilDate As Date, ByVal output As Object, ByVal seenItems As Object, _
    ByRef queued As Long, Optional ByVal expectedKey As String = "") As Boolean

    Dim entryId As String, key As String, observedMod As Date, effectiveDate As Date
    Dim sourceSize As Long, n As Long, detail As String, dateWarning As String
    If item Is Nothing Then
        MBQ_Changed expectedKey, folderPath, "SCAN_ITEM", 0, "An item is Nothing."
        Exit Function
    End If
    If Not TypeOf item Is Outlook.MailItem Then
        If Len(expectedKey) > 0 Then
            key = expectedKey
        ElseIf MBQ_ItemEntryId(item, entryId, n, detail) Then
            key = MB_SourceKey(storeId, entryId)
        End If
        If Len(key) > 0 Then
            If seenItems.Exists(key) Then
                If CStr(seenItems(key)) <> "PENDING" Then
                    MBQ_ItemRow = True
                    Exit Function
                End If
                seenItems(key) = "EXCLUDED"
            Else
                seenItems.Add key, "EXCLUDED"
            End If
        End If
        gPolicySkipped = gPolicySkipped + 1
        MBQ_ItemRow = True
        Exit Function
    End If
    If Not MBQ_ItemEntryId(item, entryId, n, detail) Then
        MBQ_Changed expectedKey, folderPath, "SCAN_ITEM_IDENTITY", n, detail
        Exit Function
    End If
    key = MB_SourceKey(storeId, entryId)
    If Len(expectedKey) > 0 And expectedKey <> key Then
        MBQ_Changed expectedKey, folderPath, "SCAN_ITEM_ID_CHANGED", 0, "Resolved item identity differs from the table row."
        Exit Function
    End If
    If seenItems.Exists(key) Then
        If CStr(seenItems(key)) <> "PENDING" Then
            MBQ_ItemRow = True
            Exit Function
        End If
    Else
        seenItems.Add key, "PENDING"
        gCandidates = gCandidates + 1
    End If

    If Not MBQ_ItemMetadata(item, observedMod, sourceSize, effectiveDate, dateWarning, n, detail) Then
        MBQ_Changed key, folderPath, "SCAN_ITEM_METADATA", n, detail
        Exit Function
    End If
    If Len(dateWarning) > 0 Then MBQ_Changed key, folderPath, "SCAN_DATE_FALLBACK", 0, dateWarning
    MBQ_QueueCandidate storeId, entryId, folderPath, observedMod, sourceSize, effectiveDate, _
        key, sinceDate, untilDate, output, seenItems, queued
    MBQ_ItemRow = True
End Function

Private Function MBQ_ItemEntryId(ByVal item As Object, ByRef entryId As String, ByRef errorNumber As Long, _
    ByRef reason As String) As Boolean
    On Error GoTo Failed
    entryId = CStr(item.EntryID)
    If Len(entryId) = 0 Then Err.Raise vbObjectError + 2840, , "MailItem has no EntryID."
    MBQ_ItemEntryId = True
    Exit Function
Failed:
    errorNumber = Err.Number
    reason = Err.Description
End Function

Private Function MBQ_ItemMetadata(ByVal item As Object, ByRef observedMod As Date, ByRef sourceSize As Long, _
    ByRef effectiveDate As Date, ByRef dateWarning As String, ByRef errorNumber As Long, ByRef reason As String) As Boolean

    Dim value As Variant, warning As String
    On Error GoTo Failed
    value = item.LastModificationTime
    If Not MBQ_Date(value, observedMod) Then Err.Raise vbObjectError + 2841, , "MailItem LastModificationTime is missing."
    value = item.Size
    If Not MBQ_Size(value, sourceSize) Then Err.Raise vbObjectError + 2842, , "MailItem Size is missing or invalid."
    If Not MBQ_ReadItemDate(item, 0, effectiveDate, warning) Then
        dateWarning = warning
        If Not MBQ_ReadItemDate(item, 1, effectiveDate, warning) Then
            If Len(warning) > 0 Then dateWarning = dateWarning & " " & warning
            If Not MBQ_ReadItemDate(item, 2, effectiveDate, warning) Then
                Err.Raise vbObjectError + 2843, , "MailItem has no usable ReceivedTime, SentOn, or CreationTime. " & warning
            End If
        End If
    End If
    MBQ_ItemMetadata = True
    Exit Function
Failed:
    errorNumber = Err.Number
    reason = Err.Description
End Function

Private Function MBQ_ReadItemDate(ByVal item As Object, ByVal whichDate As Long, ByRef result As Date, _
    ByRef warning As String) As Boolean
    Dim value As Variant, propertyName As String
    On Error GoTo Failed
    warning = vbNullString
    Select Case whichDate
        Case 0
            propertyName = "ReceivedTime"
            value = item.ReceivedTime
        Case 1
            propertyName = "SentOn"
            value = item.SentOn
        Case Else
            propertyName = "CreationTime"
            value = item.CreationTime
    End Select
    MBQ_ReadItemDate = MBQ_Date(value, result)
    Exit Function
Failed:
    warning = propertyName & " read failed: " & CStr(Err.Number) & " " & Err.Description
End Function

Private Sub MBQ_QueueCandidate(ByVal storeId As String, ByVal entryId As String, ByVal folderPath As String, _
    ByVal observedMod As Date, ByVal sourceSize As Long, ByVal effectiveDate As Date, ByVal key As String, _
    ByVal sinceDate As Date, ByVal untilDate As Date, ByVal output As Object, ByVal seenItems As Object, ByRef queued As Long)

    Dim fields As Variant, payload As String, rangeDate As Date
    rangeDate = effectiveDate
    If gModifiedWindow Then rangeDate = observedMod
    If gMode = "AUTO" Then
        If Not MBW_ShouldQueue(key, observedMod, effectiveDate, sinceDate, untilDate) Then
            seenItems(key) = "EXCLUDED"
            gPolicySkipped = gPolicySkipped + 1
            Exit Sub
        End If
    ElseIf sinceDate <> 0 Then
        If rangeDate < sinceDate Or rangeDate > untilDate Then
            seenItems(key) = "EXCLUDED"
            gPolicySkipped = gPolicySkipped + 1
            Exit Sub
        End If
    End If
    fields = Array("MBQ2", storeId, entryId, folderPath, MB_Stamp(observedMod), _
        CStr(sourceSize), MB_Stamp(effectiveDate), key)
    payload = MB_EncodeFields(fields)
    MB_WriteCheckedLine output, payload
    seenItems(key) = "QUEUED"
    queued = queued + 1
End Sub

Private Sub MBQ_Changed(ByVal key As String, ByVal folderPath As String, ByVal stage As String, _
    ByVal errorNumber As Long, ByVal detail As String)
    gScanChanged = gScanChanged + 1
    MB_LogEvent "WARN", stage, key, folderPath, errorNumber, detail
End Sub


' ---------- One exclusive OS lock per backup output root ----------

Private Sub MBH_AcquireRunLock(ByVal path As String)
    ' OPEN_ALWAYS preserves an existing lock file. The open handle, not file
    ' existence, is the lock. A process crash releases the OS handle.
    Dim lastError As Long
#If VBA7 Then
    Dim fileHandle As LongPtr
#Else
    Dim fileHandle As Long
#End If
    If MBH_RunLock <> 0 Then MBH_Fail "AcquireRunLock", "This VBA session already owns a run lock."
    path = MBH_NormalizePath(path)
    fileHandle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_READ Or MBH_GENERIC_WRITE, 0, 0, _
                                MBH_OPEN_ALWAYS, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If fileHandle = -1 Then
        lastError = Err.LastDllError
        MBH_Fail "AcquireRunLock", "Cannot acquire exclusive run lock; Win32=" & CStr(lastError) & "; path=" & path
    End If
    MBH_RunLock = fileHandle
End Sub

Private Sub MBH_ReleaseRunLock()
    ' Keep the lock file; deleting it after closing would introduce a race.
    Dim lastError As Long
    If MBH_RunLock <> 0 Then
        If MBH_CloseHandle(MBH_RunLock) = 0 Then
            lastError = Err.LastDllError
            MBH_Fail "ReleaseRunLock", "Cannot release run lock; Win32=" & CStr(lastError)
        End If
        MBH_RunLock = 0
    End If
End Sub

' ---------- SHA-256: real CNG hashes; text = UTF-16LE without BOM ----------

Private Sub MBH_CheckStatus(ByVal status As Long, ByVal operation As String)
    If status < 0 Then
        Err.Raise MBH_ERR_BASE + 1, "MailBackup.Helpers." & operation, _
                  "Windows CNG failed: NTSTATUS 0x" & Right$("00000000" & Hex$(status), 8)
    End If
End Sub

Private Sub MBH_Fail(ByVal operation As String, ByVal description As String)
    Err.Raise MBH_ERR_BASE + 2, "MailBackup.Helpers." & operation, description
End Sub

Private Sub MBH_HashOpen(ByRef context As MBH_HashContext)
    Dim algorithmName As String
    algorithmName = "SHA256"
    If MBH_Algorithm = 0 Then
        MBH_CheckStatus MBH_BCryptOpenAlgorithmProvider(MBH_Algorithm, StrPtr(algorithmName), 0, 0), "OpenSHA256"
    End If
    ' CNG allocates the object and releases its memory in BCryptDestroyHash.
    MBH_CheckStatus MBH_BCryptCreateHash(MBH_Algorithm, context.Handle, 0, 0, 0, 0, 0), "CreateSHA256"
    MBH_OpenHashes = MBH_OpenHashes + 1
End Sub

Private Sub MBH_HashClose(ByRef context As MBH_HashContext)
    Dim status As Long
    If context.Handle <> 0 Then
        status = MBH_BCryptDestroyHash(context.Handle)
        MBH_CheckStatus status, "DestroySHA256"
        context.Handle = 0
        MBH_OpenHashes = MBH_OpenHashes - 1
    End If
End Sub

Private Sub MBH_ReleaseCrypto()
    ' Call once from the run's final cleanup, after all hash calls have returned.
    If MBH_OpenHashes <> 0 Then MBH_Fail "ReleaseCrypto", "Hash objects are still active."
    If MBH_Algorithm <> 0 Then
        MBH_CheckStatus MBH_BCryptCloseAlgorithmProvider(MBH_Algorithm, 0), "CloseSHA256"
        MBH_Algorithm = 0
    End If
    Set MBH_FileSystemCache = Nothing
    Set MBH_NormalizedPaths = Nothing
End Sub

Private Function MBH_HashFinish(ByRef context As MBH_HashContext) As String
    Dim digest(0 To 31) As Byte
    Dim i As Long, result As String
    MBH_CheckStatus MBH_BCryptFinishHash(context.Handle, VarPtr(digest(0)), 32, 0), "FinishSHA256"
    For i = 0 To 31
        result = result & Right$("0" & Hex$(digest(i)), 2)
    Next i
    MBH_HashFinish = LCase$(result)
End Function

Private Function MBH_HashText(ByVal value As String) As String
    ' SHA-256 of the exact Windows BSTR payload: UTF-16LE, no BOM or terminator.
    ' Do not compare this to the hash of UTF-8/ASCII text with the same letters.
    Dim context As MBH_HashContext
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    On Error GoTo Failed
    MBH_HashOpen context
    If LenB(value) > 0 Then
        MBH_CheckStatus MBH_BCryptHashData(context.Handle, StrPtr(value), LenB(value), 0), "HashText"
    End If
    MBH_HashText = MBH_HashFinish(context)
    MBH_HashClose context
    Exit Function
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    On Error Resume Next
    MBH_HashClose context
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function MBH_HashFile(ByVal path As String) As String
    ' Streams the file bytes, not Outlook properties or a metadata approximation.
    ' ReadFile avoids VBA LOF/Get's signed-Long size/offset limitation.
    Dim context As MBH_HashContext
    Dim buffer(0 To 65535) As Byte
    Dim bytesRead As Long, lastError As Long, closeResult As Long
    Dim errorNumber As Long, errorSource As String, errorDescription As String
#If VBA7 Then
    Dim fileHandle As LongPtr
#Else
    Dim fileHandle As Long
#End If
    fileHandle = -1
    On Error GoTo Failed
    path = MBH_NormalizePath(path)
    MBH_HashOpen context
    ' Deny write/delete while hashing; concurrent read handles remain allowed.
    fileHandle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_READ, MBH_FILE_SHARE_READ, 0, _
                                 MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL Or MBH_FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If fileHandle = -1 Then
        lastError = Err.LastDllError
        MBH_Fail "HashFile", "Cannot open file; Win32=" & CStr(lastError) & "; path=" & path
    End If
    Do
        If MBH_ReadFile(fileHandle, buffer(0), 65536, bytesRead, 0) = 0 Then
            lastError = Err.LastDllError
            MBH_Fail "HashFile", "Cannot read file; Win32=" & CStr(lastError) & "; path=" & path
        End If
        If bytesRead = 0 Then Exit Do
        MBH_CheckStatus MBH_BCryptHashData(context.Handle, VarPtr(buffer(0)), bytesRead, 0), "HashFileData"
    Loop
    MBH_HashFile = MBH_HashFinish(context)
    closeResult = MBH_CloseHandle(fileHandle)
    lastError = Err.LastDllError
    If closeResult = 0 Then MBH_Fail "HashFile", "Cannot close file; Win32=" & CStr(lastError)
    fileHandle = -1
    MBH_HashClose context
    Exit Function
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    On Error Resume Next
    If fileHandle <> -1 Then closeResult = MBH_CloseHandle(fileHandle)
    MBH_HashClose context
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Function

' A held read handle permits readers but denies all write/delete handles. Use
' this around final content verification and the corresponding journal commit.
#If VBA7 Then
Private Function MBH_GuardFile(ByVal path As String) As LongPtr
    Dim handle As LongPtr
#Else
Private Function MBH_GuardFile(ByVal path As String) As Long
    Dim handle As Long
#End If
    Dim winError As Long
    path = MBH_NormalizePath(path)
    handle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_READ, MBH_FILE_SHARE_READ, 0, _
                            MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If handle = -1 Then
        winError = Err.LastDllError
        MBH_Fail "GuardFile", "Cannot protect final file verification; Win32=" & CStr(winError) & "; path=" & path
    End If
    MBH_GuardFile = handle
End Function

#If VBA7 Then
Private Sub MBH_CloseGuard(ByRef handle As LongPtr)
#Else
Private Sub MBH_CloseGuard(ByRef handle As Long)
#End If
    Dim winError As Long, detail As String
    If handle = 0 Then Exit Sub
    If MBH_CloseHandle(handle) = 0 Then
        winError = Err.LastDllError
        detail = "Cannot close verification guard; Win32=" & CStr(winError)
        ' Cleanup callers may suppress a second close exception. Persist the
        ' failure as run-level state so coverage cannot advance nevertheless.
        gCleanupIssue = gCleanupIssue & " Verification guard: " & detail
        gFatal = True
        If Len(gFatalText) > 0 Then gFatalText = gFatalText & " | "
        gFatalText = gFatalText & detail
        MBH_Fail "CloseGuard", detail
    End If
    handle = 0
End Sub

' ---------- Safe Windows path components ----------

Private Function MBH_FileSystem() As Object
    ' Reuse the late-bound COM service; do not cache filesystem query results.
    If MBH_FileSystemCache Is Nothing Then
        Set MBH_FileSystemCache = CreateObject("Scripting.FileSystemObject")
    End If
    Set MBH_FileSystem = MBH_FileSystemCache
End Function

Private Function MBH_Unit(ByVal text As String, ByVal position As Long) As Long
    MBH_Unit = CLng(AscW(Mid$(text, position, 1))) And &HFFFF&
End Function

Private Function MBH_LeftUtf16(ByVal text As String, ByVal maximumUnits As Long) As String
    ' Len/Left count UTF-16 units. Never leave a high surrogate at the cut.
    Dim lastUnit As Long
    If maximumUnits < 0 Then MBH_Fail "LeftUtf16", "maximumUnits must not be negative."
    If Len(text) <= maximumUnits Then
        MBH_LeftUtf16 = text
    ElseIf maximumUnits > 0 Then
        lastUnit = MBH_Unit(text, maximumUnits)
        If lastUnit >= &HD800& And lastUnit <= &HDBFF& Then maximumUnits = maximumUnits - 1
        MBH_LeftUtf16 = Left$(text, maximumUnits)
    End If
End Function

Private Function MBH_TrimEndDotsSpaces(ByVal value As String) As String
    Do While Len(value) > 0
        If Right$(value, 1) <> " " And Right$(value, 1) <> "." Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    MBH_TrimEndDotsSpaces = value
End Function

Private Function MBH_IsReservedName(ByVal value As String) As Boolean
    Dim stem As String, dotAt As Long, tail As String
    stem = UCase$(value)
    dotAt = InStr(1, stem, ".", vbBinaryCompare)
    If dotAt > 0 Then stem = Left$(stem, dotAt - 1)
    stem = MBH_TrimEndDotsSpaces(stem)
    Select Case stem
        Case "CON", "PRN", "AUX", "NUL", "CLOCK$", "CONIN$", "CONOUT$"
            MBH_IsReservedName = True
            Exit Function
    End Select
    If Len(stem) = 4 Then
        If Left$(stem, 3) = "COM" Or Left$(stem, 3) = "LPT" Then
            tail = Right$(stem, 1)
            If (tail >= "1" And tail <= "9") Or tail = ChrW$(&HB9) Or _
               tail = ChrW$(&HB2) Or tail = ChrW$(&HB3) Then MBH_IsReservedName = True
        End If
    End If
End Function

Private Function MBH_SafeComponent(ByVal raw As String, Optional ByVal maximumUnits As Long = 120, _
                                  Optional ByVal fallback As String = "unnamed") As String
    Dim i As Long, code As Long, nextCode As Long
    Dim one As String, result As String
    If maximumUnits < 1 Or maximumUnits > 255 Then MBH_Fail "SafeComponent", "maximumUnits must be 1..255."
    i = 1
    Do While i <= Len(raw)
        code = MBH_Unit(raw, i)
        one = Mid$(raw, i, 1)
        If code >= &HD800& And code <= &HDBFF& Then
            If i < Len(raw) Then
                nextCode = MBH_Unit(raw, i + 1)
                If nextCode >= &HDC00& And nextCode <= &HDFFF& Then
                    result = result & Mid$(raw, i, 2)
                    i = i + 2
                Else
                    result = result & "_": i = i + 1
                End If
            Else
                result = result & "_": i = i + 1
            End If
        ElseIf (code >= &HDC00& And code <= &HDFFF&) Or code < 32 Or _
               InStr(1, "<>:""/\|?*", one, vbBinaryCompare) > 0 Then
            result = result & "_": i = i + 1
        Else
            result = result & one: i = i + 1
        End If
    Loop
    result = MBH_TrimEndDotsSpaces(Trim$(result))
    If Len(result) = 0 Then
        If raw = fallback Or Len(fallback) = 0 Then
            result = "unnamed"
        Else
            result = MBH_SafeComponent(fallback, maximumUnits, vbNullString)
        End If
    End If
    result = MBH_TrimEndDotsSpaces(MBH_LeftUtf16(result, maximumUnits))
    If Len(result) = 0 Then result = MBH_LeftUtf16("unnamed", maximumUnits)
    ' Recheck after truncation: truncating e.g. CONX to 3 units creates CON.
    If MBH_IsReservedName(result) Then
        result = "_" & MBH_LeftUtf16(result, maximumUnits - 1)
    End If
    MBH_SafeComponent = result
End Function

Private Function MBH_NormalizePath(ByVal path As String) As String
    Dim fso As Object, value As String, relative As String
    Dim parts As Variant, item As Variant, driveLetter As String, cacheKey As String
    value = Replace(path, "/", "\")
    cacheKey = value
    If MBH_NormalizedPaths Is Nothing Then
        Set MBH_NormalizedPaths = CreateObject("Scripting.Dictionary")
        MBH_NormalizedPaths.CompareMode = vbBinaryCompare
    End If
    ' Only lexical normalization of absolute paths is cached. File existence,
    ' reparse attributes, hashes, permissions and account ownership are NEVER cached here.
    If MBH_NormalizedPaths.Exists(cacheKey) Then
        MBH_PathCacheHits = MBH_PathCacheHits + 1
        MBH_NormalizePath = CStr(MBH_NormalizedPaths(cacheKey))
        Exit Function
    End If
    MBH_PathCacheMisses = MBH_PathCacheMisses + 1
    If Len(value) = 0 Then MBH_Fail "NormalizePath", "An absolute path is required."
    If Left$(value, 4) = "\\?\" Or Left$(value, 4) = "\\.\" Then _
        MBH_Fail "NormalizePath", "Device/extended path namespaces are not supported."
    If Left$(value, 2) = "\\" Then
        parts = Split(Mid$(value, 3), "\")
        If UBound(parts) < 1 Then MBH_Fail "NormalizePath", "UNC server and share are required."
        If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Then MBH_Fail "NormalizePath", "Invalid UNC root."
    Else
        If Len(value) < 3 Then MBH_Fail "NormalizePath", "An absolute drive path is required."
        driveLetter = UCase$(Left$(value, 1))
        If Mid$(value, 2, 2) <> ":\" Or driveLetter < "A" Or driveLetter > "Z" Then _
            MBH_Fail "NormalizePath", "An absolute drive or UNC path is required."
    End If
    Set fso = MBH_FileSystem()
    value = fso.GetAbsolutePathName(value)
    Do While Len(value) > 3 And Right$(value, 1) = "\"
        value = Left$(value, Len(value) - 1)
    Loop
    If Len(value) > MBH_MAX_PATH_UNITS Then _
        MBH_Fail "NormalizePath", "Path exceeds conservative limit " & CStr(MBH_MAX_PATH_UNITS) & ": " & value
    If Left$(value, 2) = "\\" Then
        relative = Mid$(value, 3)
    Else
        relative = Mid$(value, 4)
    End If
    If Len(relative) > 0 Then
        parts = Split(relative, "\")
        For Each item In parts
            If Len(CStr(item)) = 0 Then MBH_Fail "NormalizePath", "Empty path component."
            If StrComp(MBH_SafeComponent(CStr(item), 255), CStr(item), vbBinaryCompare) <> 0 Then _
                MBH_Fail "NormalizePath", "Unsafe path component: " & CStr(item)
        Next item
    End If
    If MBH_NormalizedPaths.Count >= 512 Then MBH_NormalizedPaths.RemoveAll
    MBH_NormalizedPaths(cacheKey) = value
    MBH_NormalizePath = value
End Function

Private Function MBH_JoinPath(ByVal basePath As String, ByVal leaf As String) As String
    basePath = MBH_NormalizePath(basePath)
    If Len(leaf) = 0 Or leaf = "." Or leaf = ".." Then MBH_Fail "JoinPath", "A filename or directory name is required."
    If InStr(leaf, "\") > 0 Or InStr(leaf, "/") > 0 Or InStr(leaf, ":") > 0 Then _
        MBH_Fail "JoinPath", "leaf must be one path component."
    If StrComp(MBH_SafeComponent(leaf, 255), leaf, vbBinaryCompare) <> 0 Then _
        MBH_Fail "JoinPath", "leaf is not a safe Windows component."
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    MBH_JoinPath = MBH_NormalizePath(basePath & leaf)
End Function

Private Sub MBH_EnsureDirectory(ByVal path As String)
    Dim fso As Object, parentPath As String
    path = MBH_NormalizePath(path)
    Set fso = MBH_FileSystem()
    If fso.FolderExists(path) Then Exit Sub
    If fso.FileExists(path) Then MBH_Fail "EnsureDirectory", "A file occupies the directory path: " & path
    parentPath = fso.GetParentFolderName(path)
    If Len(parentPath) = 0 Or StrComp(parentPath, path, vbTextCompare) = 0 Then _
        MBH_Fail "EnsureDirectory", "Drive/share root is unavailable: " & path
    MBH_EnsureDirectory parentPath
    fso.CreateFolder path
End Sub

' ---------- Reserved stage files and no-overwrite commit ----------

Private Sub MBH_EnsureStageRegistry()
    If MBH_Stages Is Nothing Then
        Set MBH_Stages = CreateObject("Scripting.Dictionary")
        MBH_Stages.CompareMode = vbTextCompare
    End If
End Sub

Private Function MBH_NewStagePath(ByVal finalPath As String) As String
    ' Creates/reserves an EMPTY file with CREATE_NEW. Caller writes over this
    ' owned stage only. Its extension is kept for Outlook OpenSharedItem.
    Dim fso As Object, parentPath As String, extension As String
    Dim candidate As String, attempt As Long, lastError As Long, closeResult As Long
    Dim reservedHere As Boolean, registeredHere As Boolean, persistenceAttempted As Boolean
    Dim errorNumber As Long, errorSource As String, errorDescription As String
#If VBA7 Then
    Dim fileHandle As LongPtr
#Else
    Dim fileHandle As Long
#End If
    fileHandle = -1
    On Error GoTo Failed
    finalPath = MBH_NormalizePath(finalPath)
    Set fso = MBH_FileSystem()
    parentPath = fso.GetParentFolderName(finalPath)
    If Not fso.FolderExists(parentPath) Then MBH_Fail "NewStagePath", "Parent directory does not exist."
    extension = fso.GetExtensionName(finalPath)
    If Len(extension) > 0 Then extension = "." & extension
    MBH_EnsureStageRegistry
    For attempt = 1 To 64
        candidate = MBH_JoinPath(parentPath, "~mb_" & fso.GetBaseName(fso.GetTempName) & extension)
        If Not MBH_Stages.Exists(candidate) Then
            fileHandle = MBH_CreateFileW(StrPtr(candidate), MBH_GENERIC_WRITE, 0, 0, MBH_CREATE_NEW, MBH_FILE_ATTRIBUTE_NORMAL, 0)
            If fileHandle <> -1 Then
                reservedHere = True
                ' Track ownership before closing so failure cleanup remains safe.
                MBH_Stages.Add candidate, finalPath
                registeredHere = True
                closeResult = MBH_CloseHandle(fileHandle)
                lastError = Err.LastDllError
                If closeResult = 0 Then MBH_Fail "NewStagePath", "Cannot close stage; Win32=" & CStr(lastError)
                fileHandle = -1
                ' Reserve is durable before the caller may start SaveAs/write.
                ' A persistence failure preserves this stage for recovery.
                persistenceAttempted = True
                MBS_StageReserve candidate, finalPath
                MBH_NewStagePath = candidate
                Exit Function
            End If
            lastError = Err.LastDllError
            If lastError <> MBH_ERROR_FILE_EXISTS And lastError <> MBH_ERROR_ALREADY_EXISTS Then _
                MBH_Fail "NewStagePath", "Cannot reserve stage; Win32=" & CStr(lastError) & "; path=" & candidate
        End If
    Next attempt
    MBH_Fail "NewStagePath", "Could not reserve a unique stage name after 64 attempts."
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    On Error Resume Next
    If fileHandle <> -1 Then closeResult = MBH_CloseHandle(fileHandle)
    If registeredHere And Not persistenceAttempted Then
        MBH_DeleteStage candidate
    ElseIf reservedHere And Not persistenceAttempted Then
        ' This invocation created this file but registration failed.
        closeResult = MBH_DeleteFileW(StrPtr(candidate))
    End If
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Sub MBH_CommitStage(ByVal stagePath As String, ByVal finalPath As String, _
                           Optional ByVal expectedDigest As String = "", _
                           Optional ByVal expectedBytes As Double = -1)
    ' Data is flushed before publication; publication is same-directory,
    ' WRITE_THROUGH and never REPLACE_EXISTING/COPY_ALLOWED. The published file
    ' is flushed and checked again before this function can report success.
    ' Hardware/filesystem durability still requires Windows execution testing.
    Dim fso As Object, lastError As Long, actualDigest As String, actualBytes As Double
    Dim errorNumber As Long, errorSource As String, errorDescription As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    stagePath = MBH_NormalizePath(stagePath)
    finalPath = MBH_NormalizePath(finalPath)
    MBH_EnsureStageRegistry
    If Not MBH_Stages.Exists(stagePath) Then MBH_Fail "CommitStage", "Stage is not owned by this session."
    If StrComp(CStr(MBH_Stages(stagePath)), finalPath, vbTextCompare) <> 0 Then _
        MBH_Fail "CommitStage", "The final path differs from the stage reservation."
    Set fso = MBH_FileSystem()
    If StrComp(fso.GetParentFolderName(stagePath), fso.GetParentFolderName(finalPath), vbTextCompare) <> 0 Then _
        MBH_Fail "CommitStage", "Stage and final must be in the same directory."
    If Not fso.FileExists(stagePath) Then MBH_Fail "CommitStage", "Stage file is missing."
    If fso.FileExists(finalPath) Or fso.FolderExists(finalPath) Then _
        MBH_Fail "CommitStage", "Destination already exists; overwrite refused: " & finalPath
    If Len(expectedDigest) > 0 Then
        If Not MB_IsHexKey(expectedDigest) Then MBH_Fail "CommitStage", "Invalid expected file digest."
    End If
    If expectedBytes < -1 Then MBH_Fail "CommitStage", "Invalid expected file size."
    ' These are newly created, owned files; old/read-only legacy MSGs are never
    ' opened for write or flushed merely to adopt them.
    MBM_FlushPath stagePath
    guard = MBH_GuardFile(stagePath)
    actualDigest = MBH_HashFile(stagePath)
    actualBytes = CDbl(fso.GetFile(stagePath).Size)
    If Len(expectedDigest) = 0 Then expectedDigest = actualDigest
    If expectedBytes = -1 Then expectedBytes = actualBytes
    If actualDigest <> expectedDigest Or actualBytes <> expectedBytes Then _
        MBH_Fail "CommitStage", "Stage changed after its content verification."
    ' Persist verified expectations before rename, while the stage cannot be
    ' changed. Release the read guard only because rename requires delete access.
    MBS_StagePrepared stagePath, finalPath, expectedDigest, expectedBytes
    MBH_CloseGuard guard
    If MBM_MoveFileExW(StrPtr(stagePath), StrPtr(finalPath), 8) = 0 Then
        lastError = Err.LastDllError
        MBH_Fail "CommitStage", "Rename failed; Win32=" & CStr(lastError) & "; path=" & finalPath
    End If
    ' A published final is never a deletable stage, even if the following flush,
    ' verification or durable ledger update fails. Recovery retains PREPARED.
    MBH_Stages.Remove stagePath
    MBM_FlushPath finalPath
    guard = MBH_GuardFile(finalPath)
    actualDigest = MBH_HashFile(finalPath)
    actualBytes = CDbl(fso.GetFile(finalPath).Size)
    If actualDigest <> expectedDigest Or actualBytes <> expectedBytes Then _
        MBH_Fail "CommitStage", "Published file differs from the verified stage."
    MBS_StagePublished stagePath, finalPath, actualDigest, actualBytes
    MBH_CloseGuard guard
    Exit Sub
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub MBH_DeleteStage(ByVal stagePath As String)
    ' Only files registered by NewStagePath in this VBA session may be removed.
    Dim fso As Object, lastError As Long
    stagePath = MBH_NormalizePath(stagePath)
    MBH_EnsureStageRegistry
    If Not MBH_Stages.Exists(stagePath) Then MBH_Fail "DeleteStage", "Stage is not owned by this session."
    Set fso = MBH_FileSystem()
    If fso.FileExists(stagePath) Then
        If MBH_DeleteFileW(StrPtr(stagePath)) = 0 Then
            lastError = Err.LastDllError
            MBH_Fail "DeleteStage", "Cannot remove stage; Win32=" & CStr(lastError)
        End If
    End If
    MBH_Stages.Remove stagePath
    MBS_StageAborted stagePath
End Sub

' ---------- Manually runnable Windows self-test (not run in Linux) ----------

Private Function MBH_HashSelfTest() As Boolean
    Const EMPTY_SHA As String = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    Const UTF16_ABC_SHA As String = "13e228567e8249fce53337f25d7970de3bd68ab2653424c7b8f9fd05e33caedf"
    Const ASCII_ABC_SHA As String = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    Dim fso As Object, stagePath As String, finalPath As String
    Dim bytes(0 To 2) As Byte, fileNumber As Integer, fileIsOpen As Boolean
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    On Error GoTo Failed
    If MBH_HashText(vbNullString) <> EMPTY_SHA Then MBH_Fail "HashSelfTest", "Empty UTF-16 hash mismatch."
    If MBH_HashText("abc") <> UTF16_ABC_SHA Then MBH_Fail "HashSelfTest", "UTF-16 abc hash mismatch."
    Set fso = MBH_FileSystem()
    finalPath = MBH_JoinPath(CStr(fso.GetSpecialFolder(2).Path), "mailbackup_hash_selftest.bin")
    stagePath = MBH_NewStagePath(finalPath)
    If MBH_HashFile(stagePath) <> EMPTY_SHA Then MBH_Fail "HashSelfTest", "Empty file hash mismatch."
    bytes(0) = 97: bytes(1) = 98: bytes(2) = 99
    fileNumber = FreeFile
    Open stagePath For Binary Access Write Lock Read Write As #fileNumber
    fileIsOpen = True
    Put #fileNumber, 1, bytes
    Close #fileNumber
    fileIsOpen = False
    If MBH_HashFile(stagePath) <> ASCII_ABC_SHA Then MBH_Fail "HashSelfTest", "ASCII abc file hash mismatch."
    MBH_DeleteStage stagePath
    stagePath = vbNullString
    MBH_HashSelfTest = True
    Exit Function
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    On Error Resume Next
    If fileIsOpen Then Close #fileNumber
    If Len(stagePath) > 0 Then MBH_DeleteStage stagePath
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Function


' ======================= Locale / time contracts ============================
' MBC_Stamp retains the v2/v3 19-character LOCAL journal format. It does not
' reinterpret a legacy local stamp as UTC. Only the new watermark uses ...Z.
' VBA Calendar is project-wide. Helpers restore it without yielding. Run-level
' wrappers may also use Enter/Leave to protect older date/file-name code paths.
' No locale-formatted string is accepted as an archive control date.
Private Sub MBC_EnterGregorian(ByRef savedCalendar As Long, ByRef entered As Boolean)
    If entered Then Err.Raise MBC_ERR + 1, , "Gregorian calendar guard is already active."
    savedCalendar = Calendar
    Calendar = vbCalGreg
    entered = True
End Sub

Private Sub MBC_LeaveGregorian(ByVal savedCalendar As Long, ByRef entered As Boolean)
    If Not entered Then Exit Sub
    Calendar = savedCalendar
    entered = False
End Sub

Private Function MBC_Digits(ByVal value As Long, ByVal width As Long) As String
    Dim result As String, i As Long
    If value < 0 Or width < 1 Or width > 10 Then Err.Raise MBC_ERR + 2, , "Invalid unsigned integer field."
    result = String$(width, "0")
    For i = width To 1 Step -1
        Mid$(result, i, 1) = Chr$(48 + (value Mod 10))
        value = value \ 10
    Next i
    If value <> 0 Then Err.Raise MBC_ERR + 3, , "Unsigned integer exceeds field width."
    MBC_Digits = result
End Function

Private Function MBC_ReadDigits(ByVal value As String, ByVal offset As Long, ByVal count As Long) As Long
    Dim i As Long, digit As Long, result As Long
    If offset < 1 Or count < 1 Or count > 4 Or offset + count - 1 > Len(value) Then
        Err.Raise MBC_ERR + 4, , "Invalid numeric date field."
    End If
    For i = offset To offset + count - 1
        digit = AscW(Mid$(value, i, 1)) - 48
        If digit < 0 Or digit > 9 Then Err.Raise MBC_ERR + 5, , "Control date requires ASCII digits."
        result = result * 10 + digit
    Next i
    MBC_ReadDigits = result
End Function

Private Function MBC_GregorianDate(ByVal yyyy As Long, ByVal mm As Long, ByVal dd As Long, _
    Optional ByVal hh As Long = 0, Optional ByVal nn As Long = 0, Optional ByVal ss As Long = 0) As Date
    Dim savedCalendar As Long, entered As Boolean, result As Date
    Dim n As Long, reason As String
    On Error GoTo Failed
    If yyyy < 100 Or yyyy > 9999 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 _
        Or hh < 0 Or hh > 23 Or nn < 0 Or nn > 59 Or ss < 0 Or ss > 59 Then
        Err.Raise MBC_ERR + 6, , "Date field is out of range."
    End If
    MBC_EnterGregorian savedCalendar, entered
    result = DateSerial(yyyy, mm, dd)
    If CDbl(result) < 0 Then
        result = result - TimeSerial(hh, nn, ss)
    Else
        result = result + TimeSerial(hh, nn, ss)
    End If
    ' DateSerial normalizes impossible dates. Reject, for example, 2025-02-29.
    If Year(result) <> yyyy Or Month(result) <> mm Or Day(result) <> dd _
        Or Hour(result) <> hh Or Minute(result) <> nn Or Second(result) <> ss Then
        Err.Raise MBC_ERR + 7, , "Invalid Gregorian calendar date."
    End If
    MBC_LeaveGregorian savedCalendar, entered
    MBC_GregorianDate = result
    Exit Function
Failed:
    n = Err.Number: reason = Err.Description
    On Error Resume Next
    MBC_LeaveGregorian savedCalendar, entered
    On Error GoTo 0
    Err.Raise n, "MBC_GregorianDate", reason
End Function

Private Function MBC_Stamp(ByVal value As Date) As String
    Dim savedCalendar As Long, entered As Boolean, result As String
    Dim n As Long, reason As String
    If value = 0 Then Exit Function
    On Error GoTo Failed
    MBC_EnterGregorian savedCalendar, entered
    result = MBC_Digits(Year(value), 4) & "-" & MBC_Digits(Month(value), 2) & "-" & _
             MBC_Digits(Day(value), 2) & "T" & MBC_Digits(Hour(value), 2) & ":" & _
             MBC_Digits(Minute(value), 2) & ":" & MBC_Digits(Second(value), 2)
    MBC_LeaveGregorian savedCalendar, entered
    MBC_Stamp = result
    Exit Function
Failed:
    n = Err.Number: reason = Err.Description
    On Error Resume Next
    MBC_LeaveGregorian savedCalendar, entered
    On Error GoTo 0
    Err.Raise n, "MBC_Stamp", reason
End Function

Private Function MBC_ParseStamp(ByVal value As String) As Date
    Dim result As Date
    If Len(value) <> 19 Then Err.Raise MBC_ERR + 8, , "Control timestamp must have 19 characters."
    If Mid$(value, 5, 1) <> "-" Or Mid$(value, 8, 1) <> "-" Or Mid$(value, 11, 1) <> "T" _
        Or Mid$(value, 14, 1) <> ":" Or Mid$(value, 17, 1) <> ":" Then
        Err.Raise MBC_ERR + 9, , "Invalid control timestamp separators."
    End If
    result = MBC_GregorianDate(MBC_ReadDigits(value, 1, 4), MBC_ReadDigits(value, 6, 2), _
        MBC_ReadDigits(value, 9, 2), MBC_ReadDigits(value, 12, 2), _
        MBC_ReadDigits(value, 15, 2), MBC_ReadDigits(value, 18, 2))
    If MBC_Stamp(result) <> value Then Err.Raise MBC_ERR + 10, , "Noncanonical control timestamp."
    MBC_ParseStamp = result
End Function

Private Function MBC_ParseLegacyStamp(ByVal value As String) As Date
    Dim separator As String, number As Long
    ' v2/v3 Format$(..., "hh:nn:ss") expanded ':' to the OS time separator.
    ' An old fixed-width plan can therefore contain 10.20.30. The date and
    ' component order were still fixed; normalize ONLY that unambiguous legacy
    ' separator. Never use this tolerant reader for a new UTC checkpoint.
    If Len(value) <> 19 Then Err.Raise MBC_ERR + 17, , "Invalid legacy timestamp width."
    separator = Mid$(value, 14, 1)
    If Mid$(value, 17, 1) <> separator Then Err.Raise MBC_ERR + 18, , "Inconsistent legacy time separators."
    number = CLng(AscW(separator)) And &HFFFF&
    If number < 32 Or number = 127 Or (number >= 48 And number <= 57) _
        Or (number >= 65 And number <= 90) Or (number >= 97 And number <= 122) Then
        Err.Raise MBC_ERR + 19, , "Invalid legacy time separator."
    End If
    Mid$(value, 14, 1) = ":"
    Mid$(value, 17, 1) = ":"
    MBC_ParseLegacyStamp = MBC_ParseStamp(value)
End Function

Private Function MBC_FileStamp(ByVal value As Date) As String
    Dim stamp As String
    stamp = MBC_Stamp(value)
    If Len(stamp) <> 19 Then Err.Raise MBC_ERR + 11, , "Filename requires a nonempty date."
    MBC_FileStamp = Replace(Replace(Replace(stamp, "-", ""), ":", ""), "T", "_")
End Function

Private Function MBC_ParseFileStamp(ByVal value As String) As Date
    If Len(value) <> 15 Or Mid$(value, 9, 1) <> "_" Then
        Err.Raise MBC_ERR + 12, , "Filename date must be yyyyMMdd_HHmmss."
    End If
    MBC_ParseFileStamp = MBC_ParseStamp(Left$(value, 4) & "-" & Mid$(value, 5, 2) & "-" & _
        Mid$(value, 7, 2) & "T" & Mid$(value, 10, 2) & ":" & Mid$(value, 12, 2) & ":" & Right$(value, 2))
End Function

Private Function MBC_YearMonth(ByVal value As Date) As String
    Dim stamp As String
    stamp = MBC_Stamp(value)
    If Len(stamp) <> 19 Then Err.Raise MBC_ERR + 13, , "Year/month filing requires a nonempty date."
    MBC_YearMonth = Left$(stamp, 4) & "\" & Mid$(stamp, 6, 2)
End Function

Private Function MBC_ShiftMonths(ByVal value As Date, ByVal months As Long) As Date
    Dim savedCalendar As Long, entered As Boolean, result As Date
    Dim n As Long, reason As String
    On Error GoTo Failed
    MBC_EnterGregorian savedCalendar, entered
    result = DateAdd("m", months, value)
    MBC_LeaveGregorian savedCalendar, entered
    MBC_ShiftMonths = result
    Exit Function
Failed:
    n = Err.Number: reason = Err.Description
    On Error Resume Next
    MBC_LeaveGregorian savedCalendar, entered
    On Error GoTo 0
    Err.Raise n, "MBC_ShiftMonths", reason
End Function

Private Function MBC_ValidDate(ByVal value As Date) As Boolean
    ' OLE DATE numeric boundaries: 1900-01-01 (2) <= date < 2100-01-01 (73051).
    ' This intentionally preserves the existing mail-date acceptance policy.
    MBC_ValidDate = (CDbl(value) >= 2# And CDbl(value) < 73051#)
End Function

Private Function MBC_UtcNow() As Date
    Dim value As MBC_SYSTEMTIME
    MBC_GetSystemTime value
    ' A watermark is a whole UTC second. Do not derive it by converting Now:
    ' a local time in the repeated DST hour has two possible UTC meanings.
    MBC_UtcNow = MBC_GregorianDate(value.wYear, value.wMonth, value.wDay, _
        value.wHour, value.wMinute, value.wSecond)
End Function

Private Function MBC_UtcStamp(ByVal utcValue As Date) As String
    If utcValue = 0 Then Exit Function
    MBC_UtcStamp = MBC_Stamp(utcValue) & "Z"
End Function

Private Function MBC_ParseUtcStamp(ByVal value As String) As Date
    If Len(value) <> 20 Or Right$(value, 1) <> "Z" Then
        Err.Raise MBC_ERR + 14, , "UTC control timestamp must end in Z."
    End If
    MBC_ParseUtcStamp = MBC_ParseStamp(Left$(value, 19))
End Function

Private Function MBC_LocalToUtc(ByVal localValue As Date) As Date
    ' Only a convenience for legacy user-selected local ranges. Never use for
    ' a durable checkpoint or to disambiguate the repeated DST fall-back hour.
    MBC_LocalToUtc = Application.Session.DefaultStore.PropertyAccessor.LocalTimeToUTC(localValue)
End Function

Private Function MBC_UtcToLocal(ByVal utcValue As Date) As Date
    MBC_UtcToLocal = Application.Session.DefaultStore.PropertyAccessor.UTCToLocalTime(utcValue)
End Function

Private Function MBC_DateUTC(ByVal item As Object, ByVal schemaName As String, _
    ByRef result As Date, ByRef reason As String) As Boolean
    Dim value As Variant
    On Error GoTo Failed
    result = 0: reason = ""
    value = item.PropertyAccessor.GetProperty(schemaName)
    If VarType(value) <> vbDate Then
        reason = "MAPI UTC date did not return VT_DATE: " & schemaName
        Exit Function
    End If
    result = CDate(value)
    If Not MBC_ValidDate(result) Then
        result = 0
        reason = "MAPI UTC date is missing or outside the supported range: " & schemaName
        Exit Function
    End If
    MBC_DateUTC = True
    Exit Function
Failed:
    result = 0
    reason = "MAPI UTC date unavailable: " & schemaName & "; " & Trim$(Str$(Err.Number)) & " " & Err.Description
End Function

Private Function MBC_TimeRule(ByRef value As MBC_SYSTEMTIME) As String
    MBC_TimeRule = MBC_Digits(value.wYear, 4) & ":" & MBC_Digits(value.wMonth, 2) & ":" & _
        MBC_Digits(value.wDayOfWeek, 1) & ":" & MBC_Digits(value.wDay, 2) & ":" & _
        MBC_Digits(value.wHour, 2) & ":" & MBC_Digits(value.wMinute, 2) & ":" & _
        MBC_Digits(value.wSecond, 2) & ":" & MBC_Digits(value.wMilliseconds, 3)
End Function

Private Function MBC_TimeContext() As String
    Dim value As MBC_DYNAMIC_TIME_ZONE_INFORMATION, status As Long
    Dim zoneKey As String, i As Long, number As Long, payload As String
    If LenB(value) <> 432 Then Err.Raise MBC_ERR + 15, , "Unexpected Windows timezone structure layout."
    status = MBC_GetDynamicTimeZoneInformation(value)
    If status = -1 Then Err.Raise MBC_ERR + 16, , "Windows timezone information is unavailable."
    For i = 0 To 127
        number = CLng(value.TimeZoneKeyName(i)) And &HFFFF&
        If number = 0 Then Exit For
        zoneKey = zoneKey & ChrW$(number)
    Next i
    ' Display names vary with Windows UI language; omit them. Normal seasonal
    ' DST state (status=1/2) is also omitted: the transition RULES identify it.
    ' This is a settings-change detector, not a hash of every historical rule.
    payload = "MBTZ2|" & zoneKey & "|" & Trim$(Str$(value.Bias)) & "|" & _
        Trim$(Str$(value.StandardBias)) & "|" & Trim$(Str$(value.DaylightBias)) & "|" & _
        MBC_TimeRule(value.StandardDate) & "|" & MBC_TimeRule(value.DaylightDate) & "|" & _
        IIf(value.DynamicDaylightTimeDisabled = 0, "0", "1") & "|NUM=" & _
        Replace(Format$(12345.5, "0.0000000000"), ",", ".")
    ' NUM detects a locale change that can affect the preserved MBREV2 numeric
    ' formatter (including native digits/custom decimal separators). It does
    ' not alter any existing source/revision key or infer prior calendar use.
    MBC_TimeContext = MBH_HashText(payload)
End Function

Private Function MBC_RejectStamp(ByVal value As String, ByVal utcMode As Boolean) As Boolean
    Dim parsed As Date
    On Error GoTo Rejected
    If utcMode Then parsed = MBC_ParseUtcStamp(value) Else parsed = MBC_ParseStamp(value)
    Exit Function
Rejected:
    MBC_RejectStamp = True
End Function

Private Function MBC_LocaleSelfTest(ByRef detail As String) As Boolean
    Dim originalCalendar As Long, haveCalendar As Boolean, value As Date, currentUtc As Date
    Dim context As String, bad As Variant, candidate As Variant
    Dim n As Long, reason As String
    On Error GoTo Failed
    originalCalendar = Calendar: haveCalendar = True
    value = MBC_GregorianDate(2024, 2, 29, 23, 59, 58)
    Calendar = vbCalHijri
    If MBC_Stamp(value) <> "2024-02-29T23:59:58" Then Err.Raise MBC_ERR + 20, , "Gregorian stamp varies with Calendar."
    If Calendar <> vbCalHijri Then Err.Raise MBC_ERR + 21, , "Stamp helper did not restore Calendar."
    If MBC_ParseStamp("2024-02-29T23:59:58") <> value Then Err.Raise MBC_ERR + 22, , "Strict Gregorian parser failed."
    If MBC_ParseLegacyStamp("2024-02-29T23.59.58") <> value Then Err.Raise MBC_ERR + 22, , "Localized legacy timestamp failed."
    If Calendar <> vbCalHijri Then Err.Raise MBC_ERR + 23, , "Parser did not restore Calendar."
    If MBC_FileStamp(value) <> "20240229_235958" Then Err.Raise MBC_ERR + 24, , "ASCII filename stamp failed."
    If MBC_ParseFileStamp("20240229_235958") <> value Then Err.Raise MBC_ERR + 25, , "Filename parser failed."
    If MBC_YearMonth(value) <> "2024\02" Then Err.Raise MBC_ERR + 26, , "Gregorian year/month failed."
    If MBC_ShiftMonths(MBC_GregorianDate(2024, 3, 31), -1) <> MBC_GregorianDate(2024, 2, 29) Then
        Err.Raise MBC_ERR + 27, , "Gregorian calendar-month subtraction failed."
    End If
    bad = Array("2025-02-29T00:00:00", "2024-13-01T00:00:00", "2024-02-30T00:00:00", _
        "2024-01-00T00:00:00", "2024-01-01T24:00:00", "2024-01-01T00:60:00", _
        "2024-01-01T00:00:60", "2024/01/01T00:00:00", "2024-01-01 00:00:00", _
        "02/03/2024 04:05:06", "2024-01-01T00:00:00Z", " 2024-01-01T00:00:00", _
        ChrW$(&HFF12) & "024-01-01T00:00:00")
    For Each candidate In bad
        If Not MBC_RejectStamp(CStr(candidate), False) Then Err.Raise MBC_ERR + 28, , "Invalid local stamp was accepted."
        If Calendar <> vbCalHijri Then Err.Raise MBC_ERR + 29, , "Failed parse did not restore Calendar."
    Next candidate
    If MBC_UtcStamp(value) <> "2024-02-29T23:59:58Z" Then Err.Raise MBC_ERR + 30, , "UTC stamp failed."
    If MBC_ParseUtcStamp("2024-02-29T23:59:58Z") <> value Then Err.Raise MBC_ERR + 31, , "UTC parser failed."
    If Not MBC_RejectStamp("2024-02-29T23:59:58", True) Then Err.Raise MBC_ERR + 32, , "Unqualified local stamp accepted as UTC."
    currentUtc = MBC_UtcNow()
    If Not MBC_ValidDate(currentUtc) Then Err.Raise MBC_ERR + 33, , "Windows UTC clock returned an unusable date."
    context = MBC_TimeContext()
    If Len(context) <> 64 Then Err.Raise MBC_ERR + 34, , "Timezone-context hash failed."
    If Calendar <> vbCalHijri Then Err.Raise MBC_ERR + 35, , "Locale helper changed caller Calendar."
    Calendar = originalCalendar: haveCalendar = False
    detail = "PASS: ASCII Gregorian stamps, strict date rejection, Hijri restoration, UTC clock, timezone ABI."
    MBC_LocaleSelfTest = True
    Exit Function
Failed:
    n = Err.Number: reason = Err.Description
    On Error Resume Next
    If haveCalendar Then Calendar = originalCalendar
    On Error GoTo 0
    detail = "FAIL: locale/time check " & Trim$(Str$(n)) & " - " & reason
End Function


' v3.2: recover filename space previously used by the 64-hex revision suffix.
' Full identity/revision and the full subject remain in the checked journal.
' A filename is a search/display hint, never the authority for deduplication.

Private Function MBN_SenderLabel(ByVal raw As String) As String
    Dim value As String, parenthesis As Long
    value = raw
    parenthesis = InStr(1, value, "(", vbBinaryCompare)
    If parenthesis > 0 Then value = Left$(value, parenthesis - 1)
    ' Same original sender convention, without splitting a surrogate pair.
    value = MBH_LeftUtf16(value, MBN_SENDER_UNITS)
    If Len(Trim$(value)) = 0 Then value = "unknown"
    MBN_SenderLabel = value
End Function

Private Function MBN_MessageLeaf(ByVal stamp15 As String, ByVal senderName As String, _
                                 ByVal subject As String, ByVal parent As String) As String
    Dim sender As String, prefix As String, stem As String, stemBudget As Long, code As Long, i As Long
    ' The caller supplies a Gregorian, ASCII yyyyMMdd_HHmmss timestamp.
    ' Do not call Format here: its result can depend on VBA.Calendar.
    If Len(stamp15) <> 15 Or Mid$(stamp15, 9, 1) <> "_" Then
        Err.Raise MB_ERR + 820, "MBN_MessageLeaf", "Invalid invariant filename timestamp."
    End If
    For i = 1 To 15
        If i <> 9 Then
            code = AscW(Mid$(stamp15, i, 1))
            If code < 48 Or code > 57 Then
                Err.Raise MB_ERR + 820, "MBN_MessageLeaf", "Filename timestamp must use ASCII digits."
            End If
        End If
    Next i
    sender = MBN_SenderLabel(senderName)
    prefix = stamp15 & "_" & sender & "_"
    stemBudget = MBH_MAX_PATH_UNITS - Len(parent) - 1 - Len(".msg") - MBN_COLLISION_RESERVE
    If stemBudget < Len(prefix) + 1 Then
        Err.Raise MB_ERR + 821, "MBN_MessageLeaf", "Message parent path leaves no subject space."
    End If
    ' Sanitizing the final leaf preserves literal spaces and valid Unicode,
    ' including apostrophes and dots before the extension. Windows-forbidden
    ' characters become underscores. Only the tail is shortened if required.
    ' The 8-unit reserve accommodates the existing _r00001 collision suffix;
    ' MB_UnusedFinal/CommitStage still provide no-overwrite collision handling.
    If Len(Trim$(subject)) = 0 Then subject = vbNullString
    stem = MBH_LeftUtf16(prefix & subject, stemBudget)
    MBN_MessageLeaf = MBH_SafeComponent(stem & ".msg", stemBudget + 4, "message.msg")
End Function

Private Function MBN_TestRejectName(ByVal stamp15 As String, ByVal parent As String) As Boolean
    Dim ignored As String
    On Error GoTo Rejected
    ignored = MBN_MessageLeaf(stamp15, "sender", "subject", parent)
    Exit Function
Rejected:
    MBN_TestRejectName = True
End Function

Private Function MBN_SelfTest() As Boolean
    ' Pure string tests: no mail creation, filesystem mutation or public macro.
    Const STAMP As String = "20241107_093012"
    Dim parent As String, leaf As String, prefix As String, subject As String
    Dim emoji As String, stemBudget As Long, subjectBudget As Long, unicodeText As String
    parent = "C:\Mail\account\2024\11"
    leaf = MBN_MessageLeaf(STAMP, "ABCDEFGHIJK(Team)", "Quarterly report", parent)
    If leaf <> STAMP & "_ABCDEFGHIJ_Quarterly report.msg" Then Exit Function
    leaf = MBN_MessageLeaf(STAMP, "Kim (Team)", "O'Brien  Q4", parent)
    If leaf <> STAMP & "_Kim _O'Brien  Q4.msg" Then Exit Function
    leaf = MBN_MessageLeaf(STAMP, "sender", "A...  ", parent)
    If leaf <> STAMP & "_sender_A...  .msg" Then Exit Function
    subject = String$(120, "S")
    leaf = MBN_MessageLeaf(STAMP, "sender", subject, parent)
    If InStr(1, leaf, subject, vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, leaf, "__", vbBinaryCompare) > 0 Then Exit Function
    unicodeText = ChrW$(&HD55C) & ChrW$(&HAE00) & ChrW$(&H4E2D) & ChrW$(&H6587)
    leaf = MBN_MessageLeaf(STAMP, "sender", unicodeText, parent)
    If leaf <> STAMP & "_sender_" & unicodeText & ".msg" Then Exit Function
    leaf = MBN_MessageLeaf(STAMP, "sender", "CON.txt: A/B?*" & vbTab, parent)
    If leaf <> STAMP & "_sender_CON.txt_ A_B___.msg" Then Exit Function
    If MBN_MessageLeaf(STAMP, "   (Team)", "   ", parent) <> STAMP & "_unknown_.msg" Then Exit Function
    If MBN_MessageLeaf(STAMP, "sender", "", parent) <> STAMP & "_sender_.msg" Then Exit Function
    ' Account root 80 + \yyyy\mm 8 = parent 88 UTF-16 units.
    parent = "C:\" & String$(77, "p") & "\2024\11"
    prefix = STAMP & "_ABCDEFGHIJ_"
    stemBudget = MBH_MAX_PATH_UNITS - Len(parent) - 1 - 4 - MBN_COLLISION_RESERVE
    subjectBudget = stemBudget - Len(prefix)
    If subjectBudget <> 107 Then Exit Function
    leaf = MBN_MessageLeaf(STAMP, "ABCDEFGHIJ", String$(500, "X"), parent)
    If Len(leaf) <> stemBudget + 4 Then Exit Function
    If Len(parent) + 1 + Len(leaf) + MBN_COLLISION_RESERVE <> MBH_MAX_PATH_UNITS Then Exit Function
    If InStr(1, leaf, "__", vbBinaryCompare) > 0 Then Exit Function
    emoji = ChrW$(&HD83D) & ChrW$(&HDE00)
    subject = String$(subjectBudget - 1, "X") & emoji & "tail"
    leaf = MBN_MessageLeaf(STAMP, "ABCDEFGHIJ", subject, parent)
    If leaf <> prefix & String$(subjectBudget - 1, "X") & ".msg" Then Exit Function
    subject = String$(subjectBudget - 2, "X") & emoji & "tail"
    leaf = MBN_MessageLeaf(STAMP, "ABCDEFGHIJ", subject, parent)
    If leaf <> prefix & String$(subjectBudget - 2, "X") & emoji & ".msg" Then Exit Function
    If MBN_SenderLabel(String$(9, "S") & emoji & "tail") <> String$(9, "S") Then Exit Function
    If Not MBN_TestRejectName("20241107-093012", parent) Then Exit Function
    If Not MBN_TestRejectName("20241107_09A012", parent) Then Exit Function
    If Not MBN_TestRejectName(STAMP, "C:\" & String$(220, "P")) Then Exit Function
    MBN_SelfTest = True
End Function


' ===================== Restored / moved message identity ====================
Private Sub MBI_Reset()
    MBV_Reset
    Set gIdentityBuckets = Nothing: Set gIdentitySeen = Nothing: Set gIdentityHeaders = Nothing
    Set gIdentityLookupState = Nothing
    gIdentityAliases = 0: gIdentityHeaderReads = 0: gIdentityHeaderHits = 0
End Sub

Private Function MBI_SubjectKey(ByVal storeId As String, ByVal subject As String) As String
    MBI_SubjectKey = MBX_StoreToken(storeId) & "|" & MBH_HashText("MBI1|" & subject)
End Function

Private Sub MBI_CacheRecord(ByVal f As Variant)
    Dim kind As String, bucket As String, unique As String, key As String, dateKey As String
    Dim candidates As Collection, dates As Object, name As Variant, part As Object
    kind = CStr(f(1))
    If kind <> "M" And kind <> "N" And kind <> "L" And kind <> "R" Then Exit Sub
    If gIdentityBuckets Is Nothing Then
        Set gIdentityBuckets = CreateObject("Scripting.Dictionary")
        Set gIdentitySeen = CreateObject("Scripting.Dictionary")
        Set gIdentityLookupState = CreateObject("Scripting.Dictionary")
        For Each name In Array("dates", "entries", "counts", "proofs", "unknown")
            Set part = CreateObject("Scripting.Dictionary")
            gIdentityLookupState.Add CStr(name), part
        Next name
    End If
    bucket = MBI_SubjectKey(CStr(f(7)), CStr(f(12)))
    key = MBX_IndexKey(kind, CStr(f(2)))
    ' Keep distinct reference keys even when they share a file: a later journal
    ' row may replace one key. Runtime path+hash dedup avoids comparing shared
    ' bytes repeatedly and also covers verified move redirects.
    unique = bucket & "|" & key
    If Not gIdentitySeen.Exists(unique) Then
        If Not gIdentityBuckets.Exists(bucket) Then
            Set candidates = New Collection
            gIdentityBuckets.Add bucket, candidates
        Else
            Set candidates = gIdentityBuckets(bucket)
        End If
        candidates.Add key
        gIdentitySeen.Add unique, True
    End If
    ' A policy upgrade may refresh a displayed date without changing file bytes.
    ' Keep exact-date ordering independently of the all-subject append cursor.
    dateKey = MBI_DateScope(bucket, CStr(f(11)))
    unique = dateKey & "|" & key
    If Not gIdentitySeen.Exists(unique) Then
        Set dates = gIdentityLookupState("dates")
        MBI_AddLookupEntry dates, dateKey, key
        gIdentitySeen.Add unique, True
    End If
End Sub

Private Function MBI_DateScope(ByVal bucket As String, ByVal stamp As String) As String
    MBI_DateScope = bucket & "|DATE|" & MBH_HashText(stamp)
End Function

Private Sub MBI_AddLookupEntry(ByVal lookup As Object, ByVal key As String, ByVal value As String)
    Dim values As Collection, first As String
    If Not lookup.Exists(key) Then
        lookup.Add key, value
    ElseIf IsObject(lookup(key)) Then
        Set values = lookup(key): values.Add value
    Else
        first = CStr(lookup(key))
        Set values = New Collection
        values.Add first: values.Add value
        lookup.Remove key: lookup.Add key, values
    End If
End Sub

Private Function MBI_LookupGroup(ByVal lookup As Object, ByVal key As String) As Collection
    Dim values As Collection
    If lookup.Exists(key) Then
        If IsObject(lookup(key)) Then
            Set MBI_LookupGroup = lookup(key)
            Exit Function
        End If
    End If
    Set values = New Collection
    If lookup.Exists(key) Then values.Add CStr(lookup(key))
    Set MBI_LookupGroup = values
End Function

Private Function MBI_EnsureLookup(ByVal scope As String, ByVal candidates As Collection) As Boolean
    Dim entries As Object, counts As Object, proofs As Object, unknown As Object
    Dim position As Long, key As String, fields As Variant, path As String, proof As String
    Dim hint As Variant, reason As String, result As Long, selector As String
    Set entries = gIdentityLookupState("entries"): Set counts = gIdentityLookupState("counts")
    Set proofs = gIdentityLookupState("proofs"): Set unknown = gIdentityLookupState("unknown")
    If counts.Exists(scope) Then position = CLng(counts(scope))
    ' Extend only the unclassified suffix. Newly committed N/R rows never force
    ' an all-subject rebuild. Each physical proof appears once in each scope.
    Do While position < candidates.Count
        If Not MB_ContinueWork() Then Exit Function
        key = CStr(candidates(position + 1))
        If gIndex.Exists(key) Then
            fields = MB_IndexFields(key): path = MBX_IndexedPath(key, fields)
            proof = scope & "|" & MB_LegacyOwnerKey(path, CStr(fields(2)))
            If Not proofs.Exists(proof) Then
                result = MBI_ReadHeader(key, hint, reason)
                If result = 3 Then Exit Function
                If result = 1 Then
                    If CBool(hint(2)) And Len(CStr(hint(0))) > 0 Then
                        selector = "ID|" & MBH_HashText(CStr(hint(0)))
                        MBI_AddLookupEntry entries, scope & "|" & selector, key
                        MBI_AddLookupEntry entries, scope & "|PRESENT", key
                    ElseIf Not CBool(hint(2)) Then
                        selector = "SENT|" & MBH_HashText(CStr(hint(1)))
                        MBI_AddLookupEntry entries, scope & "|" & selector, key
                        MBI_AddLookupEntry entries, scope & "|ABSENT", key
                    Else
                        MBI_AddLookupEntry entries, scope & "|EMPTY", key
                    End If
                Else
                    ' A failed header remains uncertainty, never an absent ID.
                    ' Retry on the next run; do not repeat the same failed I/O
                    ' for every new mail with this subject in the current run.
                    If Not unknown.Exists(scope) Then unknown.Add scope, reason
                End If
                proofs.Add proof, True
            End If
        End If
        position = position + 1
        counts(scope) = position
    Loop
    MBI_EnsureLookup = True
End Function

Private Function MBI_QueryGroups(ByVal sourceId As String, ByVal sourceHas As Boolean, _
                                  ByVal sourceSent As String) As Variant
    If sourceHas And Len(sourceId) > 0 Then
        MBI_QueryGroups = Array("ID|" & MBH_HashText(sourceId), "ABSENT", "EMPTY")
    ElseIf Not sourceHas Then
        MBI_QueryGroups = Array("SENT|" & MBH_HashText(sourceSent), "PRESENT", "EMPTY")
    Else
        MBI_QueryGroups = Array("PRESENT", "ABSENT", "EMPTY")
    End If
End Function

Private Function MBI_ReadHeader(ByVal key As String, ByRef hint As Variant, ByRef reason As String) As Long
    Dim f As Variant, path As String, root As String, unique As String, disk As Object
    Dim obj As Object, copy As Outlook.MailItem, beforeDate As Date, bytes As Double
    Dim messageId As String, present As Boolean, sentStamp As String, n As Long, detail As String
    On Error GoTo Failed
    f = MB_IndexFields(key): path = MBX_IndexedPath(key, f)
    unique = MB_LegacyOwnerKey(path, CStr(f(2)))
    If gIdentityHeaders Is Nothing Then Set gIdentityHeaders = CreateObject("Scripting.Dictionary")
    If gIdentityHeaders.Exists(unique) Then
        gIdentityHeaderHits = gIdentityHeaderHits + 1
        hint = gIdentityHeaders(unique): MBI_ReadHeader = 1: Exit Function
    End If
    If Not MB_ContinueWork() Then MBI_ReadHeader = 3: Exit Function
    gIdentityHeaderReads = gIdentityHeaderReads + 1
    root = MBM_ResolvedRoot(MBX_IndexedOriginal(key, f), CStr(f(2)), MBX_IndexedRoot(key))
    path = MBL_CanonicalLegacyPath(path, root, True)
    Set disk = gFSO.GetFile(path)
    bytes = CDbl(disk.Size): beforeDate = disk.DateLastModified
    If bytes <= 0 Or bytes <> MB_ByteCount(CStr(f(1))) Then _
        Err.Raise MB_ERR + 738, , "Identity header candidate size changed."
    If MBH_HashFile(path) <> CStr(f(2)) Then _
        Err.Raise MB_ERR + 739, , "Identity header candidate SHA-256 changed."
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 730, , "Identity candidate is not MailItem."
    Set copy = obj
    messageId = MBL_InternetId(copy, present): sentStamp = MB_Stamp(copy.SentOn)
    copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    If MBH_HashFile(path) <> CStr(f(2)) Then Err.Raise MB_ERR + 740, , "Identity header candidate changed while read."
    Set disk = gFSO.GetFile(path)
    If CDbl(disk.Size) <> bytes Or disk.DateLastModified <> beforeDate Then _
        Err.Raise MB_ERR + 741, , "Identity header candidate size/time changed while read."
    hint = Array(messageId, sentStamp, present)
    gIdentityHeaders.Add unique, hint
    MBI_ReadHeader = 1
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not copy Is Nothing Then copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    On Error GoTo 0
    If gFatal Then Err.Raise n, "MBI_ReadHeader", detail
    reason = "Cannot verify identity candidate headers: " & detail
    If gCancel Then MBI_ReadHeader = 3 Else MBI_ReadHeader = 2
End Function

Private Function MBI_TimestampDifference(ByVal mail As Outlook.MailItem, ByVal path As String, _
                                        ByRef reason As String) As Long
    Dim obj As Object, copy As Outlook.MailItem, sourceId As String, copyId As String
    Dim sourceHas As Boolean, copyHas As Boolean, result As Long, n As Long, detail As String
    On Error GoTo Failed
    result = 2
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 730, , "Identity candidate is not MailItem."
    Set copy = obj
    sourceId = MBL_InternetId(mail, sourceHas): copyId = MBL_InternetId(copy, copyHas)
    If sourceHas And copyHas And Len(sourceId) > 0 And Len(copyId) > 0 Then
        If StrComp(sourceId, copyId, vbBinaryCompare) <> 0 Then
            result = 0
        Else
            reason = "Same Message-ID but different message times; provider/import rewriting requires review."
        End If
    ElseIf Not sourceHas And Not copyHas Then
        If MB_Stamp(mail.SentOn) <> MB_Stamp(copy.SentOn) Then
            result = 0
        Else
            reason = "No Message-ID and the same sent time but different received time; restored identity is uncertain."
        End If
    Else
        reason = "Message-ID presence changed with message times; restored identity is uncertain."
    End If
    copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    MBI_TimestampDifference = result
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not copy Is Nothing Then copy.Close olDiscard
    Set copy = Nothing: Set obj = Nothing
    On Error GoTo 0
    If gFatal Then Err.Raise n, "MBI_TimestampDifference", detail
    reason = "Cannot establish candidate time/ID relationship: " & detail
    MBI_TimestampDifference = 2
End Function

Private Function MBI_CompareContents(ByVal mail As Outlook.MailItem, ByVal path As String, _
                                     ByRef reason As String) As Long
    ' 0=explicitly different; 1=equal under the declared comparator;
    ' 2=unknown/unsupported; 3=cancel. Message-ID alone never produces 1.
    Dim n As Long, detail As String
    On Error GoTo Failed
    MBL_CompareExisting mail, path
    MBI_CompareContents = 1
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    If gFatal Then Err.Raise n, "MBI_CompareContents", detail
    reason = detail
    If gCancel Then MBI_CompareContents = 3: Exit Function
    If n = MB_ERR + 161 Then
        MBI_CompareContents = MBI_TimestampDifference(mail, path, reason)
    Else
        Select Case n - MB_ERR
            Case 158 To 160, 162 To 172, 176, 181, 182, 1120 To 1122
                MBI_CompareContents = 0
            Case Else
                MBI_CompareContents = 2
        End Select
    End If
End Function

Private Function MBI_CheckCandidate(ByVal mail As Outlook.MailItem, ByVal key As String, _
                                    ByRef path As String, ByRef root As String, _
                                    ByRef digest As String, ByRef bytes As Double, _
                                    ByRef reason As String) As Long
    Dim f As Variant, disk As Object, beforeDate As Date, result As Long, n As Long, detail As String
    On Error GoTo Failed
    f = MB_IndexFields(key)
    root = MBM_ResolvedRoot(MBX_IndexedOriginal(key, f), CStr(f(2)), MBX_IndexedRoot(key))
    path = MBX_IndexedPath(key, f, True)
    path = MBL_CanonicalLegacyPath(path, root, True)
    Set disk = gFSO.GetFile(path)
    bytes = CDbl(disk.Size): beforeDate = disk.DateLastModified
    If bytes <= 0 Or bytes <> MB_ByteCount(CStr(f(1))) Then _
        Err.Raise MB_ERR + 731, , "Indexed identity candidate size changed."
    digest = MBH_HashFile(path)
    If digest <> CStr(f(2)) Then Err.Raise MB_ERR + 732, , "Indexed identity candidate SHA-256 changed."
    result = MBI_CompareContents(mail, path, reason)
    If result = 3 Then MBI_CheckCandidate = 3: Exit Function
    ' Even an explicit content mismatch is trusted only while the indexed file
    ' still has its original bytes. This prevents a race from authorizing SaveAs.
    If MBH_HashFile(path) <> digest Then Err.Raise MB_ERR + 733, , "Identity candidate changed while compared."
    Set disk = gFSO.GetFile(path)
    If CDbl(disk.Size) <> bytes Or disk.DateLastModified <> beforeDate Then _
        Err.Raise MB_ERR + 734, , "Identity candidate size/time changed while compared."
    MBI_CheckCandidate = result
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    If gFatal Then Err.Raise n, "MBI_CheckCandidate", detail
    reason = "Indexed identity candidate needs review: " & detail
    If gCancel Then MBI_CheckCandidate = 3 Else MBI_CheckCandidate = 2
End Function

Private Function MBI_TryAlias(ByVal mail As Outlook.MailItem, ByVal q As Variant, _
                              ByVal revision As String, ByRef reason As String) As Long
    ' 0=no equal candidate (ordinary legacy/new-export policy still applies);
    ' 1=verified R alias, no mail SaveAs; 2=review, no SaveAs; 3=cancel.
    Dim bucket As String, candidates As Collection, key As Variant, meta As Variant, fields As Variant
    Dim seen As Object, unique As String, pass As Long, checked As Long, result As Long
    Dim path As String, root As String, digest As String, bytes As Double, candidateReason As String
    Dim matchedPath As String, matchedRoot As String, matchedHash As String, matchedKey As String
    Dim matchedBytes As Double, uncertain As String, conflict As Boolean, n As Long, detail As String
    Dim hint As Variant, sourceId As String, sourceHas As Boolean, sourceSent As String, plausible As Boolean
    Dim dates As Object, entries As Object, unknown As Object, sourceCandidates As Collection
    Dim scope As String, groups As Variant, group As Variant
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    reason = ""
    If gIdentityBuckets Is Nothing Then Exit Function
    bucket = MBI_SubjectKey(CStr(q(1)), mail.Subject)
    If Not gIdentityBuckets.Exists(bucket) Then Exit Function
    sourceId = MBL_InternetId(mail, sourceHas): sourceSent = MB_Stamp(mail.SentOn)
    Set dates = gIdentityLookupState("dates"): Set entries = gIdentityLookupState("entries")
    Set unknown = gIdentityLookupState("unknown")
    groups = MBI_QueryGroups(sourceId, sourceHas, sourceSent)
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For pass = 1 To 2
        ' Timestamp is a cheap ordering hint only. The fallback also handles
        ' timezone changes and old journals whose displayed local stamp differs.
        If pass = 1 Then
            scope = MBI_DateScope(bucket, CStr(q(6)))
            Set sourceCandidates = MBI_LookupGroup(dates, scope)
        Else
            scope = bucket & "|ALL"
            Set sourceCandidates = gIdentityBuckets(bucket)
        End If
        If Not MBI_EnsureLookup(scope, sourceCandidates) Then MBI_TryAlias = 3: Exit Function
        If unknown.Exists(scope) Then
            If Len(uncertain) = 0 Then uncertain = CStr(unknown(scope))
        End If
        For Each group In groups
            Set candidates = MBI_LookupGroup(entries, scope & "|" & CStr(group))
            For Each key In candidates
                If Not MB_ContinueWork() Then MBI_TryAlias = 3: Exit Function
                If gIndex.Exists(CStr(key)) Then
                    meta = MBX_Meta(CStr(key))
                    If CStr(meta(0)) <> MBX_StoreToken(CStr(q(1))) Then _
                        Err.Raise MB_ERR + 735, , "Identity candidate belongs to another store."
                    If (CStr(meta(1)) = CStr(q(6))) = (pass = 1) Then
                        fields = MB_IndexFields(CStr(key))
                        path = MBX_IndexedPath(CStr(key), fields)
                        unique = MB_LegacyOwnerKey(path, CStr(fields(2)))
                        If Not seen.Exists(unique) Then
                            seen.Add unique, True
                            ' Header inventory has no silent cap. Each indexed path
                            ' + hash is read once per run; repeated daily subjects
                            ' with different IDs do not consume the deep-check cap.
                            result = MBI_ReadHeader(CStr(key), hint, candidateReason)
                            If result = 3 Then MBI_TryAlias = 3: Exit Function
                            If result = 2 Then
                                If Len(uncertain) = 0 Then uncertain = candidateReason
                            Else
                                plausible = True
                                If sourceHas And CBool(hint(2)) And Len(sourceId) > 0 And Len(CStr(hint(0))) > 0 Then
                                    plausible = (StrComp(sourceId, CStr(hint(0)), vbBinaryCompare) = 0)
                                ElseIf Not sourceHas And Not CBool(hint(2)) Then
                                    plausible = (sourceSent = CStr(hint(1)))
                                End If
                                If plausible Then
                                    If checked >= MBI_MAX_CANDIDATES Then
                                        reason = "Plausible identity candidate limit reached; unvisited candidates remain, so automatic export is blocked."
                                        MBI_TryAlias = 2: Exit Function
                                    End If
                                    checked = checked + 1
                                    result = MBI_CheckCandidate(mail, CStr(key), path, root, digest, bytes, candidateReason)
                                    If result = 3 Then MBI_TryAlias = 3: Exit Function
                                    If result = 2 Then
                                        If Len(uncertain) = 0 Then uncertain = candidateReason
                                    ElseIf result = 1 Then
                                        If Len(matchedPath) = 0 Then
                                            matchedPath = path: matchedRoot = root: matchedHash = digest
                                            matchedBytes = bytes: matchedKey = CStr(key)
                                        ElseIf matchedHash <> digest Or matchedBytes <> bytes Then
                                            conflict = True
                                        ElseIf StrComp(path, matchedPath, vbTextCompare) < 0 Then
                                            matchedPath = path: matchedRoot = root: matchedKey = CStr(key)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next key
        Next group
        If conflict Then
            reason = "Multiple content-equal MSG candidates have different file bytes; preserve all and review."
            MBI_TryAlias = 2: Exit Function
        End If
        If Len(matchedPath) > 0 Then
            ' A full exact-date match proves this alias; other-date candidates
            ' need not be searched. Unknown candidates never become aliases.
            ' Other candidates may have taken a long time or yielded to Outlook.
            ' Recheck the selected reference, bytes AND content before commit.
            ' Deny write/delete on that selected file throughout the final
            ' comparator, source check and index writer invocation.
            guard = MBH_GuardFile(matchedPath)
            result = MBI_CheckCandidate(mail, matchedKey, path, root, digest, bytes, candidateReason)
            If result = 3 Then MBI_TryAlias = 3: GoTo AliasReturn
            If result <> 1 Then
                reason = "Chosen identity candidate no longer verifies: " & candidateReason
                MBI_TryAlias = 2: GoTo AliasReturn
            End If
            If StrComp(path, matchedPath, vbTextCompare) <> 0 Or _
               StrComp(root, matchedRoot, vbTextCompare) <> 0 Or _
               digest <> matchedHash Or bytes <> matchedBytes Then
                reason = "Chosen identity reference changed before its alias could be committed."
                MBI_TryAlias = 2: GoTo AliasReturn
            End If
            If Not MB_SourceStable(mail, q, revision) Then
                reason = "The source changed during restored-identity verification."
                MBI_TryAlias = 2: GoTo AliasReturn
            End If
            MB_RecordCommit "R", revision, matchedPath, matchedBytes, matchedHash, q, mail.Subject, _
                            0, matchedKey, MB_NEW_CONTENT_VALIDATION, matchedRoot
            MBH_CloseGuard guard
            gIdentityAliases = gIdentityAliases + 1
            MB_LogEvent "INFO", "IDENTITY_ALIAS_NO_SAVEAS", q(7), matchedPath, 0, _
                        "Current source content matched existing indexed bytes; reference=" & matchedKey
            MBI_TryAlias = 1: Exit Function
        End If
    Next pass
    If Len(uncertain) > 0 Then reason = uncertain: MBI_TryAlias = 2
AliasReturn:
    MBH_CloseGuard guard
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    If gFatal Then Err.Raise n, "MBI_TryAlias", detail
    reason = "Restored-identity lookup failed: " & detail
    If gCancel Then MBI_TryAlias = 3 Else MBI_TryAlias = 2
End Function

Private Function MBI_SyntheticComparatorTest(ByVal mail As Outlook.MailItem, ByVal path As String, _
                                            ByRef detail As String) As Boolean
    ' The caller owns this unsent test item and closes it with olDiscard.
    ' This checks the production comparator, not actual mailbox restore behavior.
    Dim reason As String, result As Long
    On Error GoTo Failed
    result = MBI_CompareContents(mail, path, reason)
    If result <> 1 Then Err.Raise MB_ERR + 736, , "Equal synthetic MSG was not accepted: " & reason
    mail.Body = mail.Body & vbCrLf & "Owned synthetic identity difference."
    result = MBI_CompareContents(mail, path, reason)
    If result <> 0 Then Err.Raise MB_ERR + 737, , "Different synthetic body was not rejected: " & reason
    detail = "PASS: same bounded content accepted; changed synthetic body rejected without alias."
    MBI_SyntheticComparatorTest = True
    Exit Function
Failed:
    detail = "FAIL: " & CStr(Err.Number) & " - " & Err.Description
End Function


' ======================== Adaptive backup coverage ==========================
' The UTC checkpoint is the START of a completed inventory/backup pass, never
' its end. A seven-day overlap repeats its boundary. There is NO elapsed-day cap
' that silently drops a long absence. Unknown source IDs and changed modification
' stamps are eligible at ANY mail date, including restored historical messages.
' Every AUTO pass inventories all declared folders WITHOUT a regional date query.
' Up to 500 recorded historical source identities also receive content comparison
' in a deterministic rotating audit. This is a per-run budget, not a deadline or
' proof against providers changing content without changing any metadata.
Private Sub MBW_Reset()
    Set gWindowKnown = Nothing: Set gWindowAudit = Nothing
    Set gWindowAuditOrder = Nothing: Set gWindowAuditSeen = Nothing
    Set gWindowAuditResults = Nothing: Set gWindowSourcePending = Nothing
    gWindowAuditPath = "": gWindowAuditCommitted = False
    gWindowAuditStateIssues = 0: gWindowAuditResultsWritten = 0: gWindowRetryQueued = 0
    gWindowAuditPass = 0: gWindowAuditLoaded = False
    gWindowStore = "": gWindowStartUTC = 0: gWindowPreviousUTC = 0
    gWindowContext = "": gWindowPath = "": gWindowCursor = "": gWindowNextCursor = ""
    gWindowFull = False: gWindowReady = False: gWindowInventoryComplete = False
    gWindowCommitted = False: gWindowNote = ""
    gWindowNew = 0: gWindowChanged = 0: gWindowRecent = 0: gWindowAuditQueued = 0
    gWindowFullQueued = 0: gWindowHistoricalSkipped = 0: gWindowBadRows = 0: gWindowPolicyQueued = 0
End Sub

Private Sub MBW_Start(ByVal storeId As String)
    gWindowStore = storeId
    Set gWindowKnown = CreateObject("Scripting.Dictionary")
    gWindowKnown.CompareMode = vbBinaryCompare
    Set gWindowAudit = CreateObject("Scripting.Dictionary")
    gWindowAudit.CompareMode = vbBinaryCompare
    Set gWindowAuditSeen = CreateObject("Scripting.Dictionary")
    gWindowAuditSeen.CompareMode = vbBinaryCompare
    Set gWindowAuditResults = CreateObject("Scripting.Dictionary")
    gWindowAuditResults.CompareMode = vbBinaryCompare
    Set gWindowSourcePending = CreateObject("Scripting.Dictionary")
    gWindowSourcePending.CompareMode = vbBinaryCompare
    Set gWindowAuditOrder = New Collection
    gWindowStartUTC = MBC_UtcNow()
    If Not MBC_ValidDate(gWindowStartUTC) Then Err.Raise MB_ERR + 742, , "UTC clock is outside the supported 1900-2099 range."
    gWindowContext = MBC_TimeContext()
End Sub

Private Sub MBW_CacheRecord(ByVal f As Variant)
    Dim kind As String
    If gWindowKnown Is Nothing Then Exit Sub
    If StrComp(CStr(f(7)), gWindowStore, vbBinaryCompare) <> 0 Then Exit Sub
    kind = CStr(f(1))
    If kind <> "M" And kind <> "N" And kind <> "L" And kind <> "R" Then Exit Sub
    ' f(4) is output-file bytes, NOT source MailItem.Size. Never compare those.
    ' Normalize only the fixed Gregorian component order used by older records.
    ' Unknown/corrupt forms remain different and force a content comparison.
    gWindowKnown(CStr(f(6))) = MBW_NormalizeKnownStamp(CStr(f(10)))
End Sub

Private Sub MBW_Prepare()
    Dim stream As Object, line As String, latest As Variant, f As Variant
    Dim saved As Boolean, n As Long, detail As String
    On Error GoTo Failed
    If gWindowKnown Is Nothing Then Err.Raise MB_ERR + 740, , "Adaptive inventory was not initialized."
    gWindowPath = MB_StatePath("coverage_" & _
        Left$(MBH_HashText(MB_EncodeFields(Array(gWindowStore, LCase$(gAutoAccount)))), 32) & ".tsv")
    MBY_NoReparse gWindowPath, False
    If gFSO.FileExists(gWindowPath) Then
        Set stream = gFSO.OpenTextFile(gWindowPath, 1, False, -1)
        Do Until stream.AtEndOfStream
            If Not MB_ContinueWork() Then stream.Close: Exit Sub
            line = stream.ReadLine
            If Len(line) > 0 Then
                If MBW_ReadCheckpoint(line, f) Then
                    latest = f: saved = True
                Else
                    gWindowBadRows = gWindowBadRows + 1
                End If
            End If
        Loop
        stream.Close: Set stream = Nothing
    End If
    If saved Then
        gWindowPreviousUTC = MBC_ParseUtcStamp(CStr(latest(4)))
        gWindowCursor = CStr(latest(7))
        If CStr(latest(5)) <> gWindowContext Or CStr(latest(6)) <> MBW_POLICY Then
            gWindowFull = True
            gWindowNote = "Time-zone rules or coverage policy changed: reconcile all current mail."
        ElseIf gWindowPreviousUTC > gWindowStartUTC Then
            gWindowFull = True
            gWindowNote = "Clock is earlier than the last checkpoint: reconcile all current mail."
        Else
            If CDbl(gWindowPreviousUTC) < 2# + MBW_OVERLAP_DAYS Then
                gWindowFull = True
                gWindowNote = "Coverage overlap reaches before the supported mail epoch: reconcile all current mail."
            Else
                gSince = MBC_UtcToLocal(DateAdd("d", -MBW_OVERLAP_DAYS, gWindowPreviousUTC))
                gWindowNote = "Last successful scan START in UTC minus 7 days; no maximum-gap cutoff."
            End If
        End If
    Else
        If gWindowKnown.Count > 0 Then
            gWindowFull = True
            gWindowNote = "No valid UTC checkpoint with existing indexed history: reconcile all current mail."
        ElseIf CDbl(gWindowStartUTC) < 2# + MBW_INITIAL_DAYS Then
            gWindowFull = True
            gWindowNote = "Initial range reaches before the supported mail epoch: reconcile all current mail."
        Else
            gSince = MBC_UtcToLocal(DateAdd("d", -MBW_INITIAL_DAYS, gWindowStartUTC))
            gWindowNote = "No UTC checkpoint: 30-day baseline plus every unknown/changed identity at any date."
        End If
    End If
    If gRecoveryForceFull Then
        gWindowFull = True
        gWindowNote = "Recovered control state or rebuilt MSG baseline: reconcile every current source before advancing coverage."
    End If
    If gWindowFull Then gSince = 0
    gUntil = MBC_UtcToLocal(gWindowStartUTC)
    MBW_LoadAuditState
    MBW_BuildAudit
    MB_LogEvent "INFO", "ADAPTIVE_RANGE", "", gWindowPath, 0, gWindowNote & _
        " UTC start=" & MBC_UtcStamp(gWindowStartUTC) & "; local since=" & MB_Stamp(gSince) & _
        "; audit identities=" & CStr(gWindowAudit.Count) & "; invalid checkpoint rows=" & CStr(gWindowBadRows)
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBW_Prepare", detail
End Sub

Private Function MBW_ReadCheckpoint(ByVal line As String, ByRef f As Variant) As Boolean
    Dim stamp As Date
    On Error GoTo Invalid
    f = MB_ReadCheckedLine(line, 9)
    If CStr(f(0)) <> "MBWINDOW1" Then Exit Function
    If StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Or _
       StrComp(CStr(f(2)), gWindowStore, vbBinaryCompare) <> 0 Or _
       StrComp(CStr(f(3)), gAutoAccount, vbTextCompare) <> 0 Then Exit Function
    stamp = MBC_ParseUtcStamp(CStr(f(4)))
    If Not MBC_ValidDate(stamp) Then Exit Function
    If Not MB_IsHexKey(CStr(f(5))) Then Exit Function
    If Len(CStr(f(7))) > 0 Then
        If Not MB_IsHexKey(CStr(f(7))) Then Exit Function
    End If
    If Len(CStr(f(6))) = 0 Or Len(CStr(f(8))) = 0 Then Exit Function
    MBW_ReadCheckpoint = True
Invalid:
End Function

Private Sub MBW_BuildAudit()
    Dim keys As Variant, i As Long, first As Long, count As Long, picked As Long
    gWindowAudit.RemoveAll
    Set gWindowAuditOrder = New Collection
    gWindowNextCursor = gWindowCursor
    count = gWindowKnown.Count
    If count = 0 Then Exit Sub
    keys = gWindowKnown.Keys
    MBW_SortKeys keys, LBound(keys), UBound(keys)
    first = LBound(keys)
    If Len(gWindowCursor) > 0 Then
        Do While first <= UBound(keys)
            If StrComp(CStr(keys(first)), gWindowCursor, vbBinaryCompare) > 0 Then Exit Do
            first = first + 1
        Loop
        If first > UBound(keys) Then first = LBound(keys)
    End If
    i = first
    Do
        gWindowAudit(CStr(keys(i))) = True
        gWindowAuditOrder.Add CStr(keys(i))
        gWindowNextCursor = CStr(keys(i))
        picked = picked + 1
        If picked >= MBW_AUDIT_LIMIT Or picked >= count Then Exit Do
        i = i + 1
        If i > UBound(keys) Then i = LBound(keys)
    Loop
End Sub

Private Sub MBW_SortKeys(ByRef keys As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, temporary As Variant
    i = first: j = last: pivot = CStr(keys((first + last) \ 2))
    Do While i <= j
        Do While StrComp(CStr(keys(i)), pivot, vbBinaryCompare) < 0: i = i + 1: Loop
        Do While StrComp(CStr(keys(j)), pivot, vbBinaryCompare) > 0: j = j - 1: Loop
        If i <= j Then
            temporary = keys(i): keys(i) = keys(j): keys(j) = temporary
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then MBW_SortKeys keys, first, j
    If i < last Then MBW_SortKeys keys, i, last
End Sub

Private Function MBW_ShouldQueue(ByVal sourceKey As String, ByVal observedMod As Date, _
    ByVal effectiveDate As Date, ByVal sinceDate As Date, ByVal untilDate As Date) As Boolean
    If gWindowKnown Is Nothing Then Err.Raise MB_ERR + 741, , "Adaptive source index is unavailable."
    MBW_ShouldQueue = True
    If Not gWindowAuditSeen Is Nothing Then
        If gWindowAudit.Exists(sourceKey) Then gWindowAuditSeen(sourceKey) = True
    End If
    If Not gWindowSourcePending Is Nothing Then
        If gWindowSourcePending.Exists(sourceKey) Then
            gWindowRetryQueued = gWindowRetryQueued + 1
            Exit Function
        End If
    End If
    If Not gWindowKnown.Exists(sourceKey) Then
        gWindowNew = gWindowNew + 1
    ElseIf CStr(gWindowKnown(sourceKey)) <> MB_Stamp(observedMod) Then
        gWindowChanged = gWindowChanged + 1
    ElseIf MBV_RequiresCurrentPolicy(sourceKey) Then
        gWindowPolicyQueued = gWindowPolicyQueued + 1
    ElseIf gWindowFull Then
        gWindowFullQueued = gWindowFullQueued + 1
    ElseIf gWindowAudit.Exists(sourceKey) Then
        gWindowAuditQueued = gWindowAuditQueued + 1
    ElseIf (effectiveDate >= sinceDate And effectiveDate <= untilDate) Or _
           (observedMod >= sinceDate And observedMod <= untilDate) Then
        gWindowRecent = gWindowRecent + 1
    Else
        MBW_ShouldQueue = False
        gWindowHistoricalSkipped = gWindowHistoricalSkipped + 1
    End If
End Function

Private Function MBW_RequireContentCheck(ByVal sourceKey As String) As Boolean
    If gMode <> "AUTO" Then Exit Function
    If gWindowFull Then MBW_RequireContentCheck = True: Exit Function
    If MBV_RequiresCurrentPolicy(sourceKey) Then MBW_RequireContentCheck = True: Exit Function
    If Not gWindowSourcePending Is Nothing Then
        If gWindowSourcePending.Exists(sourceKey) Then MBW_RequireContentCheck = True: Exit Function
    End If
    If Not gWindowAudit Is Nothing Then MBW_RequireContentCheck = gWindowAudit.Exists(sourceKey)
End Function

Private Function MBW_CanCommit() As Boolean
    If gMode <> "AUTO" Or Len(gWindowPath) = 0 Or gAutoSetupPhase <> "READY" Then Exit Function
    If Not gWindowInventoryComplete Then Exit Function
    If MBR_HasReview() Then Exit Function
    If gCancel Or gFatal Or Len(gCleanupIssue) > 0 Then Exit Function
    If gMailFailed <> 0 Or gMailChanged <> 0 Or gMailPending <> 0 Or gMailReview <> 0 Or _
       gMailPolicy <> 0 Or gContentReview <> 0 Or gAttFailed <> 0 Or gAttPending <> 0 Or _
       gFolderFailures <> 0 Or gScanChanged <> 0 Or gJournalIssues <> 0 Or gAutoFailureCount <> 0 Then Exit Function
    If gProcessed <> gPlanned Then Exit Function
    MBW_CanCommit = True
End Function

Private Sub MBW_PrepareCommit()
    gWindowReady = False
    If Not MBW_CanCommit() Then Exit Sub
    MB_FlushJournal
    If gFSO.FileExists(gJournalPath) Then MBM_FlushPath gJournalPath
    gWindowReady = True
End Sub

Private Sub MBW_CommitAfterClose()
    Dim stream As Object, payload As String, n As Long, detail As String
    If Not gWindowReady Or Not MBW_CanCommit() Then Exit Sub
    On Error GoTo Failed
    ' Caller still holds the archive/account locks; all journal/event streams
    ' are closed. A torn checkpoint never advances past its last checked row.
    MBY_NoReparse gWindowPath, False
    payload = MB_EncodeFields(Array("MBWINDOW1", gRoot, gWindowStore, gAutoAccount, _
        MBC_UtcStamp(gWindowStartUTC), gWindowContext, MBW_POLICY, gWindowCursor, gRunId))
    Set stream = gFSO.OpenTextFile(gWindowPath, 8, True, -1)
    stream.WriteLine ""
    MB_WriteCheckedLine stream, payload
    stream.Close: Set stream = Nothing
    MBM_FlushPath gWindowPath
    gWindowCommitted = True
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBW_CommitAfterClose", detail
End Sub

Private Function MBW_Summary() As String
    If Len(gWindowStore) = 0 Then Exit Function
    MBW_Summary = vbCrLf & "Adaptive coverage: " & gWindowNote & vbCrLf & _
        "UTC scan start: " & MBC_UtcStamp(gWindowStartUTC) & "; prior successful start: " & _
        MBC_UtcStamp(gWindowPreviousUTC) & vbCrLf & _
        "Inventory selections (exclusive): unknown=" & CStr(gWindowNew) & _
        "; modified=" & CStr(gWindowChanged) & "; validation-policy=" & CStr(gWindowPolicyQueued) & _
        "; context reconcile=" & CStr(gWindowFullQueued) & _
        "; rotating audit=" & CStr(gWindowAuditQueued) & "; recent=" & CStr(gWindowRecent) & vbCrLf & _
        "Older recorded unchanged items deferred: " & CStr(gWindowHistoricalSkipped) & _
        "; source-audit budget=" & CStr(MBW_AUDIT_LIMIT) & " identities/run (independent file audit also runs)" & vbCrLf & _
        "UTC coverage checkpoint advanced: " & CStr(gWindowCommitted) & _
        "; invalid checkpoint rows ignored: " & CStr(gWindowBadRows) & vbCrLf & _
        "Source-audit cursor committed independently: " & CStr(gWindowAuditCommitted) & _
        "; durable outcomes=" & CStr(gWindowAuditResultsWritten) & _
        "; outstanding source reviews=" & CStr(MBW_SourcePendingCount()) & _
        "; forced review retries=" & CStr(gWindowRetryQueued) & _
        "; invalid source-audit rows=" & CStr(gWindowAuditStateIssues)
End Function
' Source comparison progress is separate from mail coverage. A REVIEW result
' advances the audit only after its retry obligation was written and flushed.
' A cancelled/unvisited item never becomes a successful audit result.
Private Function MBW_NormalizeKnownStamp(ByVal value As String) As String
    Dim parsed As Date
    MBW_NormalizeKnownStamp = value
    On Error GoTo Unchanged
    parsed = MBC_ParseLegacyStamp(value)
    MBW_NormalizeKnownStamp = MB_Stamp(parsed)
Unchanged:
End Function

Private Function MBW_SourcePendingCount() As Long
    If Not gWindowSourcePending Is Nothing Then MBW_SourcePendingCount = gWindowSourcePending.Count
End Function

Private Function MBW_HasAuditReview() As Boolean
    MBW_HasAuditReview = (MBW_SourcePendingCount() > 0 Or gWindowAuditStateIssues > 0)
End Function

Private Sub MBW_LoadAuditState()
    Dim stream As Object, line As String, f As Variant, n As Long, detail As String
    Dim sawCursor As Boolean
    On Error GoTo Failed
    gWindowAuditLoaded = False
    gWindowAuditPath = MB_StatePath("source_audit_" & _
        Left$(MBH_HashText(MB_EncodeFields(Array(gWindowStore, LCase$(gAutoAccount)))), 32) & ".tsv")
    MBY_NoReparse gWindowAuditPath, False
    If Not gFSO.FileExists(gWindowAuditPath) Then
        gWindowAuditLoaded = True
        Exit Sub
    End If
    gWindowSourcePending.RemoveAll
    Set stream = gFSO.OpenTextFile(gWindowAuditPath, 1, False, -1)
    Do Until stream.AtEndOfStream
        If Not MB_ContinueWork() Then stream.Close: Exit Sub
        line = stream.ReadLine
        If Len(line) > 0 Then
            If sawCursor Then Err.Raise MB_ERR + 745, , "Unexpected data after source-audit snapshot terminator."
            If MBW_ReadAuditRow(line, f) Then
                Select Case CStr(f(5))
                    Case "CURSOR"
                        gWindowCursor = CStr(f(6))
                        sawCursor = True
                    Case "REVIEW"
                        gWindowSourcePending(CStr(f(4))) = CStr(f(6))
                    Case "OK"
                        If gWindowSourcePending.Exists(CStr(f(4))) Then gWindowSourcePending.Remove CStr(f(4))
                    Case "ABSENT"
                        ' Absence is an inventory observation, not proof that an
                        ' earlier content problem was repaired. Keep its review.
                End Select
            Else
                gWindowAuditStateIssues = gWindowAuditStateIssues + 1
                Err.Raise MB_ERR + 744, , "Source audit state is damaged; recover its last complete snapshot before resuming."
            End If
        End If
    Loop
    stream.Close: Set stream = Nothing
    If Not sawCursor Then Err.Raise MB_ERR + 746, , "Incomplete source-audit snapshot: terminal cursor is missing."
    gWindowAuditLoaded = True
    If gWindowAuditStateIssues > 0 Then MB_LogEvent "WARNING", "SOURCE_AUDIT_STATE_REVIEW", "", _
        gWindowAuditPath, 0, "Invalid audit-state rows retained; unresolved history may be incomplete."
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBW_LoadAuditState", detail
End Sub

Private Function MBW_ReadAuditRow(ByVal line As String, ByRef f As Variant) As Boolean
    Dim stamp As Date
    On Error GoTo Invalid
    f = MB_ReadCheckedLine(line, 9)
    If CStr(f(0)) <> "MBSOURCEAUDIT1" Then Exit Function
    If StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Or _
       StrComp(CStr(f(2)), gWindowStore, vbBinaryCompare) <> 0 Or _
       StrComp(CStr(f(3)), gAutoAccount, vbTextCompare) <> 0 Then Exit Function
    stamp = MBC_ParseUtcStamp(CStr(f(7)))
    If Not MBC_ValidDate(stamp) Or Len(CStr(f(8))) = 0 Then Exit Function
    Select Case CStr(f(5))
        Case "CURSOR"
            If Len(CStr(f(4))) > 0 Then Exit Function
            If Len(CStr(f(6))) > 0 Then
                If Not MB_IsHexKey(CStr(f(6))) Then Exit Function
            End If
        Case "OK", "REVIEW", "ABSENT"
            If Not MB_IsHexKey(CStr(f(4))) Then Exit Function
        Case Else
            Exit Function
    End Select
    MBW_ReadAuditRow = True
Invalid:
End Function

Private Sub MBW_RecordAuditResult(ByVal sourceKey As String, ByVal outcome As Long, _
                                  Optional ByVal contentLimited As Boolean = False)
    Dim detail As String, isSuccess As Boolean, relevant As Boolean
    If (gMode <> "AUTO" And gMode <> "REGISTER_LEGACY") Or Len(gWindowAuditPath) = 0 Then Exit Sub
    If outcome = 5 Then Exit Sub
    If gMode = "REGISTER_LEGACY" And outcome = 8 Then Exit Sub
    If Not MB_IsHexKey(sourceKey) Then Err.Raise MB_ERR + 743, , "Invalid source-audit identity."
    isSuccess = ((outcome = 1 Or outcome = 2) And Not contentLimited)
    relevant = Not isSuccess
    If gWindowAudit.Exists(sourceKey) Or gWindowSourcePending.Exists(sourceKey) Then relevant = True
    If Not relevant Then Exit Sub
    If isSuccess Then
        detail = "OK"
    Else
        detail = "REVIEW|outcome=" & CStr(outcome)
        If contentLimited Then detail = detail & "; special attachment/content proof remains limited"
    End If
    gWindowAuditResults(sourceKey) = detail
End Sub

Private Sub MBW_AppendAuditRow(ByVal stream As Object, ByVal sourceKey As String, _
                              ByVal state As String, ByVal detail As String)
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBSOURCEAUDIT1", gRoot, gWindowStore, _
        gAutoAccount, sourceKey, state, detail, MBC_UtcStamp(MBC_UtcNow()), gRunId))
End Sub

Private Sub MBW_FinishAudit()
    Dim stream As Object, key As Variant, value As String, cursor As String
    Dim n As Long, detail As String, completed As Boolean, reportPath As String
    Dim snapshot As Collection
    If (gMode <> "AUTO" And gMode <> "REGISTER_LEGACY") Or Len(gWindowAuditPath) = 0 Then Exit Sub
    If gWindowAuditResults Is Nothing Or gWindowAuditOrder Is Nothing Or Not gWindowAuditLoaded Then Exit Sub
    On Error GoTo Failed
    MB_EnsureDir CStr(gFSO.GetParentFolderName(gWindowAuditPath))
    MBY_NoReparse gWindowAuditPath, False
    gWindowAuditPass = gWindowAuditPass + 1
    reportPath = MBH_JoinPath(gRunDir, "source_audit_" & MBC_Digits(gWindowAuditPass, 4) & ".tsv")
    Set stream = gFSO.CreateTextFile(reportPath, False, True)
    For Each key In gWindowAuditResults.Keys
        value = CStr(gWindowAuditResults(key))
        If value = "OK" Then
            MBW_AppendAuditRow stream, CStr(key), "OK", "Current source content verified."
            If gWindowSourcePending.Exists(CStr(key)) Then gWindowSourcePending.Remove CStr(key)
        Else
            MBW_AppendAuditRow stream, CStr(key), "REVIEW", value
            gWindowSourcePending(CStr(key)) = value
        End If
        gWindowAuditResultsWritten = gWindowAuditResultsWritten + 1
    Next key
    cursor = gWindowCursor
    If gWindowInventoryComplete Then
        For Each key In gWindowAuditOrder
            completed = gWindowAuditResults.Exists(CStr(key))
            If Not completed And Not gWindowAuditSeen.Exists(CStr(key)) Then
                MBW_AppendAuditRow stream, CStr(key), "ABSENT", _
                    "Source absent from completed inventory; independent physical file audit remains applicable."
                gWindowAuditResultsWritten = gWindowAuditResultsWritten + 1
                completed = True
            End If
            If Not completed Then Exit For
            cursor = CStr(key)
        Next key
    End If
    stream.Close: Set stream = Nothing
    MBM_FlushPath reportPath
    ' The compact snapshot contains every unresolved identity plus the cursor.
    ' Atomic replacement cannot publish a cursor without its retry obligations.
    Set snapshot = New Collection
    For Each key In gWindowSourcePending.Keys
        snapshot.Add MBW_CheckedText(Array("MBSOURCEAUDIT1", gRoot, gWindowStore, gAutoAccount, _
            CStr(key), "REVIEW", CStr(gWindowSourcePending(key)), MBC_UtcStamp(MBC_UtcNow()), gRunId))
    Next key
    snapshot.Add MBW_CheckedText(Array("MBSOURCEAUDIT1", gRoot, gWindowStore, gAutoAccount, _
        "", "CURSOR", cursor, MBC_UtcStamp(MBC_UtcNow()), gRunId))
    MBS_WriteAtomic gWindowAuditPath, MBW_LinesText(snapshot)
    gWindowCursor = cursor
    gWindowAuditCommitted = True
    gWindowAuditResults.RemoveAll
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBW_FinishAudit", detail
End Sub

Private Function MBW_CheckedText(ByVal values As Variant) As String
    Dim payload As String
    payload = MB_EncodeFields(values)
    MBW_CheckedText = payload & vbTab & MBH_HashText(payload) & vbTab & "END"
End Function

Private Function MBW_LinesText(ByVal rows As Collection) As String
    Dim values() As String, i As Long
    If rows.Count = 0 Then Exit Function
    ReDim values(0 To rows.Count - 1)
    For i = 1 To rows.Count
        values(i - 1) = CStr(rows(i))
    Next i
    MBW_LinesText = Join(values, vbCrLf) & vbCrLf
End Function


' ================= Operating metadata / restart recovery ===================
Private Sub MBS_Reset()
    gStateRoot = "": gStateWasMissing = False: gStateMigrated = False
    gStageEnabled = False: Set gStageRecords = Nothing
    Set gStageExpected = CreateObject("Scripting.Dictionary")
    gStageExpected.CompareMode = vbTextCompare
    gStageMirror = "": gStageAccount = "": gStageReview = 0
    gStageRecovered = 0: gStageHeld = 0: gStageLedgerRepaired = 0: gStageProofCopies = 0
    gMaintenanceReady = False: gAutoMaintenance = False: gEarlyAccount = ""
End Sub

Private Function MB_StatePath(Optional ByVal relative As String = "") As String
    Dim value As String, part As Variant
    If Len(gRoot) = 0 Then Err.Raise MB_ERR + 900, , "State root is not initialized."
    value = MBH_JoinPath(gRoot, MBS_DataLeaf(gRoot))
    If Len(relative) > 0 Then
        For Each part In Split(relative, "\")
            value = MBH_JoinPath(value, CStr(part))
        Next part
    End If
    MB_StatePath = value
End Function

Private Function MBS_IsModernLayout(ByVal managedRoot As String) As Boolean
    managedRoot = MBH_NormalizePath(managedRoot)
    MBS_IsModernLayout = (StrComp(Mid$(managedRoot, InStrRev(managedRoot, "\") + 1), _
                                 MBP_MANAGED_LEAF, vbTextCompare) = 0)
End Function

Private Function MBS_DataLeaf(ByVal managedRoot As String) As String
    ' The immutable managed-root name selects its whole directory namespace.
    ' Deleting State/Recovery/settings cannot silently select a different path.
    If MBS_IsModernLayout(managedRoot) Then
        MBS_DataLeaf = MBS_DATA_LEAF
    Else
        ' Prior exactly owned custom roots also retain the historical layout.
        MBS_DataLeaf = MBS_DATA_LEAF_LEGACY
    End If
End Function

Private Sub MBS_ValidateLayout(ByVal managedRoot As String, ByVal fso As Object)
    Dim other As String, expected As String, folder As Object
    expected = MBS_DataLeaf(managedRoot)
    If expected = MBS_DATA_LEAF Then other = MBS_DATA_LEAF_LEGACY Else other = MBS_DATA_LEAF
    other = MBH_JoinPath(managedRoot, other)
    MBP_NoReparse other, fso
    If fso.FileExists(other) Then Err.Raise MB_ERR + 1210, , _
        "A conflicting operating-directory name is occupied by a file: " & other
    If fso.FolderExists(other) Then
        Set folder = fso.GetFolder(other)
        If folder.Files.Count > 0 Or folder.SubFolders.Count > 0 Then Err.Raise MB_ERR + 1210, , _
            "Both directory layouts or a mixed legacy layout need review. Existing state was retained: " & other
    End If
End Sub

Private Function MBS_Checked(ByVal values As Variant) As String
    Dim payload As String
    payload = MB_EncodeFields(values)
    MBS_Checked = payload & vbTab & MBH_HashText(payload) & vbTab & "END" & vbCrLf
End Function

Private Function MBS_OwnerEvidenceRow(ByVal line As String, ByVal rootPath As String, _
                                       ByVal kind As String, Optional ByVal account As String = "") As Boolean
    Dim row As Variant, boundAccount As String
    On Error GoTo Unknown
    Select Case kind
        Case "FORMAT"
            row = MB_ReadCheckedLine(line, 3)
            MBS_OwnerEvidenceRow = (CStr(row(0)) = "MBSYSTEM1" And _
                StrComp(MBH_NormalizePath(CStr(row(1))), rootPath, vbTextCompare) = 0)
        Case "STORE"
            row = MB_ReadCheckedLine(line, 4)
            MBS_OwnerEvidenceRow = (CStr(row(0)) = "MBSTORE1" And Len(CStr(row(2))) > 0 And _
                CStr(row(3)) = "SINGLE_DEFAULT_STORE" And _
                StrComp(MBH_NormalizePath(CStr(row(1))), rootPath, vbTextCompare) = 0)
        Case "ACCOUNT"
            row = MB_ReadCheckedLine(line, 4)
            If CStr(row(0)) <> "MBAC3" Or Len(CStr(row(1))) = 0 Then Exit Function
            If StrComp(MBH_NormalizePath(CStr(row(3))), rootPath, vbTextCompare) <> 0 Then Exit Function
            boundAccount = MBH_NormalizePath(CStr(row(2)))
        Case "RECOVERY"
            row = MB_ReadCheckedLine(line, 5)
            If CStr(row(0)) <> "MBRECOVERY1" Or Len(CStr(row(3))) = 0 Or CStr(row(4)) <> "1" Then Exit Function
            If StrComp(MBH_NormalizePath(CStr(row(1))), rootPath, vbTextCompare) <> 0 Then Exit Function
            boundAccount = MBH_NormalizePath(CStr(row(2)))
    End Select
    If Len(boundAccount) > 0 Then
        If MB_PathWithin(boundAccount, rootPath) Or MBL_IsFileSystemRoot(boundAccount) Then Exit Function
        If Len(account) > 0 Then
            If StrComp(boundAccount, account, vbTextCompare) <> 0 Then Exit Function
        End If
        MBS_OwnerEvidenceRow = True
    End If
Unknown:
End Function

Private Function MBS_OwnerEvidenceFile(ByVal path As String, ByVal rootPath As String, _
                                        ByVal kind As String, ByVal fso As Object, _
                                        Optional ByVal account As String = "") As Boolean
    On Error GoTo Unknown
    If Not fso.FileExists(path) Then Exit Function
    MBS_OwnerEvidenceFile = MBS_OwnerEvidenceRow(MBP_ReadOneLine(path, fso), rootPath, kind, account)
Unknown:
End Function

Private Function MBS_OwnerEvidenceMap(ByVal path As String, ByVal rootPath As String, ByVal fso As Object) As Boolean
    Dim stream As Object, line As String
    On Error GoTo Unknown
    If Not fso.FileExists(path) Then Exit Function
    MBP_NoReparse path, fso
    If CDbl(fso.GetFile(path).Size) Mod 2 <> 0 Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            If MBS_OwnerEvidenceRow(line, rootPath, "ACCOUNT") Then
                MBS_OwnerEvidenceMap = True
                Exit Do
            End If
        End If
    Loop
Unknown:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
End Function

Private Function MBS_CanRecoverOwner(ByVal rootPath As String, ByVal fso As Object) As Boolean
    Dim marker As String, child As Object, parent As String, account As Variant, saved As String
    On Error GoTo Unknown
    rootPath = MBH_NormalizePath(rootPath)
    MBP_NoReparse rootPath, fso
    ' Surviving exact checked bindings establish ownership; names never do.
    If MBS_OwnerEvidenceFile(MBH_JoinPath(rootPath, "operating-format.tsv"), rootPath, "FORMAT", fso) Then GoTo Proven
    If MBS_OwnerEvidenceFile(MBH_JoinPath(rootPath, "operating-store.tsv"), rootPath, "STORE", fso) Then GoTo Proven
    saved = MBP_GetSettingText("OperatingStore_" & MBH_HashText(LCase$(rootPath)), "")
    If Len(saved) > 0 Then
        If MBS_OwnerEvidenceRow(saved, rootPath, "STORE") Then GoTo Proven
    End If
    For Each account In Array(MBS_DATA_LEAF, MBS_DATA_LEAF_LEGACY, "_state")
        marker = MBH_JoinPath(MBH_JoinPath(rootPath, CStr(account)), MBY_MAP_FILE)
        If MBS_OwnerEvidenceMap(marker, rootPath, fso) Then GoTo Proven
    Next account
    ' A per-user locator or HKCU hint may name an external account. Its exact
    ' account/recovery marker must corroborate it before recreating owner.txt.
    account = MBP_OwnerAccountHint(Application.Session.DefaultStore.StoreID, rootPath)
    If Len(CStr(account)) > 0 Then
        If MBS_OwnerEvidenceFile(MBH_JoinPath(CStr(account), MBY_ACCOUNT_MARKER), rootPath, "ACCOUNT", fso, CStr(account)) Then GoTo Proven
        If MBS_OwnerRecoveryEvidence(CStr(account), rootPath, fso) Then GoTo Proven
    End If
    parent = CStr(fso.GetParentFolderName(rootPath))
    If Len(parent) = 0 Or Not fso.FolderExists(parent) Then Exit Function
    If MBS_OwnerEvidenceFile(MBH_JoinPath(parent, MBY_ACCOUNT_MARKER), rootPath, "ACCOUNT", fso, parent) Then GoTo Proven
    If MBS_OwnerRecoveryEvidence(parent, rootPath, fso) Then GoTo Proven
    For Each child In fso.GetFolder(parent).SubFolders
        If (CLng(child.Attributes) And 1024) = 0 Then
            account = CStr(child.Path)
            If MBS_OwnerEvidenceFile(MBH_JoinPath(CStr(account), MBY_ACCOUNT_MARKER), rootPath, "ACCOUNT", fso, CStr(account)) Then GoTo Proven
            If MBS_OwnerRecoveryEvidence(CStr(account), rootPath, fso) Then GoTo Proven
        End If
    Next child
    Exit Function
Proven:
    MBS_CanRecoverOwner = True
Unknown:
End Function

Private Function MBS_OwnerRecoveryEvidence(ByVal accountPath As String, ByVal rootPath As String, ByVal fso As Object) As Boolean
    Dim leaf As Variant, marker As String
    ' A broad sibling scan borrows only an exact checked root/account binding.
    ' Conflicts in the actually selected account are checked by RecoveryRoot.
    For Each leaf In Array(MBR_FOLDER, MBR_FOLDER_LEGACY)
        marker = MBH_JoinPath(MBH_JoinPath(accountPath, CStr(leaf)), MBR_OWNER_FILE)
        If MBS_OwnerEvidenceFile(marker, rootPath, "RECOVERY", fso, accountPath) Then
            MBS_OwnerRecoveryEvidence = True
            Exit Function
        End If
    Next leaf
End Function

Private Sub MBS_WriteAtomic(ByVal path As String, ByVal contents As String)
    ' Caller owns the metadata destination and holds the corresponding run lock.
    ' A same-directory, fully flushed temporary file replaces only metadata.
    Dim stage As String, stream As Object, n As Long, detail As String, flags As Long
    Dim digest As String, created As Boolean
    On Error GoTo Failed
    path = MBH_NormalizePath(path)
    If Not MB_PathWithin(path, gRoot) Then Err.Raise MB_ERR + 901, , "Atomic state destination is outside the owned root."
    If LCase$(Right$(path, 4)) = ".msg" Then Err.Raise MB_ERR + 901, , "Atomic metadata writer cannot replace MSG files."
    MBY_NoReparse path, False
    MBH_EnsureDirectory CStr(gFSO.GetParentFolderName(path))
    stage = MBH_JoinPath(CStr(gFSO.GetParentFolderName(path)), "_mb_atomic_" & gFSO.GetTempName)
    Set stream = gFSO.CreateTextFile(stage, False, True): created = True
    stream.Write contents
    stream.Close: Set stream = Nothing
    MBM_FlushPath stage
    digest = MBH_HashFile(stage)
    flags = 8
    If gFSO.FileExists(path) Then flags = flags Or 1
    If MBM_MoveFileExW(StrPtr(stage), StrPtr(path), flags) = 0 Then
        Err.Raise MB_ERR + 902, , "Atomic metadata publication failed. Win32=" & CStr(Err.LastDllError)
    End If
    created = False
    MBM_FlushPath path
    If MBH_HashFile(path) <> digest Then Err.Raise MB_ERR + 903, , "Published metadata hash differs."
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    ' Retain any failed owned temp for diagnosis; never remove a published file.
    On Error GoTo 0
    Err.Raise n, "MBS_WriteAtomic", detail
End Sub

Private Sub MBS_Initialize()
    Dim anchor As String, prior As Boolean, oldDir As String, f As Object, target As String
    Dim stream As Object, row As Variant, saved As String, key As String, oldHash As String
    MBS_ValidateLayout gRoot, gFSO
    gStateRoot = MB_StatePath()
    anchor = MBH_JoinPath(gRoot, "operating-format.tsv")
    key = "OperatingFormat_" & MBH_HashText(LCase$(gRoot))
    saved = MBP_GetSettingText(key, "")
    If Len(saved) > 0 Then
        If StrComp(saved, gRoot, vbTextCompare) <> 0 Then Err.Raise MB_ERR + 904, , "Operating state binding mismatch."
        prior = True
    End If
    If gFSO.FileExists(anchor) Then
        Set stream = gFSO.OpenTextFile(anchor, 1, False, -1)
        row = MB_ReadCheckedLine(stream.ReadLine, 3)
        If Not stream.AtEndOfStream Then Err.Raise MB_ERR + 905, , "Operating marker has extra data."
        stream.Close: Set stream = Nothing
        If CStr(row(0)) <> "MBSYSTEM1" Or StrComp(CStr(row(1)), gRoot, vbTextCompare) <> 0 Then _
            Err.Raise MB_ERR + 906, , "Operating marker belongs to another root."
        prior = True
    End If
    MBY_NoReparse gStateRoot, False
    If gFSO.FolderExists(gStateRoot) Then
        ' Deleting the two activation settings must never roll newer active
        ' records back to retained _state. Even a partial first migration is
        ' existing evidence: recovery preserves/repairs it instead of overlaying it.
        If gFSO.GetFolder(gStateRoot).Files.Count > 0 Or gFSO.GetFolder(gStateRoot).SubFolders.Count > 0 Then prior = True
    End If
    gStateWasMissing = (prior And Not gFSO.FileExists(MB_StatePath("journal.tsv")))
    MBH_EnsureDirectory gStateRoot
    ' Old metadata is copied once, preserving original absolute root bindings.
    ' The activation marker prevents an old _state copy becoming a false latest
    ' state after accidental deletion of the new operating directory.
    oldDir = MBH_JoinPath(gRoot, "_state")
    If Not prior And gFSO.FolderExists(oldDir) Then
        MBY_NoReparse oldDir, True
        For Each f In gFSO.GetFolder(oldDir).Files
            If LCase$(gFSO.GetExtensionName(CStr(f.Name))) = "tsv" Then
                MBY_NoReparse CStr(f.Path), True
                oldHash = MBH_HashFile(CStr(f.Path))
                target = MB_StatePath(CStr(f.Name))
                If gFSO.FileExists(target) Then
                    MBY_NoReparse target, True
                    If MBH_HashFile(target) <> oldHash Then _
                        Err.Raise MB_ERR + 907, , "An active metadata file appeared during migration; neither copy was overwritten."
                End If
                If Not gFSO.FileExists(target) Then
                    gFSO.CopyFile CStr(f.Path), target, False
                    MBM_FlushPath target
                    If MBH_HashFile(target) <> oldHash Or MBH_HashFile(CStr(f.Path)) <> oldHash Then _
                        Err.Raise MB_ERR + 907, , "Old-state migration copy differs."
                End If
            End If
        Next f
        gStateMigrated = True
    End If
    If Not gFSO.FileExists(anchor) Then _
        MBS_WriteAtomic anchor, MBS_Checked(Array("MBSYSTEM1", gRoot, MBC_UtcStamp(MBC_UtcNow())))
    MBP_SaveSettingText key, gRoot
    If Not gFSO.FileExists(MB_StatePath("README.txt")) Then
        MBS_WriteAtomic MB_StatePath("README.txt"), _
            "Outlook Mail Backup operating files. Keep this folder with your MSG archive." & vbCrLf & _
            "journal.tsv: verified file identities and hashes; moves.tsv: path redirects." & vbCrLf & _
            "account-roots.tsv: remembered account folders; coverage/audit TSVs: progress." & vbCrLf & _
            "runs: per-run reports and fixed plans. The independent recovery directory in the account" & vbCrLf & _
            "folder contains independent checked recovery generations. Re-run the ONE" & vbCrLf & _
            "macro after accidental deletion. MSG-only recovery establishes a new baseline;" & vbCrLf & _
            "it cannot prove hashes or history that no longer exist. Never alternate older" & vbCrLf & _
            "writer macros on this archive. The old _state folder is retained as history." & vbCrLf
    End If
End Sub

Private Function MBS_AccountSettingKey(ByVal storeId As String, Optional ByVal managedRoot As String = "") As String
    If Len(managedRoot) = 0 Then managedRoot = gBaseRoot
    MBS_AccountSettingKey = "AccountRoot_" & MBH_HashText(MB_EncodeFields(Array(LCase$(managedRoot), LCase$(storeId))))
End Function

Private Sub MBS_RememberAccount(ByVal storeId As String, ByVal root As String)
    MBP_RememberLocatedAccount storeId, root
    MBP_SaveSettingText MBS_AccountSettingKey(storeId), root
End Sub

Private Function MBS_EarlyAccount(ByVal storeId As String) As String
    Dim candidate As String, key As Variant, n As Long, reason As String, value As Variant
    Dim remembered As Object, candidates As Object, exactBinding As Boolean
    Set remembered = CreateObject("Scripting.Dictionary")
    remembered.CompareMode = vbTextCompare
    Set candidates = CreateObject("Scripting.Dictionary")
    candidates.CompareMode = vbTextCompare
    candidate = MBP_GetSettingText(MBS_AccountSettingKey(storeId), "")
    If Len(candidate) > 0 Then remembered(candidate) = True
    candidate = MBP_LocatedAccount(storeId, gBaseRoot)
    If Len(candidate) > 0 Then remembered(candidate) = True
    On Error Resume Next
    Err.Clear
    MBY_LoadRoots
    If Not mbyRoots Is Nothing Then
        If mbyRoots.Exists(storeId) Then remembered(CStr(mbyRoots(storeId))) = True
    End If
    Err.Clear
    On Error GoTo 0
    For Each value In remembered.Keys
        If MBY_AccountBindingUsable(CStr(value), storeId, exactBinding) Then _
            candidates(MBH_NormalizePath(CStr(value))) = True
    Next value
    ' An agreed surviving remembered mapping outranks old leftover markers.
    ' Discover exact markers only when no usable remembered account remains.
    If candidates.Count = 0 Then MBY_DiscoverBoundAccounts storeId, candidates
    If gCancel Or gFatal Then Exit Function
    candidate = ""
    If candidates.Count = 1 Then
        For Each value In candidates.Keys
            candidate = CStr(value)
        Next value
    ElseIf candidates.Count = 0 And remembered.Count = 0 Then
        candidate = MBY_InferAccountRoot(storeId)
    ElseIf candidates.Count > 1 Then
        MsgBox "More than one existing account folder is bound to this store and managed root." & vbCrLf & _
               "Choose the account archive to use; no folder will be chosen by display name.", vbInformation, "Choose account folder"
    End If
    Do
        If gCancel Or gFatal Then Exit Function
        If Len(candidate) > 0 Then
            On Error Resume Next
            Err.Clear
            candidate = MBH_NormalizePath(candidate)
            MBY_NoReparse candidate, True
            n = Err.Number: reason = Err.Description
            On Error GoTo 0
            If n = 0 Then
                If Not gFSO.FolderExists(candidate) Then n = MB_ERR + 908: reason = "Account path is not a folder."
                If MB_PathWithin(candidate, gRoot) Then n = MB_ERR + 908: reason = "Account is inside the managed state tree."
                If MBL_IsFileSystemRoot(candidate) Then n = MB_ERR + 908: reason = "Select an account folder, not a drive root."
                If Len(candidate) > 80 Then n = MB_ERR + 908: reason = "Account root exceeds 80 UTF-16 units."
            End If
            If n = 0 Then
                On Error Resume Next
                Err.Clear
                MBY_CheckMarker candidate, storeId, False
                MBY_CheckRecoveryAccountBinding candidate, storeId
                MBY_CheckNestedMarkers candidate
                If Not mbyRoots Is Nothing Then
                    For Each key In mbyRoots.Keys
                        If StrComp(CStr(key), storeId, vbTextCompare) <> 0 Then
                            If MB_PathWithin(candidate, CStr(mbyRoots(key))) Or MB_PathWithin(CStr(mbyRoots(key)), candidate) Then _
                                Err.Raise MB_ERR + 923, , "Selected account overlaps another recorded account."
                        End If
                    Next key
                End If
                n = Err.Number: reason = Err.Description
                On Error GoTo 0
            End If
            If gCancel Or gFatal Then Exit Function
            If n = 0 Then
                MBM_AcquireAccountLock candidate
                MBS_EarlyAccount = candidate
                Exit Function
            End If
            MsgBox "The remembered account folder needs review." & vbCrLf & candidate & vbCrLf & reason, vbExclamation
        End If
        candidate = MBP_PickExistingFolder("Choose the existing account folder for backup/recovery. MSG files will be preserved.", candidate)
        If Len(candidate) = 0 Then gCancel = True: Exit Function
    Loop
End Function

Private Function MBS_PlanPathAllowed(ByVal path As String) As Boolean
    MBS_PlanPathAllowed = MB_PathWithin(path, MB_StatePath("runs")) Or _
                          MB_PathWithin(path, MBH_JoinPath(gRoot, "runs"))
End Function

Private Sub MBS_StageAppend(ByVal stage As String, ByVal final As String, ByVal phase As String, _
                            ByVal digest As String, ByVal bytes As Double, Optional ByVal held As String = "")
    Dim row As Variant, old As Variant, sequence As Long, payload As String, path As Variant, stream As Object
    Dim n As Long, detail As String
    If Not gStageEnabled Then Exit Sub
    On Error GoTo Failed
    sequence = 1
    If gStageRecords.Exists(stage) Then
        old = Split(CStr(gStageRecords(stage)), vbTab)
        sequence = CLng(old(8)) + 1
    End If
    row = Array("MBSTAGE1", gRoot, stage, final, phase, digest, Format$(bytes, "0"), held, CStr(sequence), MBC_UtcStamp(MBC_UtcNow()))
    payload = MB_EncodeFields(row)
    For Each path In Array(MB_StatePath("stages.tsv"), gStageMirror)
        MBY_NoReparse CStr(path), False
        Set stream = gFSO.OpenTextFile(CStr(path), 8, True, -1)
        stream.WriteLine "": MB_WriteCheckedLine stream, payload
        stream.Close: Set stream = Nothing
        MBM_FlushPath CStr(path)
        ' Never regress behind a transition already durable in one copy.
        gStageRecords(stage) = payload
        If phase = "PUBLISHED" Or phase = "RECOVERED" Then
            If MB_IsHexKey(digest) Then
                gStageExpected(final & "|" & digest) = MB_EncodeFields(Array( _
                    final, Format$(bytes, "0"), digest, gStageAccount, "VERIFIED_PUBLICATION"))
            End If
        End If
    Next path
    gStageRecords(stage) = payload
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    gFatal = True
    gFatalText = "Cannot persist stage recovery intent: " & detail
    Err.Raise n, "MBS_StageAppend", gFatalText
End Sub

Private Sub MBS_StageReserve(ByVal stage As String, ByVal final As String)
    If Not gStageEnabled Then Exit Sub
    ' Only durable mail/attachment outputs need recovery. Read-only comparison
    ' scratch stages under private temporary directories are not archive files.
    If Not MB_PathWithin(final, gStageAccount) Then Exit Sub
    If Len(gStageMirror) > 0 Then
        If MB_PathWithin(final, CStr(gFSO.GetParentFolderName(gStageMirror))) Then Exit Sub
    End If
    If Left$(LCase$(CStr(gFSO.GetFileName(final))), 4) = "_mb_" Then Exit Sub
    MBS_StageAppend stage, final, "RESERVE", "", 0
End Sub

Private Sub MBS_StagePublished(ByVal stage As String, ByVal final As String, ByVal digest As String, ByVal bytes As Double)
    If Not gStageEnabled Then Exit Sub
    If Not gStageRecords.Exists(stage) Then Exit Sub
    MBS_StageAppend stage, final, "PUBLISHED", digest, bytes
End Sub

Private Sub MBS_StagePrepared(ByVal stage As String, ByVal final As String, ByVal digest As String, ByVal bytes As Double)
    If Not gStageEnabled Then Exit Sub
    If Not gStageRecords.Exists(stage) Then Exit Sub
    If Not MB_IsHexKey(digest) Or bytes < 0 Then Err.Raise MB_ERR + 925, , "Invalid prepared publication proof."
    MBS_StageAppend stage, final, "PREPARED", digest, bytes
End Sub

Private Sub MBS_StageAborted(ByVal stage As String)
    Dim row As Variant
    If Not gStageEnabled Then Exit Sub
    If Not gStageRecords.Exists(stage) Then Exit Sub
    row = MBS_StageFields(CStr(gStageRecords(stage)))
    MBS_StageAppend stage, CStr(row(3)), "ABORTED", "", 0
End Sub

Private Function MBS_StageFields(ByVal payload As String) As Variant
    Dim row As Variant, i As Long
    row = Split(payload, vbTab)
    If UBound(row) <> 9 Then Err.Raise MB_ERR + 909, , "Invalid stage payload."
    For i = 0 To 9: row(i) = MB_Unescape(CStr(row(i))): Next i
    MBS_StageFields = row
End Function

Private Sub MBS_LoadStages(ByVal path As String, ByVal needsMirror As Object, ByVal needsPrimary As Object)
    Dim stream As Object, line As String, row As Variant, old As Variant, payload As String
    Dim stage As String, final As String, sequence As Double, n As Long, reason As String, isMirror As Boolean
    If Not gFSO.FileExists(path) Then Exit Sub
    MBY_NoReparse path, True
    isMirror = (StrComp(path, gStageMirror, vbTextCompare) = 0)
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            On Error Resume Next
            Err.Clear
            row = MB_ReadCheckedLine(line, 10)
            n = Err.Number: reason = Err.Description
            On Error GoTo 0
            If n <> 0 Then
                ' A bad reservation never authorizes touching any candidate file.
                gStageReview = gStageReview + 1
                MB_LogEvent "WARNING", "STAGE_LEDGER_REVIEW", "", path, n, reason
            Else
                If CStr(row(0)) <> "MBSTAGE1" Or StrComp(CStr(row(1)), gRoot, vbTextCompare) <> 0 Then _
                    Err.Raise MB_ERR + 910, , "Stage ledger binding mismatch."
                stage = MBH_NormalizePath(CStr(row(2))): final = MBH_NormalizePath(CStr(row(3)))
                If Not MB_PathWithin(final, gStageAccount) Then
                    If StrComp(path, gStageMirror, vbTextCompare) = 0 Then _
                        Err.Raise MB_ERR + 911, , "Mirrored stage belongs to another account."
                    GoTo NextStageRow
                End If
                If StrComp(gFSO.GetParentFolderName(stage), gFSO.GetParentFolderName(final), vbTextCompare) <> 0 Then _
                    Err.Raise MB_ERR + 911, , "Stage ledger path is outside its account/same directory."
                If Left$(LCase$(gFSO.GetFileName(stage)), 4) <> "~mb_" Then _
                    Err.Raise MB_ERR + 912, , "Stage reservation name is invalid."
                sequence = MB_ByteCount(CStr(row(8)))
                If sequence < 1 Or sequence > 2000000000# Then Err.Raise MB_ERR + 913, , "Invalid stage sequence."
                If CStr(row(4)) = "PREPARED" Or CStr(row(4)) = "PUBLISHED" Or CStr(row(4)) = "RECOVERED" Then
                    If Not MB_IsHexKey(CStr(row(5))) Then Err.Raise MB_ERR + 925, , "Stage proof digest is invalid."
                    If MB_ByteCount(CStr(row(6))) < 0 Then Err.Raise MB_ERR + 925, , "Stage proof byte count is invalid."
                End If
                ' Historical successful publication is independent of the latest
                ' cleanup phase. A later ABORTED row cannot erase expected bytes.
                ' Old PREPARED-only then ABORTED never becomes such an obligation.
                If CStr(row(4)) = "PUBLISHED" Or CStr(row(4)) = "RECOVERED" Then
                    gStageExpected(final & "|" & CStr(row(5))) = MB_EncodeFields(Array( _
                        final, CStr(row(6)), CStr(row(5)), gStageAccount, "VERIFIED_PUBLICATION"))
                End If
                payload = MB_EncodeFields(row)
                If gStageRecords.Exists(stage) Then
                    old = MBS_StageFields(CStr(gStageRecords(stage)))
                    If sequence = MB_ByteCount(CStr(old(8))) And payload <> CStr(gStageRecords(stage)) Then _
                        Err.Raise MB_ERR + 914, , "Conflicting stage recovery records."
                    If isMirror Then
                        If sequence >= MB_ByteCount(CStr(old(8))) Then
                            If needsMirror.Exists(stage) Then needsMirror.Remove stage
                        End If
                        If sequence > MB_ByteCount(CStr(old(8))) Then needsPrimary(stage) = True
                    End If
                    If sequence > MB_ByteCount(CStr(old(8))) Then gStageRecords(stage) = payload
                Else
                    gStageRecords(stage) = payload
                    If isMirror Then needsPrimary(stage) = True
                End If
                If Not isMirror Then needsMirror(stage) = True
            End If
        End If
NextStageRow:
    Loop
    stream.Close
End Sub

Private Sub MBS_CopyStageProofs(ByVal path As String, ByVal pending As Object)
    Dim stream As Object, key As Variant, copied As Long, n As Long, detail As String
    If pending.Count = 0 Then Exit Sub
    On Error GoTo Failed
    MBY_NoReparse path, False
    Set stream = gFSO.OpenTextFile(path, 8, True, -1)
    For Each key In pending.Keys
        If Not MB_ContinueWork() Then Exit For
        ' Preserve the exact checked transition; this is replica repair, not a
        ' new phase or sequence. A later lost peer must not erase proof observed
        ' from the surviving copy during this successful run.
        stream.WriteLine ""
        MB_WriteCheckedLine stream, CStr(gStageRecords(key))
        copied = copied + 1
    Next key
    stream.Close: Set stream = Nothing
    MBM_FlushPath path
    gStageProofCopies = gStageProofCopies + copied
    MB_LogEvent "INFO", "STAGE_PROOF_COPY_REPAIRED", "", path, 0, "Verified transitions copied=" & CStr(copied)
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    gFatal = True: gFatalText = "Cannot retain independent stage proof: " & detail
    Err.Raise n, "MBS_CopyStageProofs", gFatalText
End Sub

Private Function MBS_StageReadable(ByVal path As String, ByRef digest As String, ByRef bytes As Double) As Boolean
    Dim obj As Object, before As String, probe As String, readPath As String, copied As Boolean
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Invalid
    MBY_NoReparse path, True
    bytes = CDbl(gFSO.GetFile(path).Size)
    If bytes <= 0 Then Exit Function
    guard = MBH_GuardFile(path)
    before = MBH_HashFile(path)
    readPath = path
    If LCase$(Right$(path, 4)) <> ".msg" Then
        probe = MBH_JoinPath(gRunDir, "_mb_probe_" & gFSO.GetTempName & ".msg")
        gFSO.CopyFile path, probe, False: copied = True
        readPath = probe
    End If
    Set obj = Application.Session.OpenSharedItem(readPath)
    If Not TypeOf obj Is Outlook.MailItem Then GoTo Invalid
    obj.Close olDiscard: Set obj = Nothing
    digest = MBH_HashFile(path)
    If digest <> before Or CDbl(gFSO.GetFile(path).Size) <> bytes Then GoTo Invalid
    MBH_CloseGuard guard
    If copied Then gFSO.DeleteFile probe, False: copied = False
    MBS_StageReadable = True
    Exit Function
Invalid:
    On Error Resume Next
    If Not obj Is Nothing Then obj.Close olDiscard
    MBH_CloseGuard guard
    If copied Then gFSO.DeleteFile probe, False
    On Error GoTo 0
End Function

Private Sub MBS_RecoverStage(ByVal row As Variant)
    Dim stage As String, final As String, held As String, actual As String, phase As String
    Dim digest As String, bytes As Double, target As String, n As Long, detail As String
    On Error GoTo Review
    stage = CStr(row(2)): final = CStr(row(3)): phase = CStr(row(4)): held = CStr(row(7))
    If phase = "PUBLISHED" Or phase = "RECOVERED" Or phase = "ABORTED" Or phase = "CANDIDATE" Then Exit Sub
    If phase <> "RESERVE" And phase <> "PREPARED" And phase <> "HELD" And phase <> "HOLD_PREPARE" Then _
        Err.Raise MB_ERR + 915, , "Unknown stage state."
    If phase = "HOLD_PREPARE" Then
        If Len(held) = 0 Then Err.Raise MB_ERR + 916, , "Prepared held path is missing."
        If StrComp(held, stage & ".held", vbTextCompare) <> 0 Then Err.Raise MB_ERR + 916, , "Invalid prepared held path."
        MBY_NoReparse stage, False: MBY_NoReparse held, False: MBY_NoReparse final, False
        If gFSO.FileExists(stage) Then
            If gFSO.FileExists(held) Then Err.Raise MB_ERR + 917, , "Both unfinished and held copies exist."
            If MBH_HashFile(stage) <> CStr(row(5)) Or CDbl(gFSO.GetFile(stage).Size) <> MB_ByteCount(CStr(row(6))) Then _
                Err.Raise MB_ERR + 918, , "Unfinished stage changed before preservation."
            If MBM_MoveFileExW(StrPtr(stage), StrPtr(held), 8) = 0 Then Err.Raise MB_ERR + 919, , "Cannot finish held-file move."
        End If
        If Not gFSO.FileExists(held) Then Err.Raise MB_ERR + 918, , "Held-file intent has no preserved bytes."
        If MBH_HashFile(held) <> CStr(row(5)) Or CDbl(gFSO.GetFile(held).Size) <> MB_ByteCount(CStr(row(6))) Then _
            Err.Raise MB_ERR + 920, , "Preserved stage differs from its intent."
        MBS_StageAppend stage, final, "HELD", CStr(row(5)), MB_ByteCount(CStr(row(6))), held
        phase = "HELD"
    End If
    actual = stage
    If phase = "HELD" Then
        If Len(held) = 0 Then Err.Raise MB_ERR + 916, , "Held stage path is missing."
        If StrComp(gFSO.GetParentFolderName(held), gFSO.GetParentFolderName(stage), vbTextCompare) <> 0 Then _
            Err.Raise MB_ERR + 916, , "Held file moved outside its original directory."
        If Left$(held, Len(stage) + 1) <> stage & "." Then Err.Raise MB_ERR + 916, , "Held stage suffix is invalid."
        actual = held
    End If
    MBY_NoReparse actual, False: MBY_NoReparse final, False
    If gFSO.FileExists(final) Then
        If gFSO.FileExists(actual) Then Err.Raise MB_ERR + 917, , "Both stage and destination exist; both preserved."
        If phase = "PREPARED" Then
            digest = MBH_HashFile(final): bytes = CDbl(gFSO.GetFile(final).Size)
            If digest <> CStr(row(5)) Or bytes <> MB_ByteCount(CStr(row(6))) Then _
                Err.Raise MB_ERR + 918, , "Published bytes differ from durable preparation proof."
            MBS_StageAppend stage, final, "RECOVERED", digest, bytes
        Else
            If LCase$(Right$(final, 4)) <> ".msg" Then Err.Raise MB_ERR + 917, , "Unrecorded attachment publication needs review."
            If Not MBS_StageReadable(final, digest, bytes) Then Err.Raise MB_ERR + 918, , "Unrecorded published MSG cannot be reopened."
            MBS_StageAppend stage, final, "CANDIDATE", digest, bytes
        End If
        gStageRecovered = gStageRecovered + 1
        Exit Sub
    End If
    If Not gFSO.FileExists(actual) Then
        ' Nothing is adopted and no output is declared complete. Mail inventory
        ' will retry any source that lacks a committed, verified output.
        MBS_StageAppend stage, final, "ABORTED", "", 0
        Exit Sub
    End If
    If phase = "PREPARED" Then
        digest = MBH_HashFile(actual): bytes = CDbl(gFSO.GetFile(actual).Size)
        If digest <> CStr(row(5)) Or bytes <> MB_ByteCount(CStr(row(6))) Then _
            Err.Raise MB_ERR + 918, , "Prepared stage no longer matches its durable expected bytes."
        MBM_FlushPath actual
        If MBM_MoveFileExW(StrPtr(actual), StrPtr(final), 8) = 0 Then _
            Err.Raise MB_ERR + 919, , "Prepared stage publication recovery failed."
        MBM_FlushPath final
        If MBH_HashFile(final) <> digest Or CDbl(gFSO.GetFile(final).Size) <> bytes Then _
            Err.Raise MB_ERR + 920, , "Recovered prepared publication changed."
        MBS_StageAppend stage, final, "RECOVERED", digest, bytes
        gStageRecovered = gStageRecovered + 1
        Exit Sub
    End If
    If LCase$(Right$(final, 4)) = ".msg" Then
        If MBS_StageReadable(actual, digest, bytes) Then
            MBM_FlushPath actual
            If MBM_MoveFileExW(StrPtr(actual), StrPtr(final), 8) = 0 Then _
                Err.Raise MB_ERR + 919, , "Owned stage recovery move failed."
            MBM_FlushPath final
            If MBH_HashFile(final) <> digest Then Err.Raise MB_ERR + 920, , "Recovered file hash changed."
            MBS_StageAppend stage, final, "CANDIDATE", digest, bytes
            gStageRecovered = gStageRecovered + 1
            MB_LogEvent "INFO", "STAGE_RECOVERED_NOT_YET_LINKED", "", final, 0, "MSG readability established; source-content adoption is still required."
            Exit Sub
        End If
    End If
    If phase = "RESERVE" Then
        ' A durable reservation proves ownership. Rename in the same directory
        ' out of the MSG namespace; retain every byte for manual recovery.
        target = actual & ".held"
        If gFSO.FileExists(target) Or gFSO.FolderExists(target) Then Err.Raise MB_ERR + 921, , "Held stage destination is occupied."
        digest = MBH_HashFile(actual): bytes = CDbl(gFSO.GetFile(actual).Size)
        MBS_StageAppend stage, final, "HOLD_PREPARE", digest, bytes, target
        If MBM_MoveFileExW(StrPtr(actual), StrPtr(target), 8) = 0 Then Err.Raise MB_ERR + 922, , "Cannot preserve unfinished stage."
        If MBH_HashFile(target) <> digest Then Err.Raise MB_ERR + 920, , "Preserved stage changed."
        MBS_StageAppend stage, final, "HELD", digest, bytes, target
    End If
    gStageHeld = gStageHeld + 1
    MB_LogEvent "WARNING", "UNFINISHED_STAGE_PRESERVED", "", actual, 0, "Owned incomplete bytes retained; unrelated new mail may continue."
    Exit Sub
Review:
    n = Err.Number: detail = Err.Description
    If gFatal Then Err.Raise n, "MBS_RecoverStage", detail
    gStageReview = gStageReview + 1
    MB_LogEvent "WARNING", "STAGE_RECOVERY_REVIEW", "", stage, n, detail
End Sub

Private Sub MBS_EnableStages(ByVal accountRoot As String, ByVal storeId As String)
    Dim key As Variant, row As Variant, needsMirror As Object, needsPrimary As Object
    gStageAccount = accountRoot
    gStageMirror = MBH_JoinPath(MBR_RecoveryRoot(accountRoot, storeId), "stage-intents.tsv")
    Set gStageRecords = CreateObject("Scripting.Dictionary")
    gStageRecords.CompareMode = vbTextCompare
    Set needsMirror = CreateObject("Scripting.Dictionary"): needsMirror.CompareMode = vbTextCompare
    Set needsPrimary = CreateObject("Scripting.Dictionary"): needsPrimary.CompareMode = vbTextCompare
    MBS_LoadStages MB_StatePath("stages.tsv"), needsMirror, needsPrimary
    MBS_LoadStages gStageMirror, needsMirror, needsPrimary
    If gStageReview > 0 Then
        ' The normal checked load also detects repair candidates. Avoid two
        ' extra full ledger scans on every healthy run.
        gStageLedgerRepaired = MBS_RepairStageLedgers(MB_StatePath("stages.tsv"), gStageMirror)
        If gStageLedgerRepaired > 0 Then
            gStageRecords.RemoveAll: needsMirror.RemoveAll: needsPrimary.RemoveAll
            gStageReview = 0
            MBS_LoadStages MB_StatePath("stages.tsv"), needsMirror, needsPrimary
            MBS_LoadStages gStageMirror, needsMirror, needsPrimary
        End If
    End If
    MBS_CopyStageProofs MB_StatePath("stages.tsv"), needsPrimary
    If gCancel Then Exit Sub
    MBS_CopyStageProofs gStageMirror, needsMirror
    If gCancel Then Exit Sub
    Set needsMirror = Nothing: Set needsPrimary = Nothing
    gStageEnabled = True
    For Each key In gStageRecords.Keys
        If Not MB_ContinueWork() Then Exit Sub
        row = MBS_StageFields(CStr(gStageRecords(key)))
        MBS_RecoverStage row
    Next key
    For Each key In gStageRecords.Keys
        row = MBS_StageFields(CStr(gStageRecords(key)))
        If CStr(row(4)) = "PREPARED" Or CStr(row(4)) = "PUBLISHED" Or CStr(row(4)) = "RECOVERED" Then
            If MB_IsHexKey(CStr(row(5))) Then
                gStageExpected(CStr(row(3)) & "|" & CStr(row(5))) = MB_EncodeFields(Array( _
                    CStr(row(3)), CStr(row(6)), CStr(row(5)), accountRoot, "VERIFIED_PUBLICATION"))
            End If
        End If
    Next key
End Sub

Private Function MBS_Summary() As String
    MBS_Summary = vbCrLf & "Operating folder: " & gStateRoot & vbCrLf & _
        "Old state copied once: " & CStr(gStateMigrated) & "; state missing at start: " & CStr(gStateWasMissing) & vbCrLf & _
        "Recovered staged files: " & CStr(gStageRecovered) & "; held incomplete files: " & CStr(gStageHeld) & _
        "; stage review: " & CStr(gStageReview) & vbCrLf & _
        "Repaired stage rows: " & CStr(gStageLedgerRepaired) & "; restored proof copies: " & CStr(gStageProofCopies) & vbCrLf & _
        "Mail coverage, archive audit, legacy baseline and reporting are separate results."
End Function

Private Function MBS_RelinkOne(ByVal original As String, ByVal sourceRoot As String, ByVal bytes As String, _
                               ByVal digest As String, ByVal accountRoot As String) As Boolean
    Dim source As String, target As String, bound As String, planId As String, fields As Variant
    Dim n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Review
    source = MBM_ResolvePath(original, digest)
    bound = MBM_ResolvedRoot(original, digest, sourceRoot)
    source = MBY_CanonicalFilePath(source, bound)
    If gFSO.FileExists(source) Then Exit Function
    target = MBR_FindExpectedFile(digest, MB_ByteCount(bytes), accountRoot)
    If Len(target) = 0 Then Exit Function
    If StrComp(source, target, vbTextCompare) = 0 Then Exit Function
    guard = MBH_GuardFile(target)
    MBM_VerifyFile target, accountRoot, bytes, digest
    planId = MBH_HashText(MB_EncodeFields(Array("MBRELINK1", source, target, digest, gRunId)))
    fields = Array("MBMOVE3", MBM_OperationId(planId, source, target, digest), planId, "PREPARE", _
                   accountRoot, bound, source, target, bytes, digest, MB_Stamp(Now))
    ' This records an observed byte-identical relocation; no file is moved.
    ' PREPARE followed by normal recovery retains restart semantics.
    MBM_AppendMove fields, "PREPARE"
    MBM_Recover fields
    MBH_CloseGuard guard
    MB_LogEvent "INFO", "MISSING_REFERENCE_RELINKED", "", source, 0, "Unique verified same-account bytes found at " & target
    MBS_RelinkOne = True
    Exit Function
Review:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    If gFatal Then Err.Raise n, "MBS_RelinkOne", detail
    MB_LogEvent "WARNING", "MISSING_REFERENCE_RELINK_REVIEW", "", original, n, detail
End Function

Private Function MBS_RelinkMissingReferences(ByVal accountRoot As String, ByVal storeId As String) As Long
    Dim key As Variant, fields As Variant, meta As Variant, token As String, original As String
    token = MBX_StoreToken(storeId)
    For Each key In gIndex.Keys
        If Not MB_ContinueWork() Then Exit Function
        If Left$(CStr(key), 2) = "M|" Or Left$(CStr(key), 2) = "L|" Then
            meta = MBX_Meta(CStr(key))
            If CStr(meta(0)) = token Then
                fields = MB_IndexFields(CStr(key))
                original = MBX_IndexedOriginal(CStr(key), fields)
                If MBS_RelinkOne(original, MBX_IndexedRoot(CStr(key)), CStr(fields(1)), CStr(fields(2)), accountRoot) Then _
                    MBS_RelinkMissingReferences = MBS_RelinkMissingReferences + 1
            End If
        End If
    Next key
    If Not gRecoveryFiles Is Nothing Then
        For Each key In gRecoveryFiles.Keys
            If Not MB_ContinueWork() Then Exit Function
            fields = MBM_Unpack(CStr(gRecoveryFiles(key)), 5)
            If StrComp(CStr(fields(3)), accountRoot, vbTextCompare) = 0 Then
                If MBS_RelinkOne(CStr(fields(0)), CStr(fields(3)), CStr(fields(1)), CStr(fields(2)), accountRoot) Then _
                    MBS_RelinkMissingReferences = MBS_RelinkMissingReferences + 1
            End If
        Next key
    End If
End Function

Private Sub MBS_ClearLoadedIndex(ByVal storeId As String)
    ' Reload a repaired journal into empty caches; never mix generations.
    gIndex.RemoveAll: gExternalRoots.RemoveAll: gRecordMeta.RemoveAll
    gStoreTokens.RemoveAll: gManagedFilePaths.RemoveAll
    gLegacyBindings.RemoveAll: gLegacySources.RemoveAll: gLegacyOwners.RemoveAll
    gJournalIssues = 0
    MBI_Reset
    MBW_Reset
    MBW_Start storeId
    MBY_Reset
End Sub


' ================== Independent metadata recovery / MSG baseline =============
' The independent recovery directory holds metadata only, never extra MSG copies.
' Snapshots are immutable manifests over content-addressed 1 MiB chunks, so an
' append to a large journal does not duplicate every unchanged prefix block.
' A reconstructed MSG-only baseline proves readable bytes NOW. It cannot prove
' equality to a lost historical hash, recover deleted mail, or restore lost IDs.

Private Sub MBR_Reset()
    gRecoveryForceFull = False: gRecoveryRebased = False
    gRecoveryReviewCount = 0: gRecoveryRestoredCount = 0: gRecoveryNewBaselines = 0
    mbrBaselineWrites = 0: gRecoverySnapshotsSaved = 0: mbrActiveRecoveryRoot = ""
    gRecoverySalvagedRows = 0
    gRecoveryMergedRows = 0: gRecoveryOverlayRejected = 0: gRecoveryOverlayFiles = 0
    gRecoveryNote = ""
    Set gRecoveryFiles = CreateObject("Scripting.Dictionary")
    gRecoveryFiles.CompareMode = vbTextCompare
    Set gRecoverySentHints = CreateObject("Scripting.Dictionary")
    gRecoverySentHints.CompareMode = vbTextCompare
    Set mbrDigestFiles = CreateObject("Scripting.Dictionary")
    mbrDigestFiles.CompareMode = vbBinaryCompare
End Sub

Private Sub MBR_Log(ByVal level As String, ByVal stage As String, ByVal key As String, _
                     ByVal path As String, ByVal errorNumber As Long, ByVal detail As String)
    Dim stream As Object, logPath As String, n As Long, description As String
    If Not gEvents Is Nothing Then
        MB_LogEvent level, stage, key, path, errorNumber, detail
        Exit Sub
    End If
    On Error GoTo Failed
    If Len(mbrActiveRecoveryRoot) = 0 Then Err.Raise MB_ERR + 863, , "Recovery event has no bound destination."
    logPath = MBH_JoinPath(mbrActiveRecoveryRoot, "recovery-events.tsv")
    MBY_NoReparse logPath, False
    Set stream = gFSO.OpenTextFile(logPath, 8, True, -1)
    stream.WriteLine ""
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBREVENT1", MBC_UtcStamp(MBC_UtcNow()), _
        level, stage, key, path, CStr(errorNumber), detail))
    stream.Close: Set stream = Nothing
    MBM_FlushPath logPath
    Exit Sub
Failed:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_Log", description
End Sub

Private Function MBR_HasReview() As Boolean
    MBR_HasReview = (gRecoveryReviewCount > 0)
End Function

Private Function MBR_Summary() As String
    Dim count As Long
    If Not gRecoveryFiles Is Nothing Then count = gRecoveryFiles.Count
    MBR_Summary = vbCrLf & "Recovery folder: " & mbrActiveRecoveryRoot & vbCrLf & _
        "Recovered state files: " & CStr(gRecoveryRestoredCount) & _
        "; new metadata snapshots: " & CStr(gRecoverySnapshotsSaved) & vbCrLf & _
        "Salvaged valid prior index rows: " & CStr(gRecoverySalvagedRows) & vbCrLf & _
        "Preserved pre-restore state files: " & CStr(gRecoveryOverlayFiles) & _
        "; later journal rows retained: " & CStr(gRecoveryMergedRows) & _
        "; unproven raw rows preserved: " & CStr(gRecoveryOverlayRejected) & vbCrLf & _
        "MSG-only baseline records: " & CStr(count) & "; new this run: " & CStr(gRecoveryNewBaselines) & _
        "; baseline review: " & CStr(gRecoveryReviewCount) & vbCrLf & _
        "Full reconciliation required: " & CStr(gRecoveryForceFull) & _
        "; rebuilt without prior hashes this run: " & CStr(gRecoveryRebased) & vbCrLf & gRecoveryNote
    If count > 0 Then MBR_Summary = MBR_Summary & vbCrLf & _
        "MSG_ONLY_REBASELINED proves present readable bytes; lost historical equality remains unproven."
End Function

Private Function MBR_ReadOwner(ByVal folderPath As String) As Variant
    Dim path As String, stream As Object, fields As Variant, n As Long, detail As String
    On Error GoTo Failed
    folderPath = MBH_NormalizePath(folderPath)
    path = MBH_JoinPath(folderPath, MBR_OWNER_FILE)
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 800, , "Recovery owner is truncated."
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    If stream.AtEndOfStream Then Err.Raise MB_ERR + 801, , "Recovery owner is empty."
    fields = MB_ReadCheckedLine(stream.ReadLine, 5)
    If Not stream.AtEndOfStream Then Err.Raise MB_ERR + 802, , "Recovery owner has extra data."
    stream.Close: Set stream = Nothing
    If CStr(fields(0)) <> "MBRECOVERY1" Or CStr(fields(4)) <> "1" Then _
        Err.Raise MB_ERR + 803, , "Unknown recovery owner schema."
    If Len(CStr(fields(3))) = 0 Then Err.Raise MB_ERR + 804, , "Recovery owner has no StoreID."
    If StrComp(MBH_JoinPath(CStr(fields(2)), MBR_FOLDER), folderPath, vbTextCompare) <> 0 And _
       StrComp(MBH_JoinPath(CStr(fields(2)), MBR_FOLDER_LEGACY), folderPath, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 805, , "Recovery directory moved outside its bound account."
    If Len(gRoot) > 0 Then
        If StrComp(MBH_NormalizePath(CStr(fields(1))), MBH_NormalizePath(gRoot), vbTextCompare) <> 0 Then _
            Err.Raise MB_ERR + 806, , "Recovery directory belongs to a different managed root."
    End If
    MBR_ReadOwner = fields
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_ReadOwner", detail
End Function

Private Function MBR_ExpectedRecoveryPath(ByVal accountRoot As String, ByVal managedRoot As String) As String
    If MBS_IsModernLayout(managedRoot) Then
        MBR_ExpectedRecoveryPath = MBH_JoinPath(accountRoot, MBR_FOLDER)
    Else
        MBR_ExpectedRecoveryPath = MBH_JoinPath(accountRoot, MBR_FOLDER_LEGACY)
    End If
End Function

Private Function MBR_ExistingRecoveryPath(ByVal accountRoot As String, ByVal fso As Object) As String
    Dim modern As String, legacy As String, hasModern As Boolean, hasLegacy As Boolean
    modern = MBH_JoinPath(accountRoot, MBR_FOLDER)
    legacy = MBH_JoinPath(accountRoot, MBR_FOLDER_LEGACY)
    MBP_NoReparse modern, fso: MBP_NoReparse legacy, fso
    If fso.FileExists(modern) Or fso.FileExists(legacy) Then Err.Raise MB_ERR + 1211, , _
        "A file occupies a recovery-directory name. Existing files were retained."
    hasModern = fso.FolderExists(modern): hasLegacy = fso.FolderExists(legacy)
    If hasModern And hasLegacy Then Err.Raise MB_ERR + 1212, , _
        "Both _Recovery and _MailBackup_Recovery exist. Review the independent histories before choosing a layout."
    If hasModern Then MBR_ExistingRecoveryPath = modern
    If hasLegacy Then MBR_ExistingRecoveryPath = legacy
End Function

Private Function MBR_IsRecoveryDirectory(ByVal folderPath As String) As Boolean
    Dim fields As Variant
    If Not gFSO.FileExists(MBH_JoinPath(folderPath, MBR_OWNER_FILE)) Then Exit Function
    fields = MBR_ReadOwner(folderPath)
    MBR_IsRecoveryDirectory = True
End Function

Private Function MBR_RecoveryRoot(ByVal accountRoot As String, ByVal storeId As String) As String
    Dim root As String, marker As String, fields As Variant, stream As Object, folder As Object
    Dim n As Long, detail As String, manifest As String, files As Object, chunks As Object, existing As String
    On Error GoTo Failed
    accountRoot = MBH_NormalizePath(accountRoot)
    If Len(storeId) = 0 Or Len(gRoot) = 0 Then Err.Raise MB_ERR + 807, , "Recovery needs root and StoreID bindings."
    MBY_NoReparse accountRoot, True
    root = MBR_ExpectedRecoveryPath(accountRoot, gRoot)
    existing = MBR_ExistingRecoveryPath(accountRoot, gFSO)
    If Len(existing) > 0 Then
        If StrComp(existing, root, vbTextCompare) <> 0 Then Err.Raise MB_ERR + 1213, , _
            "Recovery metadata uses a different layout from the managed root. Existing paths were retained for review: " & existing
    End If
    marker = MBH_JoinPath(root, MBR_OWNER_FILE)
    If gFSO.FolderExists(root) Then
        MBY_NoReparse root, True
        If Not gFSO.FileExists(marker) Then
            Set folder = gFSO.GetFolder(root)
            If folder.Files.Count <> 0 Or folder.SubFolders.Count <> 0 Then
                ' A deleted marker can be reconstructed from an independently
                ' checked complete generation with the exact same three bindings.
                manifest = MBR_FindSnapshot(root, accountRoot, storeId, files, chunks, True, True)
                If Len(manifest) = 0 Then Err.Raise MB_ERR + 808, , _
                    "Recovery owner is missing and no complete exactly bound snapshot proves ownership: " & root
            End If
        End If
    Else
        MBH_EnsureDirectory root
    End If
    If Not gFSO.FileExists(marker) Then
        Set stream = gFSO.CreateTextFile(marker, False, True)
        MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRECOVERY1", gRoot, accountRoot, storeId, "1"))
        stream.Close: Set stream = Nothing
        MBM_FlushPath marker
    End If
    fields = MBR_ReadOwner(root)
    If StrComp(CStr(fields(2)), accountRoot, vbTextCompare) <> 0 Or _
       StrComp(CStr(fields(3)), storeId, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 809, , "Recovery owner does not match the selected account."
    mbrActiveRecoveryRoot = root
    MBR_RecoveryRoot = root
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_RecoveryRoot", detail
End Function

Private Function MBR_IsStateLeaf(ByVal leaf As String) As Boolean
    If Len(leaf) = 0 Or InStr(leaf, "\") > 0 Or InStr(leaf, "/") > 0 Or InStr(leaf, ":") > 0 Then Exit Function
    If Left$(leaf, 1) = "." Or Left$(leaf, 4) = "~mb_" Then Exit Function
    MBR_IsStateLeaf = (LCase$(Right$(leaf, 4)) = ".tsv")
End Function

Private Function MBR_CheckedStateValid(ByVal path As String, ByRef reason As String) As Boolean
    Dim stream As Object, line As String, raw As Variant, fields As Variant, n As Long
    Dim auditKind As String, sawCursor As Boolean, leaf As String
    On Error GoTo Invalid
    reason = ""
    leaf = LCase$(CStr(gFSO.GetFileName(path)))
    If Left$(leaf, 13) = "source_audit_" Then auditKind = "SOURCE"
    If Left$(leaf, 14) = "archive_audit_" Then auditKind = "ARCHIVE"
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 810, , "Odd UTF-16 byte count."
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            If sawCursor Then Err.Raise MB_ERR + 865, , "Data follows an audit snapshot terminator."
            raw = Split(line, vbTab)
            n = UBound(raw) - 1
            If n < 1 Then Err.Raise MB_ERR + 811, , "State record has no checked framing."
            fields = MB_ReadCheckedLine(line, n)
            If auditKind = "SOURCE" Then
                If n <> 9 Then Err.Raise MB_ERR + 866, , "Source-audit state has the wrong field count."
                If CStr(fields(0)) <> "MBSOURCEAUDIT1" Then Err.Raise MB_ERR + 867, , "Wrong source-audit schema."
                sawCursor = (CStr(fields(5)) = "CURSOR")
            ElseIf auditKind = "ARCHIVE" Then
                If n <> 17 Then Err.Raise MB_ERR + 868, , "Archive-audit state has the wrong field count."
                If CStr(fields(0)) <> "MBARCHIVEAUDIT1" Then Err.Raise MB_ERR + 869, , "Wrong archive-audit schema."
                sawCursor = (CStr(fields(4)) = "CURSOR" And CStr(fields(13)) = "CURSOR")
            End If
        End If
    Loop
    stream.Close: Set stream = Nothing
    If Len(auditKind) > 0 And Not sawCursor Then Err.Raise MB_ERR + 870, , "Audit snapshot terminator is missing."
    MBR_CheckedStateValid = True
    Exit Function
Invalid:
    reason = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
End Function

Private Function MBR_BlockPath(ByVal recoveryRoot As String, ByVal digest As String) As String
    If Not MB_IsHexKey(digest) Then Err.Raise MB_ERR + 812, , "Invalid recovery chunk digest."
    MBR_BlockPath = MBH_JoinPath(MBH_JoinPath(recoveryRoot, "chunks"), LCase$(digest) & ".bin")
End Function

Private Function MBR_HashBuffer(ByRef buffer() As Byte, ByVal count As Long) As String
    Dim context As MBH_HashContext, n As Long, detail As String
    On Error GoTo Failed
    MBH_HashOpen context
    If count > 0 Then MBH_CheckStatus MBH_BCryptHashData(context.Handle, VarPtr(buffer(0)), count, 0), "RecoveryChunkHash"
    MBR_HashBuffer = MBH_HashFinish(context)
    MBH_HashClose context
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_HashClose context
    On Error GoTo 0
    Err.Raise n, "MBR_HashBuffer", detail
End Function

Private Sub MBR_StoreBlock(ByVal recoveryRoot As String, ByRef buffer() As Byte, _
                           ByVal count As Long, ByVal digest As String)
    Dim path As String, stage As String, written As Long, n As Long, detail As String, rc As Long
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    handle = -1
    On Error GoTo Failed
    path = MBR_BlockPath(recoveryRoot, digest)
    MBY_NoReparse path, False
    MBH_EnsureDirectory gFSO.GetParentFolderName(path)
    If gFSO.FileExists(path) Then
        MBY_NoReparse path, True
        If CDbl(gFSO.GetFile(path).Size) = count Then
            If MBH_HashFile(path) = digest Then Exit Sub
        End If
        ' Keep corrupted owned metadata as evidence; create a fresh verified copy.
        MBR_PreserveFile path, MBH_JoinPath(recoveryRoot, "preserved")
    End If
    stage = MBH_NewStagePath(path)
    handle = MBH_CreateFileW(StrPtr(stage), MBH_GENERIC_WRITE, 0, 0, MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If handle = -1 Then Err.Raise MB_ERR + 813, , "Cannot open recovery chunk stage."
    If MBR_WriteFile(handle, buffer(0), count, written, 0) = 0 Or written <> count Then _
        Err.Raise MB_ERR + 814, , "Incomplete recovery chunk write."
    If MBM_FlushFileBuffers(handle) = 0 Then Err.Raise MB_ERR + 815, , "Cannot flush recovery chunk."
    rc = MBH_CloseHandle(handle): handle = -1
    If rc = 0 Then Err.Raise MB_ERR + 816, , "Cannot close recovery chunk."
    If MBH_HashFile(stage) <> digest Then Err.Raise MB_ERR + 817, , "Recovery chunk stage digest mismatch."
    MBH_CommitStage stage, path
    stage = ""
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If handle <> -1 Then rc = MBH_CloseHandle(handle)
    If Len(stage) > 0 Then MBH_DeleteStage stage
    On Error GoTo 0
    Err.Raise n, "MBR_StoreBlock", detail
End Sub

Private Function MBR_CaptureState(ByVal path As String, ByVal recoveryRoot As String, _
                                   ByRef bytes As Double, ByRef digest As String) As Collection
    Dim blocks As New Collection, buffer() As Byte, context As MBH_HashContext
    Dim count As Long, blockHash As String, rc As Long, n As Long, detail As String
#If VBA7 Then
    Dim handle As LongPtr
#Else
    Dim handle As Long
#End If
    handle = -1
    On Error GoTo Failed
    bytes = 0: digest = ""
    ReDim buffer(0 To MBR_CHUNK_BYTES - 1)
    handle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_READ, MBH_FILE_SHARE_READ, 0, _
                            MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL Or MBH_FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If handle = -1 Then Err.Raise MB_ERR + 818, , "Cannot lock primary state for a snapshot: " & path
    MBH_HashOpen context
    Do
        If MBH_ReadFile(handle, buffer(0), MBR_CHUNK_BYTES, count, 0) = 0 Then _
            Err.Raise MB_ERR + 819, , "Cannot read primary state."
        If count = 0 Then Exit Do
        MBH_CheckStatus MBH_BCryptHashData(context.Handle, VarPtr(buffer(0)), count, 0), "RecoveryStateHash"
        blockHash = MBR_HashBuffer(buffer, count)
        MBR_StoreBlock recoveryRoot, buffer, count, blockHash
        blocks.Add Array(CStr(blocks.Count + 1), CStr(count), blockHash)
        bytes = bytes + count
    Loop
    digest = MBH_HashFinish(context)
    MBH_HashClose context
    rc = MBH_CloseHandle(handle): handle = -1
    If rc = 0 Then Err.Raise MB_ERR + 820, , "Cannot close primary snapshot reader."
    Set MBR_CaptureState = blocks
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If handle <> -1 Then rc = MBH_CloseHandle(handle)
    MBH_HashClose context
    On Error GoTo 0
    Err.Raise n, "MBR_CaptureState", detail
End Function

Private Function MBR_ReadSnapshot(ByVal path As String, ByVal accountRoot As String, ByVal storeId As String, _
                                   ByRef files As Object, ByRef chunks As Object, ByRef reason As String) As Boolean
    Dim stream As Object, line As String, raw As Variant, f As Variant, meta As Variant
    Dim header As Boolean, footer As Boolean, leaf As String, blocks As Collection
    Dim totalChunks As Long, key As Variant, n As Long
    On Error GoTo Invalid
    Set files = CreateObject("Scripting.Dictionary"): files.CompareMode = vbTextCompare
    Set chunks = CreateObject("Scripting.Dictionary"): chunks.CompareMode = vbTextCompare
    reason = ""
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 821, , "Snapshot has a truncated UTF-16 byte."
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            If footer Then Err.Raise MB_ERR + 822, , "Snapshot has records after its completion footer."
            raw = Split(line, vbTab)
            If UBound(raw) < 3 Then Err.Raise MB_ERR + 823, , "Snapshot record is incomplete."
            Select Case CStr(raw(1))
                Case "HEADER"
                    If header Then Err.Raise MB_ERR + 824, , "Duplicate snapshot header."
                    f = MB_ReadCheckedLine(line, 7)
                    If CStr(f(0)) <> "MBRSNAP1" Or StrComp(CStr(f(2)), gRoot, vbTextCompare) <> 0 Or _
                       StrComp(CStr(f(3)), accountRoot, vbTextCompare) <> 0 Or _
                       StrComp(CStr(f(4)), storeId, vbTextCompare) <> 0 Then _
                        Err.Raise MB_ERR + 825, , "Snapshot belongs to another root/account/store."
                    If Not MBC_ValidDate(MBC_ParseUtcStamp(CStr(f(5)))) Then _
                        Err.Raise MB_ERR + 826, , "Snapshot UTC timestamp is invalid."
                    header = True
                Case "FILE"
                    If Not header Then Err.Raise MB_ERR + 827, , "Snapshot has no header."
                    f = MB_ReadCheckedLine(line, 6): leaf = CStr(f(2))
                    If CStr(f(0)) <> "MBRSNAP1" Or Not MBR_IsStateLeaf(leaf) Then _
                        Err.Raise MB_ERR + 828, , "Unsafe snapshot state filename."
                    If files.Exists(leaf) Or Not MB_IsHexKey(CStr(f(4))) Then _
                        Err.Raise MB_ERR + 829, , "Duplicate snapshot file or invalid digest."
                    If MB_ByteCount(CStr(f(3))) < 0 Or MB_ByteCount(CStr(f(5))) < 0 Then _
                        Err.Raise MB_ERR + 830, , "Invalid snapshot size/count."
                    files.Add leaf, MB_EncodeFields(Array(f(3), f(4), f(5)))
                    Set blocks = New Collection: chunks.Add leaf, blocks
                Case "CHUNK"
                    f = MB_ReadCheckedLine(line, 6): leaf = CStr(f(2))
                    If CStr(f(0)) <> "MBRSNAP1" Or Not files.Exists(leaf) Then _
                        Err.Raise MB_ERR + 831, , "Snapshot chunk has no preceding file."
                    Set blocks = chunks(leaf)
                    If MB_ByteCount(CStr(f(3))) <> blocks.Count + 1 Then _
                        Err.Raise MB_ERR + 832, , "Snapshot chunk order is invalid."
                    If MB_ByteCount(CStr(f(4))) <= 0 Or MB_ByteCount(CStr(f(4))) > MBR_CHUNK_BYTES Or _
                       Not MB_IsHexKey(CStr(f(5))) Then Err.Raise MB_ERR + 833, , "Invalid recovery chunk."
                    blocks.Add Array(f(3), f(4), f(5))
                    totalChunks = totalChunks + 1
                Case "END"
                    f = MB_ReadCheckedLine(line, 4)
                    If Not header Or CStr(f(0)) <> "MBRSNAP1" Or MB_ByteCount(CStr(f(2))) <> files.Count Or _
                       MB_ByteCount(CStr(f(3))) <> totalChunks Then Err.Raise MB_ERR + 834, , "Incomplete snapshot footer."
                    footer = True
                Case Else
                    Err.Raise MB_ERR + 835, , "Unknown snapshot record."
            End Select
        End If
    Loop
    stream.Close: Set stream = Nothing
    If Not header Or Not footer Or Not files.Exists("journal.tsv") Then _
        Err.Raise MB_ERR + 836, , "Snapshot has no complete journal generation."
    For Each key In files.Keys
        meta = MBM_Unpack(CStr(files(key)), 3): Set blocks = chunks(key)
        If blocks.Count <> MB_ByteCount(CStr(meta(2))) Then _
            Err.Raise MB_ERR + 837, , "Snapshot file chunk count mismatch."
    Next key
    MBR_ReadSnapshot = True
    Exit Function
Invalid:
    reason = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Set files = Nothing: Set chunks = Nothing
End Function

Private Function MBR_BlocksValid(ByVal recoveryRoot As String, ByVal files As Object, _
                                 ByVal chunks As Object, ByRef reason As String) As Boolean
    Dim key As Variant, meta As Variant, blocks As Collection, block As Variant
    Dim path As String, total As Double
    On Error GoTo Invalid
    For Each key In files.Keys
        meta = MBM_Unpack(CStr(files(key)), 3): Set blocks = chunks(key): total = 0
        For Each block In blocks
            path = MBR_BlockPath(recoveryRoot, CStr(block(2)))
            MBY_NoReparse path, True
            If CDbl(gFSO.GetFile(path).Size) <> MB_ByteCount(CStr(block(1))) Then _
                Err.Raise MB_ERR + 838, , "Snapshot chunk size mismatch."
            If MBH_HashFile(path) <> CStr(block(2)) Then Err.Raise MB_ERR + 839, , "Snapshot chunk SHA-256 mismatch."
            total = total + MB_ByteCount(CStr(block(1)))
        Next block
        If total <> MB_ByteCount(CStr(meta(0))) Then Err.Raise MB_ERR + 840, , "Snapshot total bytes mismatch."
    Next key
    MBR_BlocksValid = True
    Exit Function
Invalid:
    reason = Err.Description
End Function

Private Function MBR_FindSnapshot(ByVal recoveryRoot As String, ByVal accountRoot As String, _
                                  ByVal storeId As String, ByRef files As Object, ByRef chunks As Object, _
                                  Optional ByVal verifyBlocks As Boolean = False, _
                                  Optional ByVal quiet As Boolean = False) As String
    Dim dir As String, item As Object, candidates As Object, keys As Variant, i As Long, reason As String
    Set candidates = CreateObject("Scripting.Dictionary")
    dir = MBH_JoinPath(recoveryRoot, "snapshots")
    If Not gFSO.FolderExists(dir) Then Exit Function
    MBY_NoReparse dir, True
    For Each item In gFSO.GetFolder(dir).Files
        If LCase$(Left$(CStr(item.Name), 9)) = "snapshot_" And LCase$(Right$(CStr(item.Name), 4)) = ".tsv" Then
            candidates(CStr(item.Path)) = True
        End If
    Next item
    If candidates.Count = 0 Then Exit Function
    keys = candidates.Keys: MBW_SortKeys keys, LBound(keys), UBound(keys)
    For i = UBound(keys) To LBound(keys) Step -1
        If MBR_ReadSnapshot(CStr(keys(i)), accountRoot, storeId, files, chunks, reason) Then
            If Not verifyBlocks Then
                MBR_FindSnapshot = CStr(keys(i)): Exit Function
            ElseIf MBR_BlocksValid(recoveryRoot, files, chunks, reason) Then
                MBR_FindSnapshot = CStr(keys(i)): Exit Function
            End If
        End If
        If Not quiet Then MBR_Log "WARNING", "RECOVERY_SNAPSHOT_REJECTED", "", CStr(keys(i)), 0, reason
    Next i
    Set files = Nothing: Set chunks = Nothing
End Function

Private Function MBR_StateHealthy(ByVal snapshotFiles As Object, ByRef reason As String) As Boolean
    Dim dir As String, item As Object, key As Variant, meta As Variant, path As String
    reason = ""
    dir = gFSO.GetParentFolderName(MB_StatePath("journal.tsv"))
    If Not gFSO.FileExists(MB_StatePath("journal.tsv")) Then reason = "Primary journal is missing.": Exit Function
    If Not gFSO.FolderExists(dir) Then reason = "Primary state directory is missing.": Exit Function
    For Each item In gFSO.GetFolder(dir).Files
        If MBR_IsStateLeaf(CStr(item.Name)) Then
            If Not MBR_CheckedStateValid(CStr(item.Path), reason) Then
                reason = CStr(item.Name) & ": " & reason: Exit Function
            End If
        End If
    Next item
    If Not snapshotFiles Is Nothing Then
        For Each key In snapshotFiles.Keys
            path = MB_StatePath(CStr(key))
            If Not gFSO.FileExists(path) Then reason = "State component is missing: " & CStr(key): Exit Function
            ' Journal/moves/baseline are append-only. Compact audits may shrink.
            If LCase$(CStr(key)) = "journal.tsv" Or LCase$(CStr(key)) = "moves.tsv" Or _
               LCase$(Left$(CStr(key), 9)) = "baseline_" Then
                meta = MBM_Unpack(CStr(snapshotFiles(key)), 3)
                If CDbl(gFSO.GetFile(path).Size) < MB_ByteCount(CStr(meta(0))) Then
                    reason = "Append-only state is shorter than its recovery snapshot: " & CStr(key): Exit Function
                End If
            End If
        Next key
    End If
    MBR_StateHealthy = True
End Function

Private Sub MBR_PreserveFile(ByVal path As String, ByVal preservedRoot As String)
    Dim target As String, nonce As String
    If Not gFSO.FileExists(path) Then Exit Sub
    MBY_NoReparse path, True
    MBY_NoReparse preservedRoot, False
    MBH_EnsureDirectory preservedRoot
    nonce = Left$(MBH_HashText(gFSO.GetTempName & CStr(Timer)), 12)
    target = MBH_JoinPath(preservedRoot, nonce & "_" & CStr(gFSO.GetFileName(path)))
    If MBH_MoveFileW(StrPtr(path), StrPtr(target)) = 0 Then _
        Err.Raise MB_ERR + 841, , "Cannot preserve previous state without overwrite: " & path
    MBM_FlushPath target
    MBR_Log "WARNING", "RECOVERY_PREVIOUS_STATE_PRESERVED", "", target, 0, "Previous bytes retained before recovery."
End Sub

Private Sub MBR_WriteIntent(ByVal recoveryRoot As String, ByVal accountRoot As String, ByVal storeId As String, _
                            ByVal manifest As String, ByVal phase As String)
    Dim path As String, stream As Object, n As Long, detail As String
    On Error GoTo Failed
    path = MBH_JoinPath(recoveryRoot, "restore-intents.tsv")
    MBY_NoReparse path, False
    Set stream = gFSO.OpenTextFile(path, 8, True, -1)
    stream.WriteLine ""
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRRESTORE1", gRoot, accountRoot, storeId, manifest, phase, MBC_UtcStamp(MBC_UtcNow())))
    stream.Close: Set stream = Nothing
    MBM_FlushPath path
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_WriteIntent", detail
End Sub

Private Function MBR_PendingIntent(ByVal recoveryRoot As String, ByVal accountRoot As String, ByVal storeId As String) As String
    Dim path As String, stream As Object, line As String, f As Variant, latest As Variant
    Dim n As Long, detail As String, saved As Boolean
    On Error GoTo Failed
    path = MBH_JoinPath(recoveryRoot, "restore-intents.tsv")
    If Not gFSO.FileExists(path) Then Exit Function
    MBY_NoReparse path, True
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            On Error Resume Next
            Err.Clear: f = MB_ReadCheckedLine(line, 7): n = Err.Number
            On Error GoTo Failed
            If n = 0 Then
                If CStr(f(0)) <> "MBRRESTORE1" Or StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Or _
                   StrComp(CStr(f(2)), accountRoot, vbTextCompare) <> 0 Or _
                   StrComp(CStr(f(3)), storeId, vbTextCompare) <> 0 Then _
                    Err.Raise MB_ERR + 842, , "Restore intent belongs to another binding."
                If CStr(f(5)) <> "START" And CStr(f(5)) <> "DONE" Then _
                    Err.Raise MB_ERR + 843, , "Unknown restore intent phase."
                latest = f: saved = True
            Else
                MBR_Log "WARNING", "RECOVERY_TORN_INTENT", "", path, n, "A complete prior intent remains authoritative."
            End If
        End If
    Loop
    stream.Close: Set stream = Nothing
    If saved Then
        If CStr(latest(5)) = "START" Then MBR_PendingIntent = CStr(latest(4))
    End If
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_PendingIntent", detail
End Function

Private Sub MBR_ReconstructFile(ByVal recoveryRoot As String, ByVal leaf As String, _
                                ByVal meta As Variant, ByVal blocks As Collection)
    Dim finalPath As String, stage As String, block As Variant, path As String, buffer() As Byte
    Dim count As Long, written As Long, rc As Long, total As Double, n As Long, detail As String
#If VBA7 Then
    Dim outputHandle As LongPtr, inputHandle As LongPtr
#Else
    Dim outputHandle As Long, inputHandle As Long
#End If
    outputHandle = -1: inputHandle = -1
    On Error GoTo Failed
    finalPath = MB_StatePath(leaf)
    If gFSO.FileExists(finalPath) Then
        If CDbl(gFSO.GetFile(finalPath).Size) = MB_ByteCount(CStr(meta(0))) Then
            If MBH_HashFile(finalPath) = CStr(meta(1)) Then Exit Sub
        End If
        MBR_PreserveFile finalPath, MB_StatePath("preserved")
    End If
    MBH_EnsureDirectory gFSO.GetParentFolderName(finalPath)
    stage = MBH_NewStagePath(finalPath)
    outputHandle = MBH_CreateFileW(StrPtr(stage), MBH_GENERIC_WRITE, 0, 0, MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL, 0)
    If outputHandle = -1 Then Err.Raise MB_ERR + 844, , "Cannot create restored state stage."
    ReDim buffer(0 To MBR_CHUNK_BYTES - 1)
    For Each block In blocks
        path = MBR_BlockPath(recoveryRoot, CStr(block(2)))
        inputHandle = MBH_CreateFileW(StrPtr(path), MBH_GENERIC_READ, MBH_FILE_SHARE_READ, 0, MBH_OPEN_EXISTING, MBH_FILE_ATTRIBUTE_NORMAL, 0)
        If inputHandle = -1 Then Err.Raise MB_ERR + 845, , "Cannot open snapshot chunk."
        If MBH_ReadFile(inputHandle, buffer(0), MBR_CHUNK_BYTES, count, 0) = 0 Then _
            Err.Raise MB_ERR + 846, , "Cannot read snapshot chunk."
        If count <> MB_ByteCount(CStr(block(1))) Or MBR_HashBuffer(buffer, count) <> CStr(block(2)) Then _
            Err.Raise MB_ERR + 847, , "Snapshot chunk changed while restoring."
        If MBR_WriteFile(outputHandle, buffer(0), count, written, 0) = 0 Or written <> count Then _
            Err.Raise MB_ERR + 848, , "Incomplete restored state write."
        rc = MBH_CloseHandle(inputHandle): inputHandle = -1
        If rc = 0 Then Err.Raise MB_ERR + 849, , "Cannot close snapshot chunk."
        total = total + count
    Next block
    If MBM_FlushFileBuffers(outputHandle) = 0 Then Err.Raise MB_ERR + 850, , "Cannot flush restored state."
    rc = MBH_CloseHandle(outputHandle): outputHandle = -1
    If rc = 0 Then Err.Raise MB_ERR + 851, , "Cannot close restored state."
    If total <> MB_ByteCount(CStr(meta(0))) Or MBH_HashFile(stage) <> CStr(meta(1)) Then _
        Err.Raise MB_ERR + 852, , "Restored state failed final size/digest verification."
    MBH_CommitStage stage, finalPath: stage = ""
    gRecoveryRestoredCount = gRecoveryRestoredCount + 1
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If inputHandle <> -1 Then rc = MBH_CloseHandle(inputHandle)
    If outputHandle <> -1 Then rc = MBH_CloseHandle(outputHandle)
    If Len(stage) > 0 Then MBH_DeleteStage stage
    On Error GoTo 0
    Err.Raise n, "MBR_ReconstructFile", detail
End Sub

Private Sub MBR_PreserveExtraState(ByVal snapshotFiles As Object)
    Dim dir As String, item As Object, paths As New Collection, path As Variant
    dir = gFSO.GetParentFolderName(MB_StatePath("journal.tsv"))
    If Not gFSO.FolderExists(dir) Then Exit Sub
    For Each item In gFSO.GetFolder(dir).Files
        If MBR_IsStateLeaf(CStr(item.Name)) Then
            If snapshotFiles Is Nothing Then
                paths.Add CStr(item.Path)
            ElseIf Not snapshotFiles.Exists(CStr(item.Name)) Then
                paths.Add CStr(item.Path)
            End If
        End If
    Next item
    For Each path In paths
        MBR_PreserveFile CStr(path), MB_StatePath("preserved")
    Next path
End Sub

Private Function MBR_ReconcileLeaf(ByVal accountRoot As String, ByVal storeId As String) As String
    MBR_ReconcileLeaf = "reconcile_" & Left$(MBH_HashText(MB_EncodeFields(Array(gRoot, accountRoot, storeId))), 32) & ".tsv"
End Function

Private Function MBR_TokenNeeds(ByVal path As String, ByVal accountRoot As String, ByVal storeId As String) As Boolean
    Dim stream As Object, f As Variant, n As Long, detail As String
    If Not gFSO.FileExists(path) Then Exit Function
    On Error GoTo Damaged
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 875, , "Reconciliation token is truncated."
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    If stream.AtEndOfStream Then Err.Raise MB_ERR + 875, , "Reconciliation token is empty."
    f = MB_ReadCheckedLine(stream.ReadLine, 7)
    If Not stream.AtEndOfStream Then Err.Raise MB_ERR + 875, , "Reconciliation token has extra rows."
    stream.Close: Set stream = Nothing
    If CStr(f(0)) <> "MBRECONCILE1" Or StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Or _
       StrComp(CStr(f(2)), accountRoot, vbTextCompare) <> 0 Or StrComp(CStr(f(3)), storeId, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 876, , "Reconciliation token belongs to another archive/account/store."
    If CStr(f(4)) <> "NEEDS" And CStr(f(4)) <> "DONE" Then Err.Raise MB_ERR + 875, , "Unknown reconciliation token state."
    MBR_TokenNeeds = (CStr(f(4)) = "NEEDS")
    Exit Function
Damaged:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    If n = MB_ERR + 876 Then Err.Raise n, "MBR_TokenNeeds", detail
    ' An unreadable token never means completed reconciliation.
    MBR_TokenNeeds = True
    MBR_Log "WARNING", "RECOVERY_RECONCILE_TOKEN_REVIEW", "", path, n, detail
End Function

Private Sub MBR_WriteToken(ByVal path As String, ByVal recoveryRoot As String, _
                           ByVal accountRoot As String, ByVal storeId As String, ByVal phase As String)
    Dim leaf As String, stage As String, stream As Object, digest As String, n As Long, detail As String
    Dim expectedPrimary As String, expectedMirror As String, created As Boolean
    On Error GoTo Failed
    leaf = MBR_ReconcileLeaf(accountRoot, storeId)
    expectedPrimary = MB_StatePath(leaf): expectedMirror = MBH_JoinPath(recoveryRoot, leaf)
    If StrComp(path, expectedPrimary, vbTextCompare) <> 0 And StrComp(path, expectedMirror, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 877, , "Reconciliation token destination is outside its exact owned binding."
    If phase <> "NEEDS" And phase <> "DONE" Then Err.Raise MB_ERR + 878, , "Unknown reconciliation phase."
    MBY_NoReparse path, False
    MBH_EnsureDirectory gFSO.GetParentFolderName(path)
    stage = MBH_JoinPath(CStr(gFSO.GetParentFolderName(path)), "_mb_reconcile_" & gFSO.GetTempName)
    Set stream = gFSO.CreateTextFile(stage, False, True): created = True
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRECONCILE1", gRoot, accountRoot, storeId, _
        phase, MBC_UtcStamp(MBC_UtcNow()), gRunId))
    stream.Close: Set stream = Nothing
    MBM_FlushPath stage: digest = MBH_HashFile(stage)
    ' Only these two exact metadata tokens may be replaced, never a MSG.
    If MBM_MoveFileExW(StrPtr(stage), StrPtr(path), 9) = 0 Then _
        Err.Raise MB_ERR + 879, , "Cannot atomically publish reconciliation requirement."
    created = False
    MBM_FlushPath path
    If MBH_HashFile(path) <> digest Then Err.Raise MB_ERR + 880, , "Published reconciliation token changed."
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_WriteToken", detail
End Sub

Private Sub MBR_RequireReconcile(ByVal recoveryRoot As String, ByVal accountRoot As String, ByVal storeId As String)
    Dim leaf As String
    leaf = MBR_ReconcileLeaf(accountRoot, storeId)
    ' The independent NEEDS is durable before destructive metadata replacement.
    MBR_WriteToken MBH_JoinPath(recoveryRoot, leaf), recoveryRoot, accountRoot, storeId, "NEEDS"
    MBR_WriteToken MB_StatePath(leaf), recoveryRoot, accountRoot, storeId, "NEEDS"
    gRecoveryForceFull = True
End Sub

Private Sub MBR_CompleteReconcile(ByVal accountRoot As String, ByVal storeId As String)
    Dim recoveryRoot As String, leaf As String
    If Not gWindowCommitted Then Err.Raise MB_ERR + 881, , "Reconciliation can close only after durable mail coverage."
    If Not gRecoveryForceFull Then Exit Sub
    recoveryRoot = MBR_RecoveryRoot(accountRoot, storeId)
    leaf = MBR_ReconcileLeaf(accountRoot, storeId)
    MBR_WriteToken MB_StatePath(leaf), recoveryRoot, accountRoot, storeId, "DONE"
    MBR_WriteToken MBH_JoinPath(recoveryRoot, leaf), recoveryRoot, accountRoot, storeId, "DONE"
    ' If either write fails, the other NEEDS copy forces a safe repeat.
    gRecoveryForceFull = False
    gRecoveryNote = gRecoveryNote & " Full current-mail reconciliation completed."
End Sub

Private Function MBR_TryIndexRow(ByVal line As String, ByRef reason As String) As Boolean
    Dim f As Variant, n As Long
    On Error GoTo Rejected
    f = MBR_ParseIndexFields(line)
    If Not MB_IsHexKey(CStr(f(2))) Or Not MB_IsHexKey(CStr(f(5))) Or Not MB_IsHexKey(CStr(f(6))) Then _
        Err.Raise MB_ERR + 25, , "Invalid source/revision/file digest."
    ' This is the REAL index semantic validator, in isolated temporary caches.
    ' R rows must refer to a previously accepted same-store, same-byte record.
    MBX_CacheRecord f
    MBR_TryIndexRow = True
    Exit Function
Rejected:
    n = Err.Number: reason = Err.Description
    If (n >= MB_ERR + 15 And n <= MB_ERR + 25) Or _
       (n >= MB_ERR + 263 And n <= MB_ERR + 269) Or n = MB_ERR + 326 Or _
       (n >= MB_ERR + 600 And n <= MB_ERR + 610) Or _
       (n >= MB_ERR + 1180 And n <= MB_ERR + 1183) Or n = MBH_ERR_BASE + 2 Then Exit Function
    ' Unexpected resource/crypto/runtime failures are not corrupt-row evidence.
    Err.Raise n, "MBR_TryIndexRow", reason
End Function

Private Function MBR_SalvageJournal(ByVal recoveryRoot As String) As String
    Dim oldIndex As Object, oldExternal As Object, oldMeta As Object, oldStores As Object
    Dim oldBindings As Object, oldSources As Object, oldOwners As Object
    Dim oldBuckets As Object, oldSeen As Object, oldKnown As Object, oldHeads As Object, oldLookup As Object
    Dim inputStream As Object, outputStream As Object, source As String, stage As String, target As String, dir As String
    Dim line As String, reason As String, accepted As Long, rejected As Long, cached As Boolean
    Dim digest As String, bytes As Double, n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    source = MB_StatePath("journal.tsv")
    If Not gFSO.FileExists(source) Then Exit Function
    MBY_NoReparse source, True
    guard = MBH_GuardFile(source)
    Set oldIndex = gIndex: Set oldExternal = gExternalRoots: Set oldMeta = gRecordMeta: Set oldStores = gStoreTokens
    Set oldBindings = gLegacyBindings: Set oldSources = gLegacySources: Set oldOwners = gLegacyOwners
    Set oldBuckets = gIdentityBuckets: Set oldSeen = gIdentitySeen: Set oldKnown = gWindowKnown
    Set oldLookup = gIdentityLookupState
    Set oldHeads = gRevisionHeads
    cached = True
    Set gIndex = CreateObject("Scripting.Dictionary")
    Set gExternalRoots = CreateObject("Scripting.Dictionary")
    Set gRecordMeta = CreateObject("Scripting.Dictionary")
    Set gStoreTokens = CreateObject("Scripting.Dictionary"): gStoreTokens.CompareMode = vbTextCompare
    Set gLegacyBindings = CreateObject("Scripting.Dictionary")
    Set gLegacySources = CreateObject("Scripting.Dictionary")
    Set gLegacyOwners = CreateObject("Scripting.Dictionary"): gLegacyOwners.CompareMode = vbTextCompare
    Set gIdentityBuckets = Nothing: Set gIdentitySeen = Nothing: Set gRevisionHeads = Nothing
    Set gWindowKnown = CreateObject("Scripting.Dictionary")
    dir = MBH_JoinPath(recoveryRoot, "salvages")
    MBY_NoReparse dir, False
    MBH_EnsureDirectory dir
    target = MBH_JoinPath(dir, "journal_" & Left$(MBH_HashText(gRunId & gFSO.GetTempName), 24) & ".tsv")
    stage = MBH_NewStagePath(target)
    Set inputStream = gFSO.OpenTextFile(source, 1, False, -1)
    Set outputStream = gFSO.CreateTextFile(stage, True, True)
    Do Until inputStream.AtEndOfStream
        line = inputStream.ReadLine
        If Len(line) > 0 Then
            If MBR_TryIndexRow(line, reason) Then
                outputStream.WriteLine line: accepted = accepted + 1
            Else
                rejected = rejected + 1
                MBR_Log "WARNING", "RECOVERY_INDEX_ROW_REJECTED", CStr(accepted + rejected), source, 0, reason
            End If
        End If
    Loop
    inputStream.Close: Set inputStream = Nothing
    outputStream.Close: Set outputStream = Nothing
    If accepted > 0 Then
        MBM_FlushPath stage
        digest = MBH_HashFile(stage): bytes = CDbl(gFSO.GetFile(stage).Size)
        MBH_CommitStage stage, target, digest, bytes: stage = ""
        MBR_SalvageJournal = "SALVAGE|" & CStr(gFSO.GetFileName(target)) & "|" & _
            Format$(bytes, "0") & "|" & digest & "|" & CStr(accepted)
        gRecoverySalvagedRows = accepted
    End If
    GoTo Cleanup
Failed:
    n = Err.Number: detail = Err.Description
Cleanup:
    On Error Resume Next
    If Not inputStream Is Nothing Then inputStream.Close
    If Not outputStream Is Nothing Then outputStream.Close
    If Len(stage) > 0 Then MBH_DeleteStage stage
    MBH_CloseGuard guard
    If cached Then
        Set gIndex = oldIndex: Set gExternalRoots = oldExternal: Set gRecordMeta = oldMeta: Set gStoreTokens = oldStores
        Set gLegacyBindings = oldBindings: Set gLegacySources = oldSources: Set gLegacyOwners = oldOwners
        Set gIdentityBuckets = oldBuckets: Set gIdentitySeen = oldSeen: Set gWindowKnown = oldKnown
        Set gIdentityLookupState = oldLookup
        Set gRevisionHeads = oldHeads
    End If
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, "MBR_SalvageJournal", detail
End Function

Private Sub MBR_RestoreSalvage(ByVal recoveryRoot As String, ByVal pointer As String)
    Dim f As Variant, path As String, finalPath As String, stage As String, reason As String
    Dim n As Long, detail As String
    On Error GoTo Failed
    f = Split(pointer, "|")
    If UBound(f) <> 4 Then Err.Raise MB_ERR + 871, , "Invalid salvage pointer."
    If CStr(f(0)) <> "SALVAGE" Or Not MBR_IsStateLeaf(CStr(f(1))) Or Not MB_IsHexKey(CStr(f(3))) Then _
        Err.Raise MB_ERR + 872, , "Unsafe salvage pointer."
    path = MBH_JoinPath(MBH_JoinPath(recoveryRoot, "salvages"), CStr(f(1)))
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) <> MB_ByteCount(CStr(f(2))) Or MBH_HashFile(path) <> CStr(f(3)) Then _
        Err.Raise MB_ERR + 873, , "Salvaged historical proof no longer matches its durable intent."
    If Not MBR_CheckedStateValid(path, reason) Then Err.Raise MB_ERR + 874, , "Salvaged rows are damaged: " & reason
    finalPath = MB_StatePath("journal.tsv")
    If gFSO.FileExists(finalPath) Then MBR_PreserveFile finalPath, MB_StatePath("preserved")
    stage = MBH_NewStagePath(finalPath)
    gFSO.CopyFile path, stage, True
    MBH_CommitStage stage, finalPath, CStr(f(3)), MB_ByteCount(CStr(f(2))): stage = ""
    gRecoverySalvagedRows = CLng(MB_ByteCount(CStr(f(4))))
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Len(stage) > 0 Then MBH_DeleteStage stage
    On Error GoTo 0
    Err.Raise n, "MBR_RestoreSalvage", detail
End Sub

Private Sub MBR_RestoreState(ByVal accountRoot As String, ByVal storeId As String, _
                             Optional ByVal forceRepair As Boolean = False)
    Dim recoveryRoot As String, manifest As String, pending As String, reason As String
    Dim files As Object, chunks As Object, key As Variant, meta As Variant, blocks As Collection
    Dim emptyFiles As Object, stream As Object, healthy As Boolean, intent As String
    Dim leaf As String, needsPrimary As Boolean, needsMirror As Boolean
    Dim overlayFiles As Object, baseReference As String, baseDigest As String
    accountRoot = MBH_NormalizePath(accountRoot)
    recoveryRoot = MBR_RecoveryRoot(accountRoot, storeId)
    leaf = MBR_ReconcileLeaf(accountRoot, storeId)
    needsPrimary = MBR_TokenNeeds(MB_StatePath(leaf), accountRoot, storeId)
    needsMirror = MBR_TokenNeeds(MBH_JoinPath(recoveryRoot, leaf), accountRoot, storeId)
    If needsPrimary Or needsMirror Or gStateMigrated Then gRecoveryForceFull = True
    pending = MBR_PendingIntent(recoveryRoot, accountRoot, storeId)
    manifest = MBR_FindSnapshot(recoveryRoot, accountRoot, storeId, files, chunks)
    healthy = MBR_StateHealthy(files, reason)
    If forceRepair Then healthy = False
    If healthy And Len(pending) = 0 Then
        If gRecoveryForceFull Then
            MBR_RequireReconcile recoveryRoot, accountRoot, storeId
            gRecoveryNote = "RECOVERY_RECONCILE_PENDING: state is readable; repeat the incomplete full reconciliation without restoring again."
        End If
        Exit Sub
    End If
    MBR_RequireReconcile recoveryRoot, accountRoot, storeId
    If Left$(pending, 6) = "MERGE|" Then
        intent = pending
    Else
        ' Convert an earlier snapshot/SALVAGE/MSG_ONLY intent into one durable
        ' plan without discarding any still-available current proof.
        If Len(pending) > 0 Then
            If MBR_HasPreservedState() Then Err.Raise MB_ERR + 1157, , _
                "Legacy interrupted recovery has preserved history outside a bound overlay. Review those originals before resuming; no proof may be discarded."
            baseReference = pending
        Else
            manifest = MBR_FindSnapshot(recoveryRoot, accountRoot, storeId, files, chunks, True)
            baseReference = "MSG_ONLY"
            If Len(manifest) > 0 Then baseReference = manifest
        End If
        intent = MBR_CaptureOverlay(recoveryRoot, accountRoot, storeId, baseReference)
    End If
    MBR_ReadOverlay recoveryRoot, accountRoot, storeId, intent, overlayFiles, baseReference, baseDigest
    If baseReference <> "MSG_ONLY" And Left$(baseReference, 8) <> "SALVAGE|" Then
        MBR_LoadOverlayBase recoveryRoot, accountRoot, storeId, baseReference, baseDigest, files, chunks
    Else
        Set files = Nothing: Set chunks = Nothing
    End If
    ' Immutable raw copies + a checked bound manifest exist BEFORE START.
    ' Restart always uses that SAME plan, not the partly reconstructed primary.
    MBR_WriteIntent recoveryRoot, accountRoot, storeId, intent, "START"
    If Not files Is Nothing Then
        MBR_PreserveExtraState files
        For Each key In files.Keys
            meta = MBM_Unpack(CStr(files(key)), 3): Set blocks = chunks(key)
            MBR_ReconstructFile recoveryRoot, CStr(key), meta, blocks
        Next key
    Else
        MBR_PreserveExtraState emptyFiles
        MBH_EnsureDirectory gFSO.GetParentFolderName(MB_StatePath("journal.tsv"))
        If Left$(baseReference, 8) = "SALVAGE|" Then
            MBR_RestoreSalvage recoveryRoot, baseReference
        ElseIf Not gFSO.FileExists(MB_StatePath("journal.tsv")) Then
            Set stream = gFSO.CreateTextFile(MB_StatePath("journal.tsv"), False, True)
            stream.Close: Set stream = Nothing
            MBM_FlushPath MB_StatePath("journal.tsv")
        End If
        gRecoveryRebased = True
    End If
    MBR_ApplyOverlay recoveryRoot, overlayFiles, accountRoot, storeId
    ' No DONE is emitted if valid historical proofs conflict or any owned
    ' merge could not be published. Both original byte sets remain available.
    MBR_WriteIntent recoveryRoot, accountRoot, storeId, intent, "DONE"
    If Not files Is Nothing Then
        gRecoveryNote = "RECOVERED_SNAPSHOT: snapshot restored and surviving later proof merged; full current-mail reconciliation required."
        MBR_Log "WARNING", "RECOVERY_SNAPSHOT_RESTORED", "", baseReference, 0, gRecoveryNote
    Else
        gRecoveryNote = "MSG_ONLY_REBASELINED: surviving checked proof retained; files without historical proof require fresh content matching."
        If gRecoveryMergedRows > 0 Or gRecoverySalvagedRows > 0 Then _
            gRecoveryNote = "PARTIAL_STATE_REBUILT: surviving valid expected hashes retained; missing historical equality remains unproven."
        MBR_Log "WARNING", "RECOVERY_MSG_ONLY", "", accountRoot, 0, gRecoveryNote
    End If
    If gRecoveryOverlayRejected > 0 Then gRecoveryNote = gRecoveryNote & _
        " Some raw rows could not be proven; immutable originals remain in recovery/overlays."
    ' A restored or retained old token cannot erase this recovery obligation.
    MBR_RequireReconcile recoveryRoot, accountRoot, storeId
End Sub

Private Function MBR_Fingerprint(ByVal files As Object) As String
    Dim keys As Variant, i As Long, value As String
    If files Is Nothing Then Exit Function
    If files.Count = 0 Then Exit Function
    keys = files.Keys: MBW_SortKeys keys, LBound(keys), UBound(keys)
    For i = LBound(keys) To UBound(keys)
        value = value & MB_EncodeFields(Array(LCase$(CStr(keys(i))), CStr(files(keys(i))))) & vbLf
    Next i
    MBR_Fingerprint = MBH_HashText(value)
End Function

Private Function MBR_NextGeneration(ByVal snapshotsDir As String) As String
    Dim item As Object, candidate As String, number As Double, maximum As Double, i As Long, valid As Boolean
    For Each item In gFSO.GetFolder(snapshotsDir).Files
        candidate = CStr(item.Name)
        If LCase$(Left$(candidate, 9)) = "snapshot_" And Len(candidate) > 22 Then
            candidate = Mid$(candidate, 10, 12): valid = True
            For i = 1 To Len(candidate)
                If Mid$(candidate, i, 1) < "0" Or Mid$(candidate, i, 1) > "9" Then valid = False
            Next i
            If valid Then
                number = MB_ByteCount(candidate)
                If number > maximum Then maximum = number
            End If
        End If
    Next item
    If maximum >= 999999999998# Then Err.Raise MB_ERR + 864, , "Recovery generation counter is exhausted."
    MBR_NextGeneration = Format$(maximum + 1, "000000000000")
End Function

Private Sub MBR_SaveSnapshot(ByVal accountRoot As String, ByVal storeId As String)
    Dim recoveryRoot As String, stateDir As String, snapshotsDir As String, manifest As String, stage As String
    Dim files As Object, chunks As Object, previousFiles As Object, previousChunks As Object, items As Object
    Dim item As Object, key As Variant, keys As Variant, i As Long, blocks As Collection, block As Variant
    Dim bytes As Double, digest As String, reason As String, stream As Object, count As Long
    Dim n As Long, detail As String, previous As String, meta As Variant
    On Error GoTo Failed
    ' Caller closes ALL writers and holds the account/root locks until return.
    If Not gFSO.FileExists(MB_StatePath("journal.tsv")) Then Exit Sub
    recoveryRoot = MBR_RecoveryRoot(accountRoot, storeId)
    stateDir = gFSO.GetParentFolderName(MB_StatePath("journal.tsv"))
    Set files = CreateObject("Scripting.Dictionary"): files.CompareMode = vbTextCompare
    Set chunks = CreateObject("Scripting.Dictionary"): chunks.CompareMode = vbTextCompare
    Set items = CreateObject("Scripting.Dictionary"): items.CompareMode = vbTextCompare
    For Each item In gFSO.GetFolder(stateDir).Files
        If MBR_IsStateLeaf(CStr(item.Name)) Then items(CStr(item.Name)) = CStr(item.Path)
    Next item
    If Not items.Exists("journal.tsv") Then Exit Sub
    keys = items.Keys: MBW_SortKeys keys, LBound(keys), UBound(keys)
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        If Not MBR_CheckedStateValid(CStr(items(key)), reason) Then _
            Err.Raise MB_ERR + 856, , "Invalid primary state is not promoted to recovery: " & CStr(key) & ": " & reason
        Set blocks = MBR_CaptureState(CStr(items(key)), recoveryRoot, bytes, digest)
        files.Add CStr(key), MB_EncodeFields(Array(Format$(bytes, "0"), digest, CStr(blocks.Count)))
        chunks.Add CStr(key), blocks
    Next i
    previous = MBR_FindSnapshot(recoveryRoot, accountRoot, storeId, previousFiles, previousChunks)
    If Len(previous) > 0 Then
        If MBR_Fingerprint(files) = MBR_Fingerprint(previousFiles) Then Exit Sub
    End If
    snapshotsDir = MBH_JoinPath(recoveryRoot, "snapshots")
    MBY_NoReparse snapshotsDir, False
    MBH_EnsureDirectory snapshotsDir
    manifest = MBH_JoinPath(snapshotsDir, "snapshot_" & MBR_NextGeneration(snapshotsDir) & "_" & MBC_FileStamp(MBC_UtcNow()) & "_" & _
                            Left$(MBH_HashText(gRunId & gFSO.GetTempName), 12) & ".tsv")
    stage = MBH_NewStagePath(manifest)
    Set stream = gFSO.CreateTextFile(stage, True, True)
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRSNAP1", "HEADER", gRoot, accountRoot, storeId, MBC_UtcStamp(MBC_UtcNow()), gRunId))
    For i = LBound(keys) To UBound(keys)
        key = keys(i): Set blocks = chunks(key)
        meta = MBM_Unpack(CStr(files(key)), 3)
        MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRSNAP1", "FILE", CStr(key), meta(0), meta(1), meta(2)))
        For Each block In blocks
            MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRSNAP1", "CHUNK", CStr(key), block(0), block(1), block(2)))
            count = count + 1
        Next block
    Next i
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBRSNAP1", "END", CStr(files.Count), CStr(count)))
    stream.Close: Set stream = Nothing
    MBM_FlushPath stage
    MBH_CommitStage stage, manifest: stage = ""
    gRecoverySnapshotsSaved = gRecoverySnapshotsSaved + 1
    MBR_Log "INFO", "RECOVERY_SNAPSHOT_SAVED", "", manifest, 0, _
                "Metadata-only immutable chunks; MSG files were neither copied nor changed."
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(stage) > 0 Then MBH_DeleteStage stage
    On Error GoTo 0
    Err.Raise n, "MBR_SaveSnapshot", detail
End Sub

Private Function MBR_BaselinePath(ByVal accountRoot As String, ByVal storeId As String) As String
    MBR_BaselinePath = MB_StatePath("baseline_" & Left$(MBH_HashText(MB_EncodeFields(Array(gRoot, accountRoot, storeId))), 32) & ".tsv")
End Function

Private Sub MBR_LoadBaseline(ByVal accountRoot As String, ByVal storeId As String)
    Dim path As String, stream As Object, line As String, f As Variant, resolved As String, root As String
    Dim existing As Variant, n As Long, detail As String
    On Error GoTo Failed
    If gRecoveryFiles Is Nothing Then MBR_Reset
    path = MBR_BaselinePath(accountRoot, storeId)
    If Not gFSO.FileExists(path) Then Exit Sub
    MBY_NoReparse path, True
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            f = MB_ReadCheckedLine(line, 9)
            If CStr(f(0)) <> "MBBASE1" Or StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Or _
               StrComp(CStr(f(2)), storeId, vbTextCompare) <> 0 Or StrComp(CStr(f(3)), accountRoot, vbTextCompare) <> 0 Or _
               CStr(f(8)) <> "MSG_ONLY_REBASELINED" Or Not MB_IsHexKey(CStr(f(6))) Or MB_ByteCount(CStr(f(5))) <= 0 Then _
                Err.Raise MB_ERR + 857, , "Baseline ownership/hash/provenance is invalid."
            root = MBM_ResolvedRoot(CStr(f(4)), CStr(f(6)), accountRoot)
            resolved = MBY_CanonicalFilePath(MBM_ResolvePath(CStr(f(4)), CStr(f(6))), root)
            If gRecoveryFiles.Exists(resolved) Then
                existing = MBM_Unpack(CStr(gRecoveryFiles(resolved)), 5)
                If CStr(existing(2)) <> CStr(f(6)) Or MB_ByteCount(CStr(existing(1))) <> MB_ByteCount(CStr(f(5))) Then _
                    Err.Raise MB_ERR + 858, , "A baseline path has conflicting historical expected bytes."
            End If
            gRecoveryFiles(resolved) = MB_EncodeFields(Array(resolved, f(5), f(6), root, f(8)))
            MBR_AddDigestPath resolved, MB_ByteCount(CStr(f(5))), CStr(f(6)), root
        End If
    Loop
    stream.Close: Set stream = Nothing
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBR_LoadBaseline", detail
End Sub

Private Function MBR_CaptureBaseline(ByVal path As String, ByVal accountRoot As String, ByVal storeId As String, _
                                     ByRef reason As String) As Boolean
    Dim obj As Object, bytes As Double, digest As String, stream As Object, payload As String, stamp As Date
    Dim n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    guard = 0
    On Error GoTo Invalid
    path = MBY_CanonicalFilePath(path, accountRoot, True)
    guard = MBH_GuardFile(path)
    bytes = CDbl(gFSO.GetFile(path).Size)
    If bytes <= 0 Then Err.Raise MB_ERR + 859, , "MSG-only file is empty."
    digest = MBH_HashFile(path)
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 860, , "MSG-only file does not reopen as MailItem."
    stamp = obj.SentOn
    If MBC_ValidDate(stamp) Then gRecoverySentHints(path) = MBC_FileStamp(stamp)
    obj.Close olDiscard: Set obj = Nothing
    If CDbl(gFSO.GetFile(path).Size) <> bytes Or MBH_HashFile(path) <> digest Then _
        Err.Raise MB_ERR + 861, , "MSG-only file changed during baseline creation."
    payload = MB_EncodeFields(Array("MBBASE1", gRoot, storeId, accountRoot, path, Format$(bytes, "0"), _
                              digest, MBC_UtcStamp(MBC_UtcNow()), "MSG_ONLY_REBASELINED"))
    MBY_NoReparse MBR_BaselinePath(accountRoot, storeId), False
    Set stream = gFSO.OpenTextFile(MBR_BaselinePath(accountRoot, storeId), 8, True, -1)
    stream.WriteLine "": MB_WriteCheckedLine stream, payload
    stream.Close: Set stream = Nothing
    mbrBaselineWrites = mbrBaselineWrites + 1
    If mbrBaselineWrites >= MB_JOURNAL_FLUSH_EVERY Then
        MBM_FlushPath MBR_BaselinePath(accountRoot, storeId)
        mbrBaselineWrites = 0
    End If
    gRecoveryFiles(path) = MB_EncodeFields(Array(path, Format$(bytes, "0"), digest, accountRoot, "MSG_ONLY_REBASELINED"))
    MBR_AddDigestPath path, bytes, digest, accountRoot
    MBH_CloseGuard guard
    gRecoveryNewBaselines = gRecoveryNewBaselines + 1
    MBR_CaptureBaseline = True
    Exit Function
Invalid:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not obj Is Nothing Then obj.Close olDiscard
    If Not stream Is Nothing Then stream.Close
    MBH_CloseGuard guard
    On Error GoTo 0
    reason = "Error " & CStr(n) & ": " & detail
End Function

Private Sub MBR_AddDigestPath(ByVal path As String, ByVal bytes As Double, ByVal digest As String, ByVal accountRoot As String)
    Dim key As String, paths As Collection, oldPath As String, candidate As Variant
    If mbrDigestFiles Is Nothing Then
        Set mbrDigestFiles = CreateObject("Scripting.Dictionary")
        mbrDigestFiles.CompareMode = vbBinaryCompare
    End If
    key = LCase$(MBH_NormalizePath(accountRoot)) & "|" & digest & "|" & Format$(bytes, "0")
    path = MBH_NormalizePath(path)
    If mbrDigestFiles.Exists(key) Then
        If IsObject(mbrDigestFiles(key)) Then
            Set paths = mbrDigestFiles(key)
            For Each candidate In paths
                If StrComp(CStr(candidate), path, vbTextCompare) = 0 Then Exit Sub
            Next candidate
            paths.Add path
        Else
            oldPath = CStr(mbrDigestFiles(key))
            If StrComp(oldPath, path, vbTextCompare) = 0 Then Exit Sub
            Set paths = New Collection
            paths.Add oldPath: paths.Add path
            mbrDigestFiles.Remove key
            mbrDigestFiles.Add key, paths
        End If
    Else
        ' One string per ordinary unique hash; allocate a collection only when
        ' a second distinct path has the same expected bytes. Avoid 100k COM maps.
        mbrDigestFiles.Add key, path
    End If
End Sub

Private Function MBR_FindExpectedFile(ByVal digest As String, ByVal bytes As Double, ByVal accountRoot As String) As String
    Dim key As String, paths As Collection, candidate As Variant, chosen As String, count As Long
    If mbrDigestFiles Is Nothing Then Exit Function
    If Not MB_IsHexKey(digest) Or bytes <= 0 Then Exit Function
    key = LCase$(MBH_NormalizePath(accountRoot)) & "|" & digest & "|" & Format$(bytes, "0")
    If Not mbrDigestFiles.Exists(key) Then Exit Function
    If IsObject(mbrDigestFiles(key)) Then
        Set paths = mbrDigestFiles(key)
    Else
        Set paths = New Collection
        paths.Add CStr(mbrDigestFiles(key))
    End If
    For Each candidate In paths
        If gFSO.FileExists(CStr(candidate)) Then
            MBY_CanonicalFilePath CStr(candidate), accountRoot, True
            If CDbl(gFSO.GetFile(CStr(candidate)).Size) = bytes Then
                If MBH_HashFile(CStr(candidate)) = digest Then
                    chosen = CStr(candidate): count = count + 1
                    If count > 1 Then Exit Function
                End If
            End If
        End If
    Next candidate
    If count = 1 Then MBR_FindExpectedFile = chosen
End Function

Private Sub MBR_MergeStageProof(ByVal accountRoot As String)
    Dim key As Variant, f As Variant, existing As Variant, root As String, path As String
    If gStageExpected Is Nothing Then Exit Sub
    For Each key In gStageExpected.Keys
        f = MBM_Unpack(CStr(gStageExpected(key)), 5)
        If CStr(f(4)) = "VERIFIED_PUBLICATION" And LCase$(Right$(CStr(f(0)), 4)) = ".msg" Then
            root = MBM_ResolvedRoot(CStr(f(0)), CStr(f(2)), CStr(f(3)))
            path = MBY_CanonicalFilePath(MBM_ResolvePath(CStr(f(0)), CStr(f(2))), root)
            If StrComp(root, accountRoot, vbTextCompare) = 0 Then
                If gRecoveryFiles.Exists(path) Then
                    existing = MBM_Unpack(CStr(gRecoveryFiles(path)), 5)
                    If CStr(existing(2)) <> CStr(f(2)) Or MB_ByteCount(CStr(existing(1))) <> MB_ByteCount(CStr(f(1))) Then
                        gRecoveryReviewCount = gRecoveryReviewCount + 1
                        MBR_Log "WARNING", "RECOVERY_EXPECTED_PROOF_CONFLICT", "", path, 0, _
                            "Prior baseline and verified publication disagree; neither is silently overwritten."
                    Else
                        gRecoveryFiles(path) = MB_EncodeFields(Array(path, f(1), f(2), root, "VERIFIED_PUBLICATION"))
                    End If
                Else
                    gRecoveryFiles(path) = MB_EncodeFields(Array(path, f(1), f(2), root, "VERIFIED_PUBLICATION"))
                End If
                MBR_AddDigestPath path, MB_ByteCount(CStr(f(1))), CStr(f(2)), root
            End If
        End If
    Next key
End Sub

Private Sub MBR_BuildBaseline(ByVal accountRoot As String, ByVal storeId As String)
    Dim pending As New Collection, known As Object, folder As Object, child As Object, item As Object
    Dim key As Variant, fields As Variant, path As String, current As String, reason As String
    Dim n As Long, detail As String
    On Error GoTo Failed
    If gRecoveryFiles Is Nothing Then MBR_Reset
    MBR_LoadBaseline accountRoot, storeId
    MBR_MergeStageProof accountRoot
    Set known = CreateObject("Scripting.Dictionary"): known.CompareMode = vbTextCompare
    For Each key In gIndex.Keys
        fields = MB_IndexFields(CStr(key))
        known(MBX_IndexedPath(CStr(key), fields)) = True
    Next key
    pending.Add accountRoot
    Do While pending.Count > 0
        If Not MB_ContinueWork() Then GoTo Finish
        current = CStr(pending(pending.Count)): pending.Remove pending.Count
        MBY_NoReparse current, True
        If Not MBR_IsRecoveryDirectory(current) And Not MB_PathWithin(current, gRoot) Then
            Set folder = gFSO.GetFolder(current)
            For Each item In folder.Files
                If Not MB_ContinueWork() Then GoTo Finish
                If LCase$(gFSO.GetExtensionName(CStr(item.Name))) = "msg" Then
                    path = MBH_NormalizePath(CStr(item.Path))
                    If Not known.Exists(path) And Not gRecoveryFiles.Exists(path) Then
                        If Not MBR_CaptureBaseline(path, accountRoot, storeId, reason) Then
                            gRecoveryReviewCount = gRecoveryReviewCount + 1
                            MBR_Log "WARNING", "RECOVERY_BASELINE_REVIEW", "", path, 0, reason
                        End If
                    End If
                End If
            Next item
            For Each child In folder.SubFolders
                If (CLng(child.Attributes) And 1024) <> 0 Then _
                    Err.Raise MB_ERR + 862, , "Recovery inventory cannot traverse a reparse folder."
                pending.Add CStr(child.Path)
            Next child
        End If
    Loop
Finish:
    If gFSO.FileExists(MBR_BaselinePath(accountRoot, storeId)) Then
        MBM_FlushPath MBR_BaselinePath(accountRoot, storeId)
        mbrBaselineWrites = 0
    End If
    If gRecoveryNewBaselines > 0 Then
        MBR_Log "WARNING", "RECOVERY_NEW_BASELINE", "", accountRoot, 0, _
            "Newly readable MSG baselines=" & CStr(gRecoveryNewBaselines) & _
            "; no statement about equality to deleted historical hashes."
    End If
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If gFSO.FileExists(MBR_BaselinePath(accountRoot, storeId)) Then MBM_FlushPath MBR_BaselinePath(accountRoot, storeId)
    On Error GoTo 0
    Err.Raise n, "MBR_BuildBaseline", detail
End Sub


' ==================== Independent physical archive audit ===================
' Committed index hashes and immutable MSG-only recovery baselines are distinct
' proof origins. A fresh baseline establishes current bytes/readability only.
' Existing expected hashes are never replaced by a newly observed file hash.
' One physical path+size+hash proof is read once per run even if many mailbox
' identities refer to it. Missing sources do not remove their file from audit.
Private Sub MBA_Reset()
    Set gArchiveAuditEntries = Nothing: Set gArchiveAuditIssues = Nothing
    Set gArchiveAuditCheckedPaths = Nothing: Set gArchiveAuditVisited = Nothing
    gArchiveAuditStatePath = "": gArchiveAuditCursor = ""
    gArchiveAuditStore = "": gArchiveAuditAccount = ""
    gArchiveAuditChecked = 0: gArchiveAuditPassed = 0: gArchiveAuditFailed = 0
    gArchiveAuditPending = 0: gArchiveAuditFilesRead = 0
    gArchiveAuditReferenceReview = 0: gArchiveAuditStateIssues = 0
    gArchiveAuditTotal = 0: gArchiveAuditCommitted = False
End Sub

Private Function MBA_ProofId(ByVal path As String, ByVal bytes As String, ByVal digest As String) As String
    MBA_ProofId = MBH_HashText(MB_EncodeFields(Array(LCase$(MBH_NormalizePath(path)), _
        Format$(MB_ByteCount(bytes), "0"), LCase$(digest))))
End Function

Private Sub MBA_Initialize()
    Set gArchiveAuditEntries = CreateObject("Scripting.Dictionary")
    Set gArchiveAuditIssues = CreateObject("Scripting.Dictionary")
    Set gArchiveAuditCheckedPaths = CreateObject("Scripting.Dictionary")
    Set gArchiveAuditVisited = CreateObject("Scripting.Dictionary")
    gArchiveAuditEntries.CompareMode = vbBinaryCompare
    gArchiveAuditIssues.CompareMode = vbBinaryCompare
    gArchiveAuditCheckedPaths.CompareMode = vbBinaryCompare
    gArchiveAuditVisited.CompareMode = vbBinaryCompare
    gArchiveAuditStore = gWindowStore
    gArchiveAuditAccount = MBH_NormalizePath(gAutoAccount)
    gArchiveAuditStatePath = MB_StatePath("archive_audit_" & _
        Left$(MBH_HashText(MB_EncodeFields(Array(gArchiveAuditStore, LCase$(gArchiveAuditAccount)))), 32) & ".tsv")
End Sub

Private Sub MBA_AddEntry(ByVal path As String, ByVal bytes As String, ByVal digest As String, _
                         ByVal boundRoot As String, ByVal isMsg As Boolean, _
                         ByVal reference As String, ByVal provenance As String)
    Dim id As String, prior As Variant
    path = MBY_CanonicalFilePath(path, boundRoot)
    If Not MB_IsHexKey(digest) Then Err.Raise MB_ERR + 751, , "Invalid physical-audit expected digest."
    If MB_ByteCount(bytes) < 0 Then Err.Raise MB_ERR + 752, , "Negative physical-audit expected size."
    id = MBA_ProofId(path, bytes, digest)
    If gArchiveAuditEntries.Exists(id) Then
        prior = MBM_Unpack(CStr(gArchiveAuditEntries(id)), 8)
        ' The physical bytes are equal; preserve the stronger committed origin.
        If CStr(prior(6)) = "COMMITTED_INDEX" Then Exit Sub
        If provenance <> "COMMITTED_INDEX" Then Exit Sub
    End If
    gArchiveAuditEntries(id) = MB_EncodeFields(Array(path, bytes, LCase$(digest), _
        MBH_NormalizePath(boundRoot), IIf(isMsg, "1", "0"), reference, provenance, ""))
End Sub

Private Sub MBA_AddIndexReference(ByVal key As String, ByVal storeToken As String)
    Dim f As Variant, metadata As Variant, path As String, root As String, isMsg As Boolean
    Dim n As Long, detail As String
    On Error GoTo Review
    If Not gRecordMeta.Exists(key) Then Err.Raise MB_ERR + 753, , "Indexed reference has no account metadata."
    metadata = MBX_Meta(key)
    If CStr(metadata(0)) <> storeToken Then Exit Sub
    f = MB_IndexFields(key)
    path = MBX_IndexedPath(key, f)
    root = MBM_ResolvedRoot(MBX_IndexedOriginal(key, f), CStr(f(2)), MBX_IndexedRoot(key))
    isMsg = (Left$(key, 2) = "M|" Or Left$(key, 2) = "L|")
    MBA_AddEntry path, CStr(f(1)), CStr(f(2)), root, isMsg, key, "COMMITTED_INDEX"
    Exit Sub
Review:
    n = Err.Number: detail = Err.Description
    gArchiveAuditReferenceReview = gArchiveAuditReferenceReview + 1
    MB_LogEvent "WARNING", "ARCHIVE_REFERENCE_REVIEW", key, path, n, detail
End Sub

Private Sub MBA_AddRecoveryReference(ByVal key As String)
    Dim f As Variant, path As String, root As String, n As Long, detail As String
    On Error GoTo Review
    f = MBM_Unpack(CStr(gRecoveryFiles(key)), 5)
    root = MBM_ResolvedRoot(CStr(f(0)), CStr(f(2)), CStr(f(3)))
    If StrComp(root, gArchiveAuditAccount, vbTextCompare) <> 0 Then Exit Sub
    path = MBM_ResolvePath(CStr(f(0)), CStr(f(2)))
    MBA_AddEntry path, CStr(f(1)), CStr(f(2)), root, True, "", CStr(f(4))
    Exit Sub
Review:
    n = Err.Number: detail = Err.Description
    gArchiveAuditReferenceReview = gArchiveAuditReferenceReview + 1
    MB_LogEvent "WARNING", "RECOVERY_BASELINE_REVIEW", key, path, n, detail
End Sub

Private Sub MBA_AddStageReference(ByVal key As String)
    Dim f As Variant, path As String, root As String, n As Long, detail As String
    On Error GoTo Review
    f = MBM_Unpack(CStr(gStageExpected(key)), 5)
    If CStr(f(4)) <> "VERIFIED_PUBLICATION" Then Exit Sub
    root = MBM_ResolvedRoot(CStr(f(0)), CStr(f(2)), CStr(f(3)))
    If StrComp(root, gArchiveAuditAccount, vbTextCompare) <> 0 Then Exit Sub
    path = MBM_ResolvePath(CStr(f(0)), CStr(f(2)))
    MBA_AddEntry path, CStr(f(1)), CStr(f(2)), root, (LCase$(Right$(path, 4)) = ".msg"), "", "VERIFIED_PUBLICATION"
    Exit Sub
Review:
    n = Err.Number: detail = Err.Description
    gArchiveAuditReferenceReview = gArchiveAuditReferenceReview + 1
    MB_LogEvent "WARNING", "PUBLICATION_PROOF_REVIEW", key, path, n, detail
End Sub

Private Sub MBA_ResolvePending()
    Dim keys As Variant, key As Variant, f As Variant, root As String, path As String, id As String
    Dim prior As Variant
    If gArchiveAuditIssues.Count = 0 Then Exit Sub
    keys = gArchiveAuditIssues.Keys
    For Each key In keys
        f = MBM_Unpack(CStr(gArchiveAuditIssues(key)), 8)
        root = MBM_ResolvedRoot(CStr(f(0)), CStr(f(2)), CStr(f(3)))
        path = MBM_ResolvePath(CStr(f(0)), CStr(f(2)))
        path = MBY_CanonicalFilePath(path, root)
        id = MBA_ProofId(path, CStr(f(1)), CStr(f(2)))
        If id <> CStr(key) Then
            f(0) = path: f(3) = root
            If gArchiveAuditIssues.Exists(id) Then
                prior = MBM_Unpack(CStr(gArchiveAuditIssues(id)), 8)
                If CStr(prior(6)) = "COMMITTED_INDEX" Then f = prior
            End If
            ' Relocation changes no expected byte/hash proof and resolves no
            ' issue by itself. The relocated file still has to pass its audit.
            gArchiveAuditIssues(id) = MB_EncodeFields(f)
            gArchiveAuditIssues.Remove CStr(key)
        End If
    Next key
End Sub

Private Sub MBA_BuildEntries()
    Dim key As Variant, storeToken As String
    storeToken = MBX_StoreToken(gArchiveAuditStore)
    For Each key In gIndex.Keys
        If Not MB_ContinueWork() Then Exit Sub
        MBA_AddIndexReference CStr(key), storeToken
    Next key
    If Not gRecoveryFiles Is Nothing Then
        For Each key In gRecoveryFiles.Keys
            If Not MB_ContinueWork() Then Exit Sub
            MBA_AddRecoveryReference CStr(key)
        Next key
    End If
    If Not gStageExpected Is Nothing Then
        For Each key In gStageExpected.Keys
            If Not MB_ContinueWork() Then Exit Sub
            MBA_AddStageReference CStr(key)
        Next key
    End If
    ' A missing index/baseline reference does not erase an unresolved proof.
    ' It remains in the audit set until the recorded bytes actually verify.
    For Each key In gArchiveAuditIssues.Keys
        If Not gArchiveAuditEntries.Exists(CStr(key)) Then
            gArchiveAuditEntries(CStr(key)) = CStr(gArchiveAuditIssues(key))
        End If
    Next key
    gArchiveAuditTotal = gArchiveAuditEntries.Count
End Sub

Private Function MBA_ReadStateRow(ByVal line As String, ByRef f As Variant) As Boolean
    Dim stamp As Date, size As Double
    On Error GoTo Invalid
    f = MB_ReadCheckedLine(line, 17)
    If CStr(f(0)) <> "MBARCHIVEAUDIT1" Then Exit Function
    If StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Or _
       StrComp(CStr(f(2)), gArchiveAuditStore, vbBinaryCompare) <> 0 Or _
       StrComp(CStr(f(3)), gArchiveAuditAccount, vbTextCompare) <> 0 Then Exit Function
    stamp = MBC_ParseUtcStamp(CStr(f(15)))
    If Not MBC_ValidDate(stamp) Or Len(CStr(f(16))) = 0 Then Exit Function
    If CStr(f(4)) = "CURSOR" Then
        If CStr(f(13)) <> "CURSOR" Then Exit Function
        If Len(CStr(f(5))) > 0 Then
            If Not MB_IsHexKey(CStr(f(5))) Then Exit Function
        End If
    ElseIf CStr(f(4)) = "RESULT" Then
        If CStr(f(13)) <> "REVIEW" And CStr(f(13)) <> "OK" Then Exit Function
        If Not MB_IsHexKey(CStr(f(5))) Or Not MB_IsHexKey(CStr(f(8))) Then Exit Function
        If Len(CStr(f(6))) = 0 Or Len(CStr(f(9))) = 0 Then Exit Function
        If CStr(f(10)) <> "0" And CStr(f(10)) <> "1" Then Exit Function
        size = MB_ByteCount(CStr(f(7)))
        If MBA_ProofId(CStr(f(6)), CStr(f(7)), CStr(f(8))) <> CStr(f(5)) Then Exit Function
        If StrComp(MBH_NormalizePath(CStr(f(9))), gArchiveAuditAccount, vbTextCompare) <> 0 And _
           StrComp(MBH_NormalizePath(CStr(f(9))), gRoot, vbTextCompare) <> 0 Then
            ' A registered path can be in a separately selected root. The exact
            ' root still has to be a current known owner before any file read.
            If Not MBA_KnownRoot(CStr(f(9))) Then Exit Function
        End If
    Else
        Exit Function
    End If
    MBA_ReadStateRow = True
Invalid:
End Function

Private Function MBA_KnownRoot(ByVal root As String) As Boolean
    Dim item As Variant
    If gArchiveRoots Is Nothing Then Exit Function
    For Each item In gArchiveRoots
        If StrComp(MBH_NormalizePath(CStr(item)), MBH_NormalizePath(root), vbTextCompare) = 0 Then
            MBA_KnownRoot = True
            Exit Function
        End If
    Next item
End Function

Private Sub MBA_LoadState()
    Dim stream As Object, line As String, f As Variant, n As Long, detail As String
    Dim sawCursor As Boolean
    On Error GoTo Failed
    MBY_NoReparse gArchiveAuditStatePath, False
    If Not gFSO.FileExists(gArchiveAuditStatePath) Then Exit Sub
    Set stream = gFSO.OpenTextFile(gArchiveAuditStatePath, 1, False, -1)
    Do Until stream.AtEndOfStream
        If Not MB_ContinueWork() Then stream.Close: Exit Sub
        line = stream.ReadLine
        If Len(line) > 0 Then
            If sawCursor Then Err.Raise MB_ERR + 760, , "Unexpected data after archive-audit snapshot terminator."
            If Not MBA_ReadStateRow(line, f) Then
                gArchiveAuditStateIssues = gArchiveAuditStateIssues + 1
                Err.Raise MB_ERR + 754, , "Physical audit state is damaged; recover its last complete snapshot before resuming."
            End If
            If CStr(f(4)) = "CURSOR" Then
                gArchiveAuditCursor = CStr(f(5))
                sawCursor = True
            ElseIf CStr(f(13)) = "REVIEW" Then
                gArchiveAuditIssues(CStr(f(5))) = MB_EncodeFields(Array(f(6), f(7), f(8), f(9), f(10), f(11), f(12), f(14)))
            Else
                If gArchiveAuditIssues.Exists(CStr(f(5))) Then gArchiveAuditIssues.Remove CStr(f(5))
            End If
        End If
    Loop
    stream.Close: Set stream = Nothing
    If Not sawCursor Then Err.Raise MB_ERR + 761, , "Incomplete archive-audit snapshot: terminal cursor is missing."
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBA_LoadState", detail
End Sub

#If VBA7 Then
Private Function MBA_CheckProof(ByVal f As Variant, ByRef actualPath As String, ByRef reason As String, _
                                ByRef guard As LongPtr) As Boolean
#Else
Private Function MBA_CheckProof(ByVal f As Variant, ByRef actualPath As String, ByRef reason As String, _
                                ByRef guard As Long) As Boolean
#End If
    Dim root As String, cacheId As String, obj As Object, file As Object
    Dim bytes As Double, beforeDate As Date, digest As String, n As Long
    On Error GoTo Review
    reason = ""
    bytes = MB_ByteCount(CStr(f(1)))
    root = MBM_ResolvedRoot(CStr(f(0)), CStr(f(2)), CStr(f(3)))
    actualPath = MBM_ResolvePath(CStr(f(0)), CStr(f(2)))
    actualPath = MBY_CanonicalFilePath(actualPath, root, True)
    guard = MBH_GuardFile(actualPath)
    Set file = gFSO.GetFile(actualPath)
    If CDbl(file.Size) <> bytes Then Err.Raise MB_ERR + 755, , "Archived file size differs from its recorded baseline."
    beforeDate = file.DateLastModified
    cacheId = MBA_ProofId(actualPath, CStr(f(1)), CStr(f(2)))
    ' No mtime/size shortcut: another writer could preserve both. The read
    ' guard remains held until this item's result has been recorded by caller.
    gArchiveAuditFilesRead = gArchiveAuditFilesRead + 1
    digest = MBH_HashFile(actualPath)
    If digest <> CStr(f(2)) Then Err.Raise MB_ERR + 756, , "Archived SHA-256 differs from its recorded baseline."
    If CStr(f(4)) = "1" Then
        Set obj = Application.Session.OpenSharedItem(actualPath)
        If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 757, , "Archived MSG cannot reopen as a MailItem."
        obj.Close olDiscard
        Set obj = Nothing
    End If
    Set file = gFSO.GetFile(actualPath)
    If CDbl(file.Size) <> bytes Or file.DateLastModified <> beforeDate Then _
        Err.Raise MB_ERR + 758, , "Archived file changed during verification."
    If MBH_HashFile(actualPath) <> CStr(f(2)) Then _
        Err.Raise MB_ERR + 759, , "Archived file bytes changed while reopening."
    gArchiveAuditCheckedPaths(cacheId) = True
    MBA_CheckProof = True
    Exit Function
Review:
    n = Err.Number: reason = "Error " & CStr(n) & ": " & Err.Description
    On Error Resume Next
    If Not obj Is Nothing Then obj.Close olDiscard
    Set obj = Nothing
    On Error GoTo 0
End Function

Private Function MBA_ResultFields(ByVal id As String, ByVal f As Variant, ByVal outcome As String, _
                                  ByVal detail As String) As Variant
    MBA_ResultFields = Array("MBARCHIVEAUDIT1", gRoot, gArchiveAuditStore, gArchiveAuditAccount, _
        "RESULT", id, f(0), f(1), f(2), f(3), f(4), f(5), f(6), outcome, detail, MBC_UtcStamp(MBC_UtcNow()), gRunId)
End Function

Private Sub MBA_AuditOne(ByVal id As String, ByVal report As Object)
    Dim f As Variant, actualPath As String, reason As String, ok As Boolean
    Dim n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    If gArchiveAuditVisited.Exists(id) Then Exit Sub
    f = MBM_Unpack(CStr(gArchiveAuditEntries(id)), 8)
    ok = MBA_CheckProof(f, actualPath, reason, guard)
    If ok Then
        MB_WriteCheckedLine report, MB_EncodeFields(MBA_ResultFields(id, f, "OK", _
            "Recorded bytes and readability verified at " & actualPath))
        If gArchiveAuditIssues.Exists(id) Then gArchiveAuditIssues.Remove id
        gArchiveAuditPassed = gArchiveAuditPassed + 1
    Else
        MB_WriteCheckedLine report, MB_EncodeFields(MBA_ResultFields(id, f, "REVIEW", reason))
        f(7) = reason
        gArchiveAuditIssues(id) = MB_EncodeFields(f)
        gArchiveAuditFailed = gArchiveAuditFailed + 1
        MB_LogEvent "WARNING", "ARCHIVE_FILE_REVIEW", CStr(f(5)), actualPath, 0, reason
    End If
    MBH_CloseGuard guard
    gArchiveAuditVisited(id) = True
    gArchiveAuditChecked = gArchiveAuditChecked + 1
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    Err.Raise n, "MBA_AuditOne", detail
End Sub

Private Sub MBA_Run()
    Dim keys As Variant, pending As Variant, key As Variant, i As Long, first As Long, visited As Long
    Dim count As Long, cursor As String, report As Object, reportPath As String
    Dim snapshot As Collection, f As Variant, n As Long, detail As String, retries As Long
    If gMode <> "AUTO" Or Len(gAutoAccount) = 0 Or Len(gWindowStore) = 0 Then Exit Sub
    If gIndex Is Nothing Or gFSO Is Nothing Or gFatal Or gArchiveAuditCommitted Then Exit Sub
    On Error GoTo Failed
    MBA_Initialize
    MBA_LoadState
    MBA_ResolvePending
    gArchiveAuditPending = gArchiveAuditIssues.Count
    If gCancel Then Exit Sub
    MBA_BuildEntries
    If gCancel Then Exit Sub
    reportPath = MBH_JoinPath(gRunDir, "archive_audit.tsv")
    Set report = gFSO.CreateTextFile(reportPath, False, True)
    ' Retry a bounded group promptly; the independent full rotation below
    ' guarantees progress even when these old errors never disappear.
    If gArchiveAuditIssues.Count > 0 Then
        pending = gArchiveAuditIssues.Keys
        MBW_SortKeys pending, LBound(pending), UBound(pending)
        For Each key In pending
            If retries >= MBA_RETRY_BUDGET Then Exit For
            If Not MB_ContinueWork() Then Exit For
            MBA_AuditOne CStr(key), report
            retries = retries + 1
        Next key
    End If
    cursor = gArchiveAuditCursor
    count = gArchiveAuditEntries.Count
    If count > 0 Then
        keys = gArchiveAuditEntries.Keys
        MBW_SortKeys keys, LBound(keys), UBound(keys)
        first = LBound(keys)
        If Len(cursor) > 0 Then
            Do While first <= UBound(keys)
                If StrComp(CStr(keys(first)), cursor, vbBinaryCompare) > 0 Then Exit Do
                first = first + 1
            Loop
            If first > UBound(keys) Then first = LBound(keys)
        End If
        i = first
        Do While visited < count And gArchiveAuditChecked < MBA_FILE_BUDGET
            If Not MB_ContinueWork() Then Exit Do
            MBA_AuditOne CStr(keys(i)), report
            cursor = CStr(keys(i))
            visited = visited + 1
            i = i + 1
            If i > UBound(keys) Then i = LBound(keys)
        Loop
    End If
    report.Close: Set report = Nothing
    MBM_FlushPath reportPath
    Set snapshot = New Collection
    For Each key In gArchiveAuditIssues.Keys
        f = MBM_Unpack(CStr(gArchiveAuditIssues(key)), 8)
        snapshot.Add MBW_CheckedText(MBA_ResultFields(CStr(key), f, "REVIEW", CStr(f(7))))
    Next key
    snapshot.Add MBW_CheckedText(Array("MBARCHIVEAUDIT1", gRoot, gArchiveAuditStore, gArchiveAuditAccount, _
        "CURSOR", cursor, "", "", "", "", "", "", "", "CURSOR", "", MBC_UtcStamp(MBC_UtcNow()), gRunId))
    ' Cursor and all unresolved proofs are published as one atomic snapshot,
    ' only after the outcome report has been closed and flushed.
    MBS_WriteAtomic gArchiveAuditStatePath, MBW_LinesText(snapshot)
    gArchiveAuditCursor = cursor
    gArchiveAuditPending = gArchiveAuditIssues.Count + gArchiveAuditReferenceReview + gArchiveAuditStateIssues
    gArchiveAuditCommitted = True
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Close
    On Error GoTo 0
    Err.Raise n, "MBA_Run", detail
End Sub

Private Function MBA_PendingCount() As Long
    If Not gArchiveAuditIssues Is Nothing Then MBA_PendingCount = gArchiveAuditIssues.Count
    MBA_PendingCount = MBA_PendingCount + gArchiveAuditReferenceReview + gArchiveAuditStateIssues
End Function

Private Function MBA_HasReview() As Boolean
    MBA_HasReview = (MBA_PendingCount() > 0)
End Function

Private Function MBA_Summary() As String
    If Len(gArchiveAuditStore) = 0 Then Exit Function
    MBA_Summary = vbCrLf & "Independent archived-file audit: checked=" & CStr(gArchiveAuditChecked) & _
        "; passed=" & CStr(gArchiveAuditPassed) & "; failed=" & CStr(gArchiveAuditFailed) & _
        "; physical file reads=" & CStr(gArchiveAuditFilesRead) & vbCrLf & _
        "Known physical proofs=" & CStr(gArchiveAuditTotal) & "; per-run budget=" & CStr(MBA_FILE_BUDGET) & _
        "; unresolved=" & CStr(MBA_PendingCount()) & "; state committed=" & CStr(gArchiveAuditCommitted) & vbCrLf & _
        "Audit verifies recorded bytes/readability without requiring a live Outlook source. " & _
        "MSG-only rebaselines establish current bytes, not lost historical integrity."
End Function


' Disposable native Windows acceptance for metadata snapshot recovery.
' This test uses synthetic metadata and a synthetic sentinel only. It does not
' read real mail, change registry settings, or simulate an actual power loss.
Private Function MBS_NativeRecoverySelfTest(ByRef detail As String) As Boolean
    Const TEST_STORE As String = "MB_NATIVE_RECOVERY_TEST"
    Dim fso As Object, stream As Object, files As Object, chunks As Object
    Dim tempRoot As String, testBase As String, account As String, primary As String
    Dim ownerPath As String, ownerHash As String, ownerToken As String
    Dim probePath As String, probeHash As String, journalHash As String
    Dim sentinelPath As String, sentinelHash As String, text As String, recovery As String
    Dim manifest As String, reason As String, cleanup As String, errorText As String
    Dim newerJournalHash As String, overlayIntent As String, syntheticRow As Variant, stamp As String
    Dim owned As Boolean, initialized As Boolean, rootLocked As Boolean, accountLocked As Boolean
    Dim passed As Boolean, cleaned As Boolean, n As Long, savedBefore As Long, gateError As Long
    detail = ""
    ' Acceptance owns gBusy; a live archive lock or stream is not disposable.
    If MBH_RunLock <> 0 Then detail = "FAIL: a live root lock prevents the disposable test.": Exit Function
    If Not mbmAccountLocks Is Nothing Then
        If mbmAccountLocks.Count > 0 Then detail = "FAIL: live account locks prevent the disposable test.": Exit Function
    End If
    If Not gJournal Is Nothing Then detail = "FAIL: a live journal prevents the disposable test.": Exit Function
    If Not gEvents Is Nothing Then detail = "FAIL: a live event stream prevents the disposable test.": Exit Function
    On Error GoTo Failed
    MB_Reset: initialized = True
    gMode = "SELFTEST_RECOVERY": gPhase = "Testing disposable metadata recovery"
    gClockStart = MB_Clock(): gRunStart = Now: gUntil = gRunStart
    Set fso = CreateObject("Scripting.FileSystemObject"): Set gFSO = fso
    Set gDirs = CreateObject("Scripting.Dictionary"): gDirs.CompareMode = vbTextCompare
    Set gIndex = CreateObject("Scripting.Dictionary"): gIndex.CompareMode = vbBinaryCompare
    tempRoot = MBH_NormalizePath(CStr(fso.GetSpecialFolder(2).Path))
    testBase = MBH_JoinPath(tempRoot, "MBREC_" & fso.GetTempName)
    ' CreateFolder claims a NEW directory. Never reuse or clean an older test.
    fso.CreateFolder testBase: owned = True
    ownerPath = MBH_JoinPath(testBase, "native-owner.txt")
    ownerToken = "Owned disposable recovery test: " & MBC_UtcStamp(MBC_UtcNow()) & "|" & fso.GetTempName
    Set stream = fso.CreateTextFile(ownerPath, False, True)
    stream.Write ownerToken: stream.Close: Set stream = Nothing
    ownerHash = MBH_HashFile(ownerPath)

    gRoot = MBH_JoinPath(testBase, MBP_MANAGED_LEAF): gBaseRoot = gRoot
    account = MBH_JoinPath(testBase, "account")
    MBH_EnsureDirectory gRoot: MBH_EnsureDirectory account
    gStateRoot = MB_StatePath(): primary = gStateRoot
    fso.CreateFolder primary
    ' The run directory stays outside the deliberately deleted primary folder.
    gRunId = "native_recovery"
    gRunDir = MBH_JoinPath(gRoot, "native-run")
    MBH_EnsureDirectory gRunDir
    gJournalPath = MB_StatePath("journal.tsv")
    MBH_AcquireRunLock MBH_JoinPath(gRoot, "backup.lock"): rootLocked = True
    MBM_AcquireAccountLock account: accountLocked = True

    sentinelPath = MBH_JoinPath(account, "protected-sentinel.bin")
    Set stream = fso.CreateTextFile(sentinelPath, False, False)
    stream.Write "Synthetic protected archive bytes. Never metadata."
    stream.Close: Set stream = Nothing
    sentinelHash = MBH_HashFile(sentinelPath)
    Set stream = fso.CreateTextFile(gJournalPath, False, True)
    stream.Write vbNullString: stream.Close: Set stream = Nothing
    MBM_FlushPath gJournalPath
    journalHash = MBH_HashFile(gJournalPath)
    probePath = MB_StatePath("probe.tsv")
    ' Cross the 1 MiB boundary to exercise multi-chunk capture/reassembly.
    text = ChrW$(&HD55C) & ChrW$(&HAE00) & vbTab & "checked Unicode metadata" & vbCrLf & _
           String$(MBR_CHUNK_BYTES \ 2 + 64, "R")
    MBS_WriteAtomic probePath, MBS_Checked(Array("MBNATIVETEST1", "probe", text))
    probeHash = MBH_HashFile(probePath)
    If CDbl(fso.GetFile(probePath).Size) <= MBR_CHUNK_BYTES Then _
        Err.Raise MB_ERR + 970, , "Native probe did not cross the recovery chunk boundary."
    If Not MBR_CheckedStateValid(probePath, reason) Then _
        Err.Raise MB_ERR + 970, , "Synthetic probe framing is invalid: " & reason

    MBR_SaveSnapshot account, TEST_STORE
    If gRecoverySnapshotsSaved <> 1 Then Err.Raise MB_ERR + 970, , "Initial snapshot was not published exactly once."
    recovery = MBR_RecoveryRoot(account, TEST_STORE)
    manifest = MBR_FindSnapshot(recovery, account, TEST_STORE, files, chunks, True)
    If Len(manifest) = 0 Then Err.Raise MB_ERR + 970, , "No complete verified native recovery snapshot exists."
    If files.Count <> 2 Or Not files.Exists("journal.tsv") Or Not files.Exists("probe.tsv") Then _
        Err.Raise MB_ERR + 970, , "Snapshot did not contain exactly the two synthetic state files."
    If Not MBS_NativeOwnedTreeSafe(fso, testBase, tempRoot, ownerHash, reason) Then _
        Err.Raise MB_ERR + 970, , "Cannot prove disposable test ownership before state deletion: " & reason
    If StrComp(primary, MB_StatePath(), vbTextCompare) <> 0 Or Not MB_PathWithin(primary, testBase) Then _
        Err.Raise MB_ERR + 970, , "Disposable primary state escaped its owned test tree."
    ' Only the newly created disposable primary folder is removed. Account
    ' sentinel, recovery generations and the ownership token remain untouched.
    fso.DeleteFolder primary, True
    If fso.FolderExists(primary) Then Err.Raise MB_ERR + 970, , "Disposable primary folder deletion failed."
    If MBH_HashFile(sentinelPath) <> sentinelHash Then Err.Raise MB_ERR + 970, , "Sentinel changed during state deletion."

    MBR_Reset
    MBR_RestoreState account, TEST_STORE
    If Not gRecoveryForceFull Then Err.Raise MB_ERR + 970, , "Restored metadata did not require full reconciliation."
    If gRecoveryRebased Then Err.Raise MB_ERR + 970, , "Snapshot restore was mislabeled as MSG-only rebaselining."
    If gRecoveryRestoredCount <> 2 Then Err.Raise MB_ERR + 970, , "Snapshot restore did not reconstruct two state files."
    If MBH_HashFile(probePath) <> probeHash Or MBH_HashFile(gJournalPath) <> journalHash Then _
        Err.Raise MB_ERR + 970, , "Restored metadata bytes differ from their original hashes."
    If MBH_HashFile(sentinelPath) <> sentinelHash Then Err.Raise MB_ERR + 970, , "Restore modified the protected sentinel."
    If Len(MBR_PendingIntent(recovery, account, TEST_STORE)) <> 0 Then _
        Err.Raise MB_ERR + 970, , "Snapshot restore left an unfinished restore intent."

    ' A new process must still require reconciliation while the durable token
    ' is pending, even when restored state is structurally healthy.
    MBR_Reset
    MBR_RestoreState account, TEST_STORE
    If Not gRecoveryForceFull Or gRecoveryRestoredCount <> 0 Then _
        Err.Raise MB_ERR + 970, , "Pending reconciliation did not survive a healthy restart without replaying snapshot files."
    If MBH_HashFile(probePath) <> probeHash Or MBH_HashFile(gJournalPath) <> journalHash Or _
       MBH_HashFile(sentinelPath) <> sentinelHash Then _
        Err.Raise MB_ERR + 970, , "Healthy pending-reconciliation restart changed protected bytes."
    On Error Resume Next
    Err.Clear: MBR_CompleteReconcile account, TEST_STORE
    gateError = Err.Number: Err.Clear
    On Error GoTo Failed
    If gateError <> MB_ERR + 881 Then Err.Raise MB_ERR + 970, , "Reconciliation accepted an acknowledgement before coverage."
    ' Synthetic acknowledgement only: this test has no real mail or coverage.
    ' The production completion gate is deliberately satisfied for this owned
    ' fixture, then cleared immediately; no actual coverage checkpoint is made.
    gWindowCommitted = True
    MBR_CompleteReconcile account, TEST_STORE
    gWindowCommitted = False
    MBR_Reset
    MBR_RestoreState account, TEST_STORE
    If gRecoveryForceFull Or gRecoveryRestoredCount <> 0 Then _
        Err.Raise MB_ERR + 970, , "Completed synthetic reconciliation did not leave a stable healthy restart."
    If MBH_HashFile(probePath) <> probeHash Or MBH_HashFile(gJournalPath) <> journalHash Or _
       MBH_HashFile(sentinelPath) <> sentinelHash Then _
        Err.Raise MB_ERR + 970, , "Repeated healthy restore changed protected bytes."
    ' Normalize the post-acknowledgement snapshot first; an immediately repeated
    ' snapshot of unchanged metadata must then reuse the complete generation.
    MBR_SaveSnapshot account, TEST_STORE
    savedBefore = gRecoverySnapshotsSaved
    MBR_SaveSnapshot account, TEST_STORE
    If gRecoverySnapshotsSaved <> savedBefore Then _
        Err.Raise MB_ERR + 970, , "Unchanged metadata created another recovery generation."
    ' Ancillary-only loss must not roll a newer valid expected file hash back
    ' to this older snapshot. This row is SYNTHETIC metadata: no MSG is created,
    ' opened or represented as a real mail backup by this disposable scenario.
    manifest = MBR_FindSnapshot(recovery, account, TEST_STORE, files, chunks, True)
    stamp = MB_Stamp(Now)
    syntheticRow = Array("MBS3", "N", MBH_HashText("native revision"), _
        MBH_JoinPath(account, "synthetic-metadata-only.msg"), "37", MBH_HashText("native expected bytes"), _
        MB_SourceKey(TEST_STORE, "native entry"), TEST_STORE, "native entry", "native folder", stamp, stamp, _
        "native proof only", "0", "", MB_NEW_CONTENT_VALIDATION, account)
    Set stream = fso.OpenTextFile(gJournalPath, 8, True, -1)
    MB_WriteCheckedLine stream, MB_EncodeFields(syntheticRow)
    stream.Close: Set stream = Nothing: MBM_FlushPath gJournalPath
    newerJournalHash = MBH_HashFile(gJournalPath)
    If Not MBS_NativeOwnedTreeSafe(fso, testBase, tempRoot, ownerHash, reason) Then _
        Err.Raise MB_ERR + 970, , "Disposable tree ownership changed before ancillary deletion."
    fso.DeleteFile probePath, True
    MBR_Reset
    MBR_RestoreState account, TEST_STORE
    If MBH_HashFile(gJournalPath) <> newerJournalHash Or MBH_HashFile(probePath) <> probeHash Or _
       MBH_HashFile(sentinelPath) <> sentinelHash Or gRecoveryMergedRows <> 1 Then _
        Err.Raise MB_ERR + 970, , "Ancillary-only recovery rolled back newer expected file proof."
    If Not gRecoveryForceFull Then Err.Raise MB_ERR + 970, , "Merged recovery lost reconciliation requirement."

    ' Model a process ending after START and only the old journal being restored.
    ' The next call must replay the immutable ORIGINAL overlay, not recapture the
    ' now-empty partially restored primary and forget its newer valid row.
    overlayIntent = MBR_CaptureOverlay(recovery, account, TEST_STORE, manifest)
    MBR_WriteIntent recovery, account, TEST_STORE, overlayIntent, "START"
    If Not MBS_NativeOwnedTreeSafe(fso, testBase, tempRoot, ownerHash, reason) Then _
        Err.Raise MB_ERR + 970, , "Disposable tree ownership changed before partial-restore fixture."
    Set stream = fso.CreateTextFile(gJournalPath, True, True)
    stream.Write vbNullString: stream.Close: Set stream = Nothing
    MBM_FlushPath gJournalPath
    MBR_Reset
    MBR_RestoreState account, TEST_STORE
    If MBH_HashFile(gJournalPath) <> newerJournalHash Or MBH_HashFile(probePath) <> probeHash Or _
       MBH_HashFile(sentinelPath) <> sentinelHash Or Len(MBR_PendingIntent(recovery, account, TEST_STORE)) <> 0 Then _
        Err.Raise MB_ERR + 970, , "Interrupted restore did not recover the original overlay's exact later proof."
    If Not gRecoveryForceFull Then Err.Raise MB_ERR + 970, , "Resumed overlay restore falsely completed coverage."
    MBR_Reset
    MBR_RestoreState account, TEST_STORE
    If gRecoveryRestoredCount <> 0 Or MBH_HashFile(gJournalPath) <> newerJournalHash Then _
        Err.Raise MB_ERR + 970, , "Merged healthy repeat rewrote prior expected proof."
    If gCancel Or gFatal Or Len(gCleanupIssue) > 0 Then _
        Err.Raise MB_ERR + 970, , "Native recovery verification was cancelled or stopped."
    passed = True
    GoTo Cleanup
Failed:
    n = Err.Number: errorText = Err.Description
Cleanup:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    If accountLocked Then
        Err.Clear: MBM_ReleaseAccountLocks
        If Err.Number <> 0 Then cleanup = cleanup & " Account lock: " & Err.Description
    End If
    If rootLocked Then
        Err.Clear: MBH_ReleaseRunLock
        If Err.Number <> 0 Then cleanup = cleanup & " Root lock: " & Err.Description
    End If
    On Error GoTo 0
    ' On any test failure retain every disposable byte for diagnosis. Cleanup
    ' after success requires both exclusive creation and the unchanged token.
    If passed And owned And Len(cleanup) = 0 And Not gCancel And Not gFatal Then
        If MBS_NativeOwnedTreeSafe(fso, testBase, tempRoot, ownerHash, reason) Then
            On Error Resume Next
            Err.Clear: fso.DeleteFolder testBase, True
            If Err.Number <> 0 Then
                cleanup = cleanup & " Disposable directory cleanup: " & Err.Description
            Else
                cleaned = True
            End If
            On Error GoTo 0
        Else
            cleanup = cleanup & " Disposable ownership check: " & reason
        End If
    End If
    On Error Resume Next
    Err.Clear: MBH_ReleaseCrypto
    If Err.Number <> 0 Then cleanup = cleanup & " Crypto cleanup: " & Err.Description
    If initialized Then
        gStageEnabled = False: gWindowCommitted = False
        Set gDirs = Nothing: Set gIndex = Nothing: Set gFSO = Nothing
        gRoot = "": gBaseRoot = "": gRunDir = "": gRunId = "": gJournalPath = "": gStateRoot = ""
    End If
    On Error GoTo 0
    If passed And cleaned And n = 0 And Len(cleanup) = 0 And Not gCancel And Not gFatal Then
        detail = "PASS: disposable multi-chunk Unicode metadata snapshot, deleted-state exact-hash restore, " & _
                 "sentinel preservation, durable reconcile token, healthy repeat, unchanged-snapshot reuse, " & _
                 "ancillary-loss newer-proof preservation and interrupted overlay resume. " & _
                 "Synthetic metadata only; no real mail, registry changes or actual power-loss test."
        MBS_NativeRecoverySelfTest = True
    Else
        detail = "FAIL: disposable metadata recovery " & CStr(n) & " - " & errorText & cleanup
        If owned And Not cleaned Then detail = detail & "; disposable evidence retained at " & testBase
    End If
End Function

Private Function MBS_NativeOwnedTreeSafe(ByVal fso As Object, ByVal testBase As String, _
                                        ByVal tempRoot As String, ByVal ownerHash As String, _
                                        ByRef reason As String) As Boolean
    Dim pending As New Collection, current As String, folder As Object, item As Object, ownerPath As String
    On Error GoTo Rejected
    reason = ""
    If Len(testBase) = 0 Or Len(ownerHash) = 0 Then Err.Raise MB_ERR + 971, , "Missing disposable ownership proof."
    If StrComp(CStr(fso.GetParentFolderName(testBase)), tempRoot, vbTextCompare) <> 0 Or _
       Left$(CStr(fso.GetFileName(testBase)), 6) <> "MBREC_" Then _
        Err.Raise MB_ERR + 971, , "Disposable directory is not a direct owned Temp child."
    MBY_NoReparse testBase, True
    ownerPath = MBH_JoinPath(testBase, "native-owner.txt")
    MBY_NoReparse ownerPath, True
    If MBH_HashFile(ownerPath) <> ownerHash Then Err.Raise MB_ERR + 971, , "Disposable ownership token changed."
    pending.Add testBase
    Do While pending.Count > 0
        current = CStr(pending(pending.Count)): pending.Remove pending.Count
        Set folder = fso.GetFolder(current)
        If (CLng(folder.Attributes) And 1024) <> 0 Then Err.Raise MB_ERR + 971, , "Disposable tree contains a reparse folder."
        For Each item In folder.Files
            If (CLng(item.Attributes) And 1024) <> 0 Then Err.Raise MB_ERR + 971, , "Disposable tree contains a reparse file."
        Next item
        For Each item In folder.SubFolders
            If (CLng(item.Attributes) And 1024) <> 0 Then Err.Raise MB_ERR + 971, , "Disposable tree contains a reparse folder."
            pending.Add CStr(item.Path)
        Next item
    Loop
    MBS_NativeOwnedTreeSafe = True
    Exit Function
Rejected:
    reason = CStr(Err.Number) & ": " & Err.Description
End Function


' ================= One default Outlook store per managed root ===============
' Central snapshots cover the whole state directory. A second store must not
' share that directory and later restore an older view of another store's data.
' The root-level marker and HKCU copy survive deletion of the central data.
' If every historical binding is gone, this is a NEW observation of the chosen
' store, not proof that the old root was historically exclusive to that store.

Private Function MBS_StoreBindingValue(ByVal line As String, Optional ByVal rootPath As String = "") As String
    Dim fields As Variant
    If Len(rootPath) = 0 Then rootPath = gRoot
    fields = MB_ReadCheckedLine(line, 4)
    If CStr(fields(0)) <> "MBSTORE1" Or CStr(fields(3)) <> "SINGLE_DEFAULT_STORE" Then _
        Err.Raise MB_ERR + 930, , "Unknown managed-root store binding."
    If StrComp(CStr(fields(1)), MBH_NormalizePath(rootPath), vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 931, , "Store binding belongs to another managed root."
    If Len(CStr(fields(2))) = 0 Then Err.Raise MB_ERR + 932, , "Store binding has no StoreID."
    MBS_StoreBindingValue = CStr(fields(2))
End Function

Private Sub MBS_CheckCandidateStore(ByVal rootPath As String, ByVal storeId As String, ByVal fso As Object)
    ' Read-only gate before remembering a folder, with no dependence on gRoot.
    ' A foreign saved root should lead to the folder chooser, not trap every run.
    Dim path As String, saved As String, found As String, bound As Boolean
    path = MBH_JoinPath(rootPath, "operating-store.tsv")
    If fso.FileExists(path) Then
        found = MBS_StoreBindingValue(MBP_ReadOneLine(path, fso), rootPath)
        If StrComp(found, storeId, vbTextCompare) <> 0 Then _
            Err.Raise MB_ERR + 934, , "[FOREIGN_STORE] This managed root is bound to another Outlook store: " & rootPath
        bound = True
    End If
    saved = MBP_GetSettingText("OperatingStore_" & MBH_HashText(LCase$(MBH_NormalizePath(rootPath))), "")
    If Len(saved) > 0 Then
        found = MBS_StoreBindingValue(saved, rootPath)
        If StrComp(found, storeId, vbTextCompare) <> 0 Then _
            Err.Raise MB_ERR + 934, , "[FOREIGN_STORE] This managed root is bound to another Outlook store: " & rootPath
        bound = True
    End If
    If Not bound Then MBS_PreflightCandidateEvidence rootPath, storeId, fso
End Sub

Private Function MBS_ReadStoreBinding(ByVal path As String) As String
    Dim stream As Object, line As String, n As Long, detail As String
    On Error GoTo Failed
    MBY_NoReparse path, True
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then _
        Err.Raise MB_ERR + 933, , "Managed-root store binding is truncated."
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    If stream.AtEndOfStream Then Err.Raise MB_ERR + 933, , "Managed-root store binding is empty."
    line = stream.ReadLine
    If Not stream.AtEndOfStream Then Err.Raise MB_ERR + 933, , "Managed-root store binding has extra data."
    stream.Close: Set stream = Nothing
    MBS_ReadStoreBinding = MBS_StoreBindingValue(line)
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBS_ReadStoreBinding", detail
End Function

Private Sub MBS_RequireStore(ByVal foundStore As String, ByVal storeId As String, ByVal source As String)
    If Len(foundStore) = 0 Then Exit Sub
    If StrComp(foundStore, storeId, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 934, , "[FOREIGN_STORE] This managed root contains a binding to another Outlook store: " & source & _
                  ". Use a separate managed backup root for each default store; existing files were preserved."
End Sub

Private Function MBS_TryStoreRow(ByVal line As String, ByVal fieldCount As Long, ByRef fields As Variant) As Boolean
    Dim n As Long, detail As String
    On Error GoTo Invalid
    fields = MB_ReadCheckedLine(line, fieldCount)
    MBS_TryStoreRow = True
    Exit Function
Invalid:
    n = Err.Number: detail = Err.Description
    ' A damaged row is not ownership evidence. The normal recovery loader will
    ' preserve/repair it; unrelated I/O or cryptographic failures must propagate.
    If n = MB_ERR + 15 Or n = MB_ERR + 16 Or n = MB_ERR + 17 Then Exit Function
    Err.Raise n, "MBS_TryStoreRow", detail
End Function

Private Sub MBS_CheckStoreRow(ByVal line As String, ByVal kind As String, ByVal storeId As String, _
                              ByVal source As String, Optional ByVal accountPath As String = "", _
                              Optional ByVal rootPath As String = "", Optional ByVal strictAsOwned As Boolean = False)
    Dim fields As Variant, fieldCount As Long, foundStore As String
    If Len(rootPath) = 0 Then rootPath = gRoot
    Select Case kind
        Case "JOURNAL"
            If Left$(line, 5) = "MBS2" & vbTab Then
                fieldCount = 16
            ElseIf Left$(line, 5) = "MBS3" & vbTab Or Left$(line, 5) = "MBS4" & vbTab Then
                fieldCount = 17
            Else
                GoTo Unproven
            End If
            If Not MBS_TryStoreRow(line, fieldCount, fields) Then GoTo Unproven
            foundStore = CStr(fields(7))
        Case "ACCOUNT"
            If Not MBS_TryStoreRow(line, 4, fields) Then GoTo Unproven
            If CStr(fields(0)) <> "MBAC3" Then GoTo Unproven
            If StrComp(CStr(fields(3)), rootPath, vbTextCompare) <> 0 Then GoTo Unproven
            If Len(accountPath) > 0 Then
                If StrComp(CStr(fields(2)), accountPath, vbTextCompare) <> 0 Then GoTo Unproven
            End If
            foundStore = CStr(fields(1))
        Case "RECOVERY"
            If Not MBS_TryStoreRow(line, 5, fields) Then GoTo Unproven
            If CStr(fields(0)) <> "MBRECOVERY1" Or CStr(fields(4)) <> "1" Then GoTo Unproven
            If StrComp(CStr(fields(1)), rootPath, vbTextCompare) <> 0 Then GoTo Unproven
            If StrComp(CStr(fields(2)), accountPath, vbTextCompare) <> 0 Then GoTo Unproven
            foundStore = CStr(fields(3))
        Case Else
            Err.Raise MB_ERR + 935, , "Unknown store-evidence kind."
    End Select
    If Len(foundStore) = 0 Then GoTo Unproven
    MBS_RequireStore foundStore, storeId, source
    Exit Sub
Unproven:
    If strictAsOwned Then Err.Raise MB_ERR + 937, , _
        "The existing managed root has unreadable or conflicting store evidence; it was preserved for review: " & source
End Sub

Private Sub MBS_PreflightStoreFile(ByVal path As String, ByVal kind As String, ByVal storeId As String, _
                                   ByVal rootPath As String, ByVal fso As Object, ByVal strictAsOwned As Boolean, _
                                   Optional ByVal accountPath As String = "")
    Dim stream As Object, line As String, n As Long, detail As String
    On Error GoTo Failed
    If Not fso.FileExists(path) Then Exit Sub
    MBP_NoReparse path, fso
    If CDbl(fso.GetFile(path).Size) Mod 2 <> 0 Then
        If strictAsOwned Then Err.Raise MB_ERR + 937, , "Truncated managed-root store evidence needs review: " & path
        Exit Sub
    End If
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then MBS_CheckStoreRow line, kind, storeId, path, accountPath, rootPath, strictAsOwned
    Loop
    stream.Close: Set stream = Nothing
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBS_PreflightStoreFile", detail
End Sub

Private Sub MBS_PreflightCandidateAccount(ByVal accountPath As String, ByVal storeId As String, _
                                          ByVal rootPath As String, ByVal fso As Object)
    ' Only an exact checked sibling marker is evidence for this root. Unrelated
    ' malformed sibling files are not borrowed as authority to bind any store.
    Dim leaf As Variant, recovery As String
    MBS_PreflightStoreFile MBH_JoinPath(accountPath, MBY_ACCOUNT_MARKER), "ACCOUNT", storeId, rootPath, fso, False, accountPath
    For Each leaf In Array(MBR_FOLDER, MBR_FOLDER_LEGACY)
        recovery = MBH_JoinPath(accountPath, CStr(leaf))
        MBS_PreflightStoreFile MBH_JoinPath(recovery, MBR_OWNER_FILE), _
                           "RECOVERY", storeId, rootPath, fso, False, accountPath
    Next leaf
End Sub

Private Sub MBS_PreflightCandidateEvidence(ByVal rootPath As String, ByVal storeId As String, ByVal fso As Object)
    Dim stateLeaf As Variant, state As String, parent As String, folder As Object, account As String
    ' The first upgrade of an old root has no MBSTORE1 marker. Inspect owned
    ' state before the chooser returns, so a different default store can choose
    ' its separate root through the same single public macro.
    For Each stateLeaf In Array(MBS_DATA_LEAF, MBS_DATA_LEAF_LEGACY, "_state")
        state = MBH_JoinPath(rootPath, CStr(stateLeaf))
        MBS_PreflightStoreFile MBH_JoinPath(state, "journal.tsv"), "JOURNAL", storeId, rootPath, fso, True
        MBS_PreflightStoreFile MBH_JoinPath(state, MBY_MAP_FILE), "ACCOUNT", storeId, rootPath, fso, True
    Next stateLeaf
    account = MBP_OwnerAccountHint(storeId, rootPath)
    If Len(account) > 0 Then MBS_PreflightCandidateAccount account, storeId, rootPath, fso
    parent = CStr(fso.GetParentFolderName(rootPath))
    If Len(parent) = 0 Or Not fso.FolderExists(parent) Then Exit Sub
    MBP_NoReparse parent, fso
    MBS_PreflightCandidateAccount parent, storeId, rootPath, fso
    For Each folder In fso.GetFolder(parent).SubFolders
        If (CLng(folder.Attributes) And 1024) = 0 Then _
            MBS_PreflightCandidateAccount CStr(folder.Path), storeId, rootPath, fso
    Next folder
End Sub

Private Sub MBS_CheckStoreFile(ByVal path As String, ByVal kind As String, ByVal storeId As String, _
                               Optional ByVal accountPath As String = "")
    Dim stream As Object, line As String, n As Long, detail As String
    On Error GoTo Failed
    If Not gFSO.FileExists(path) Then Exit Sub
    MBY_NoReparse path, True
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then MBS_CheckStoreRow line, kind, storeId, path, accountPath
    Loop
    stream.Close: Set stream = Nothing
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "MBS_CheckStoreFile", detail
End Sub

Private Sub MBS_CheckAccountStoreEvidence(ByVal accountPath As String, ByVal storeId As String)
    Dim leaf As Variant, recovery As String
    MBS_CheckStoreFile MBH_JoinPath(accountPath, MBY_ACCOUNT_MARKER), "ACCOUNT", storeId, accountPath
    For Each leaf In Array(MBR_FOLDER, MBR_FOLDER_LEGACY)
        recovery = MBH_JoinPath(accountPath, CStr(leaf))
        MBS_CheckStoreFile MBH_JoinPath(recovery, MBR_OWNER_FILE), "RECOVERY", storeId, accountPath
    Next leaf
End Sub

Private Sub MBS_CheckInitialStoreEvidence(ByVal storeId As String)
    Dim oldState As String, parent As String, folder As Object
    MBS_CheckStoreFile MB_StatePath("journal.tsv"), "JOURNAL", storeId
    MBS_CheckStoreFile MB_StatePath(MBY_MAP_FILE), "ACCOUNT", storeId
    oldState = MBH_JoinPath(gRoot, "_state")
    MBS_CheckStoreFile MBH_JoinPath(oldState, "journal.tsv"), "JOURNAL", storeId
    MBS_CheckStoreFile MBH_JoinPath(oldState, MBY_MAP_FILE), "ACCOUNT", storeId
    parent = CStr(gFSO.GetParentFolderName(gRoot))
    If Len(parent) = 0 Then Exit Sub
    MBY_NoReparse parent, True
    ' The archive may be the parent itself or a sibling of the state folder.
    ' Custom roots elsewhere remain guarded when their exact owner is opened.
    MBS_CheckAccountStoreEvidence parent, storeId
    For Each folder In gFSO.GetFolder(parent).SubFolders
        If (CLng(folder.Attributes) And 1024) = 0 Then _
            MBS_CheckAccountStoreEvidence CStr(folder.Path), storeId
    Next folder
End Sub

Private Sub MBS_BindStore(ByVal storeId As String)
    Dim path As String, settingKey As String, saved As String, diskStore As String
    Dim registryStore As String, payload As String, checkedLine As String
    If Len(storeId) = 0 Or Len(gRoot) = 0 Then Err.Raise MB_ERR + 936, , "Root store binding needs a root and StoreID."
    MBY_NoReparse gRoot, True
    path = MBH_JoinPath(gRoot, "operating-store.tsv")
    settingKey = "OperatingStore_" & MBH_HashText(LCase$(MBH_NormalizePath(gRoot)))
    If gFSO.FileExists(path) Then diskStore = MBS_ReadStoreBinding(path)
    saved = MBP_GetSettingText(settingKey, "")
    If Len(saved) > 0 Then registryStore = MBS_StoreBindingValue(saved)
    MBS_RequireStore diskStore, storeId, path
    MBS_RequireStore registryStore, storeId, "HKCU " & settingKey
    ' A surviving exact binding is authority: no full journal scan on every run.
    If Len(diskStore) = 0 And Len(registryStore) = 0 Then MBS_CheckInitialStoreEvidence storeId
    payload = MB_EncodeFields(Array("MBSTORE1", MBH_NormalizePath(gRoot), LCase$(storeId), "SINGLE_DEFAULT_STORE"))
    checkedLine = payload & vbTab & MBH_HashText(payload) & vbTab & "END"
    ' Each durable copy can repair a DELETED peer. A malformed or foreign peer
    ' stops above and is retained for review; it never authorizes a new store.
    If Len(diskStore) = 0 Then MBS_WriteAtomic path, checkedLine & vbCrLf
    If Len(registryStore) = 0 Then MBP_SaveSettingText settingKey, checkedLine
End Sub

Private Sub MBS_CheckRestoredStore(ByVal storeId As String)
    ' A structurally healthy no-op restore does not need another O(N) scan.
    ' Force-repair can replay identical bytes, so its explicit outcome also
    ' triggers validation even if no physical file was reconstructed.
    If gRecoveryRestoredCount = 0 And Not gRecoveryRebased And _
       InStr(1, gRecoveryNote, "RECOVERED_SNAPSHOT:", vbBinaryCompare) <> 1 Then Exit Sub
    MBS_CheckStoreFile MB_StatePath("journal.tsv"), "JOURNAL", storeId
    MBS_CheckStoreFile MB_StatePath(MBY_MAP_FILE), "ACCOUNT", storeId
End Sub


' Repair only a damaged stage record whose complete original payload is also
' present as an independently checked record in the other ledger. Unknown
' fragments are never dropped, inferred or converted into ownership proof.
Private Function MBSR_Payload(ByVal raw As Variant) As String
    Dim values(0 To 9) As String, i As Long
    For i = 0 To 9
        values(i) = CStr(raw(i))
    Next i
    MBSR_Payload = Join(values, vbTab)
End Function

Private Function MBSR_Identity(ByVal f As Variant) As String
    Dim sequence As Double, stamp As Date, stage As String, final As String
    If CStr(f(0)) <> "MBSTAGE1" Or StrComp(CStr(f(1)), gRoot, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 960, , "Stage repair payload has a different archive binding."
    sequence = MB_ByteCount(CStr(f(8)))
    If sequence < 1 Or sequence > 2000000000# Or sequence <> Fix(sequence) Then _
        Err.Raise MB_ERR + 961, , "Stage repair payload has an invalid sequence."
    stamp = MBC_ParseUtcStamp(CStr(f(9)))
    If Not MBC_ValidDate(stamp) Then Err.Raise MB_ERR + 962, , "Stage repair payload lacks a complete UTC timestamp."
    stage = MBH_NormalizePath(CStr(f(2)))
    final = MBH_NormalizePath(CStr(f(3)))
    If Len(stage) = 0 Or Len(final) = 0 Then Err.Raise MB_ERR + 963, , "Stage repair payload lacks a path."
    MBSR_Identity = LCase$(stage) & "|" & Format$(sequence, "0")
End Function

Private Function MBSR_CheckedRecord(ByVal line As String, ByRef payload As String, ByRef identity As String) As Boolean
    Dim f As Variant, raw As Variant
    On Error GoTo Invalid
    f = MB_ReadCheckedLine(line, 10)
    identity = MBSR_Identity(f)
    raw = Split(line, vbTab)
    payload = MBSR_Payload(raw)
    MBSR_CheckedRecord = True
Invalid:
End Function

Private Function MBSR_UnprovenPayload(ByVal line As String, ByRef payload As String, ByRef identity As String) As Boolean
    Dim raw As Variant, f As Variant
    On Error GoTo Invalid
    If Len(line) > 16384 Then Exit Function
    raw = Split(line, vbTab)
    ' Only the normal checksum/END framing may be damaged or truncated. Extra
    ' fields could contain another unknown record and are never discarded.
    If UBound(raw) < 9 Or UBound(raw) > 11 Then Exit Function
    If UBound(raw) >= 10 Then
        If Len(CStr(raw(10))) > 64 Then Exit Function
    End If
    If UBound(raw) = 11 Then
        If Len(CStr(raw(11))) > 3 Then Exit Function
    End If
    payload = MBSR_Payload(raw)
    f = MBM_Unpack(payload, 10)
    identity = MBSR_Identity(f)
    ' This is only a candidate, NOT a validated record. The other file must
    ' supply an exact payload with its original valid checksum and END marker.
    MBSR_UnprovenPayload = True
Invalid:
End Function

Private Function MBSR_RepairOne(ByVal target As String, ByVal counterpart As String, _
                                ByVal recoveryRoot As String) As Long
    Const MAX_BAD_ROWS As Long = 64
    Dim reader As Object, writer As Object, bad As Object, wanted As Object, proven As Object
    Dim line As String, payload As String, identity As String, rowKey As Variant
    Dim rowNumber As Long, targetRows As Long, rewritten As Long, expected As String
    Dim sourceHash As String, newHash As String, preservedRoot As String, preserved As String, stage As String
    Dim n As Long, detail As String
#If VBA7 Then
    Dim targetGuard As LongPtr, counterpartGuard As LongPtr, stageGuard As LongPtr
#Else
    Dim targetGuard As Long, counterpartGuard As Long, stageGuard As Long
#End If
    On Error GoTo Failed
    If Not gFSO.FileExists(target) Or Not gFSO.FileExists(counterpart) Then Exit Function
    MBY_NoReparse target, True: MBY_NoReparse counterpart, True
    targetGuard = MBH_GuardFile(target)
    counterpartGuard = MBH_GuardFile(counterpart)
    Set bad = CreateObject("Scripting.Dictionary")
    Set wanted = CreateObject("Scripting.Dictionary")
    Set proven = CreateObject("Scripting.Dictionary")
    bad.CompareMode = vbBinaryCompare: wanted.CompareMode = vbTextCompare: proven.CompareMode = vbBinaryCompare
    Set reader = gFSO.OpenTextFile(target, 1, False, -1)
    Do Until reader.AtEndOfStream
        line = reader.ReadLine: rowNumber = rowNumber + 1
        If Len(line) > 0 Then
            If Not MBSR_CheckedRecord(line, payload, identity) Then
                If bad.Count >= MAX_BAD_ROWS Then GoTo Unproven
                If Not MBSR_UnprovenPayload(line, payload, identity) Then GoTo Unproven
                If wanted.Exists(identity) Then
                    If CStr(wanted(identity)) <> payload Then GoTo Unproven
                End If
                wanted(identity) = payload
                bad(CStr(rowNumber)) = payload
            End If
        End If
    Loop
    targetRows = rowNumber
    reader.Close: Set reader = Nothing
    If bad.Count = 0 Then GoTo Unproven
    Set reader = gFSO.OpenTextFile(counterpart, 1, False, -1)
    Do Until reader.AtEndOfStream
        line = reader.ReadLine
        If Len(line) > 0 Then
            If MBSR_CheckedRecord(line, payload, identity) Then
                If wanted.Exists(identity) Then
                    If CStr(wanted(identity)) <> payload Then GoTo Unproven
                    proven(payload) = line
                End If
            End If
        End If
    Loop
    reader.Close: Set reader = Nothing
    For Each rowKey In bad.Keys
        If Not proven.Exists(CStr(bad(rowKey))) Then GoTo Unproven
    Next rowKey
    sourceHash = MBH_HashFile(target)
    stage = MBH_JoinPath(CStr(gFSO.GetParentFolderName(target)), "_mb_stage_repair_" & gFSO.GetTempName)
    Set writer = gFSO.CreateTextFile(stage, False, True)
    Set reader = gFSO.OpenTextFile(target, 1, False, -1)
    rowNumber = 0
    Do Until reader.AtEndOfStream
        line = reader.ReadLine: rowNumber = rowNumber + 1
        If bad.Exists(CStr(rowNumber)) Then
            expected = CStr(bad(CStr(rowNumber)))
            If Not MBSR_UnprovenPayload(line, payload, identity) Then _
                Err.Raise MB_ERR + 964, , "Stage repair source changed during guarded rewrite."
            If payload <> expected Then Err.Raise MB_ERR + 964, , "Stage repair payload changed during rewrite."
            writer.WriteLine CStr(proven(expected))
            rewritten = rewritten + 1
        Else
            ' Preserve every valid/blank historical row, including other accounts.
            writer.WriteLine line
        End If
    Loop
    reader.Close: Set reader = Nothing
    writer.Close: Set writer = Nothing
    If rowNumber <> targetRows Or rewritten <> bad.Count Then _
        Err.Raise MB_ERR + 965, , "Stage repair row reconciliation failed."
    MBM_FlushPath stage
    stageGuard = MBH_GuardFile(stage)
    newHash = MBH_HashFile(stage)
    preservedRoot = MBH_JoinPath(recoveryRoot, "preserved")
    MBY_NoReparse preservedRoot, False
    MBH_EnsureDirectory preservedRoot
    preserved = MBH_JoinPath(preservedRoot, "stage_repair_" & gRunId & "_" & _
        gFSO.GetBaseName(gFSO.GetTempName) & ".raw")
    gFSO.CopyFile target, preserved, False
    MBM_FlushPath preserved
    If MBH_HashFile(preserved) <> sourceHash Or MBH_HashFile(target) <> sourceHash Then _
        Err.Raise MB_ERR + 966, , "Original stage ledger was not preserved exactly."
    ' No replacement occurs until the original is independently durable and
    ' every changed row has an exact independently verified counterpart.
    MBH_CloseGuard stageGuard
    MBH_CloseGuard targetGuard
    MBH_CloseGuard counterpartGuard
    If MBM_MoveFileExW(StrPtr(stage), StrPtr(target), 9) = 0 Then _
        Err.Raise MB_ERR + 967, , "Atomic stage-ledger repair failed. Win32=" & CStr(Err.LastDllError)
    MBM_FlushPath target
    If MBH_HashFile(target) <> newHash Then Err.Raise MB_ERR + 968, , "Repaired stage ledger changed after publication."
    MB_LogEvent "WARNING", "STAGE_LEDGER_REPAIRED_FROM_COUNTERPART", "", target, 0, _
        "Repaired checked rows=" & CStr(rewritten) & "; exact original preserved at " & preserved
    MBSR_RepairOne = rewritten
    Exit Function
Unproven:
    If Not reader Is Nothing Then reader.Close: Set reader = Nothing
    MBH_CloseGuard targetGuard
    MBH_CloseGuard counterpartGuard
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not reader Is Nothing Then reader.Close
    If Not writer Is Nothing Then writer.Close
    MBH_CloseGuard stageGuard
    MBH_CloseGuard targetGuard
    MBH_CloseGuard counterpartGuard
    ' Keep any owned intermediate for diagnosis; never remove unknown bytes.
    On Error GoTo 0
    Err.Raise n, "MBSR_RepairOne", detail
End Function

Private Function MBS_RepairStageLedgers(ByVal primary As String, ByVal mirror As String) As Long
    Dim recoveryRoot As String
    primary = MBH_NormalizePath(primary): mirror = MBH_NormalizePath(mirror)
    If StrComp(primary, MB_StatePath("stages.tsv"), vbTextCompare) <> 0 Or _
       StrComp(mirror, gStageMirror, vbTextCompare) <> 0 Or _
       StrComp(CStr(gFSO.GetFileName(mirror)), "stage-intents.tsv", vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 969, , "Stage repair destinations do not match their exact owned bindings."
    recoveryRoot = CStr(gFSO.GetParentFolderName(mirror))
    If Not MBR_IsRecoveryDirectory(recoveryRoot) Then _
        Err.Raise MB_ERR + 970, , "Stage repair requires a verified independent recovery directory."
    MBS_RepairStageLedgers = MBSR_RepairOne(primary, mirror, recoveryRoot)
    MBS_RepairStageLedgers = MBS_RepairStageLedgers + MBSR_RepairOne(mirror, primary, recoveryRoot)
End Function


' ==================== Verified same-source version changes ==================
' Source identity (StoreID + EntryID) does not prove content equality. It chooses
' one previously accepted file proof in O(1); that proof's immutable expected
' SHA/size and bounded content comparison decide reuse, new version, or REVIEW.
' The ordinary restored-identity search may still find an older equal version
' after CHANGED, so editing and later reverting content need not duplicate MSGs.
Private Sub MBV_Reset()
    Set gRevisionHeads = Nothing
    gRevisionCompared = 0: gRevisionReused = 0: gRevisionDifferent = 0: gRevisionReview = 0
End Sub

Private Sub MBV_CacheRecord(ByVal f As Variant)
    Dim kind As String
    kind = CStr(f(1))
    If kind <> "M" And kind <> "N" And kind <> "L" And kind <> "R" Then Exit Sub
    If gRevisionHeads Is Nothing Then
        Set gRevisionHeads = CreateObject("Scripting.Dictionary")
        gRevisionHeads.CompareMode = vbBinaryCompare
    End If
    ' Called only after MBX_CacheRecord accepts the complete semantic record.
    ' Last accepted journal order reflects the last observed source revision;
    ' provider timestamps may move backwards and must not choose this head.
    gRevisionHeads(CStr(f(6))) = MBX_IndexKey(kind, CStr(f(2)))
End Sub

Private Function MBV_RequiresCurrentPolicy(ByVal sourceKey As String) As Boolean
    Dim key As String, fields As Variant
    If gRevisionHeads Is Nothing Then Exit Function
    If Not gRevisionHeads.Exists(sourceKey) Then Exit Function
    key = CStr(gRevisionHeads(sourceKey))
    If Not gIndex.Exists(key) Then MBV_RequiresCurrentPolicy = True: Exit Function
    fields = MB_IndexFields(key)
    If Left$(key, 2) = "L|" Then
        MBV_RequiresCurrentPolicy = (CStr(fields(3)) <> MB_LEGACY_VALIDATION)
    Else
        MBV_RequiresCurrentPolicy = (CStr(fields(3)) <> MB_NEW_CONTENT_VALIDATION)
    End If
    ' Evaluate only the current inventory's source. Historical orphan records
    ' do not force all live mail into perpetual full scans. A retained current
    ' BASIC/limited or old-policy proof cannot hide behind a newer checkpoint.
End Function

Private Function MBV_TryKnownRevision(ByVal mail As Outlook.MailItem, ByVal q As Variant, _
                                      ByVal revision As String, ByRef reason As String) As Long
    Dim key As String, fields As Variant, meta As Variant, current As Variant
    Dim path As String, root As String, digest As String, bytes As Double, result As Long
    Dim expectedPath As String, expectedRoot As String, expectedHash As String, expectedBytes As Double
    Dim n As Long, detail As String, targetKey As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    reason = ""
    If gRevisionHeads Is Nothing Then Exit Function
    If Not gRevisionHeads.Exists(CStr(q(7))) Then Exit Function
    key = CStr(gRevisionHeads(CStr(q(7))))
    If Not gIndex.Exists(key) Then Err.Raise MB_ERR + 1100, , "Known source head has no accepted file proof."
    meta = MBX_Meta(key)
    If CStr(meta(2)) <> CStr(q(7)) Or CStr(meta(0)) <> MBX_StoreToken(CStr(q(1))) Then _
        Err.Raise MB_ERR + 1101, , "Known source head belongs to another source or store."
    fields = MB_IndexFields(key)
    expectedHash = CStr(fields(2)): expectedBytes = MB_ByteCount(CStr(fields(1)))
    expectedPath = MBX_IndexedPath(key, fields, True)
    expectedRoot = MBM_ResolvedRoot(MBX_IndexedOriginal(key, fields), expectedHash, MBX_IndexedRoot(key))
    gRevisionCompared = gRevisionCompared + 1
    guard = MBH_GuardFile(expectedPath)
    result = MBI_CheckCandidate(mail, key, path, root, digest, bytes, reason)
    If result = 3 Then MBV_TryKnownRevision = MBV_CANCELLED: GoTo Finished
    If result = 2 Then GoTo NeedsReview
    If result <> 0 And result <> 1 Then Err.Raise MB_ERR + 1102, , "Unknown same-source comparison result."
    ' The comparator writes ByRef outputs. They never replace the saved expected
    ' proof, including when it reports explicitly different current content.
    If StrComp(path, expectedPath, vbTextCompare) <> 0 Or _
       StrComp(root, expectedRoot, vbTextCompare) <> 0 Or digest <> expectedHash Or bytes <> expectedBytes Then _
        Err.Raise MB_ERR + 1103, , "Known source file reference changed during comparison."
    If Not MB_SourceStable(mail, q, revision) Then
        reason = "The current source changed during known-revision comparison.": GoTo NeedsReview
    End If
    targetKey = "M|" & revision
    If result = 0 Then
        ' A provider hiding a content change behind the SAME current M revision
        ' must not cause that immutable key to be overwritten with new bytes.
        If gIndex.Exists(targetKey) Then
            reason = "Current content differs but its provider revision already has a file proof; preserve both facts for review."
            GoTo NeedsReview
        End If
        MBH_CloseGuard guard
        gRevisionDifferent = gRevisionDifferent + 1
        MB_LogEvent "INFO", "KNOWN_SOURCE_NEW_VERSION", q(7), expectedPath, 0, _
            "Prior expected bytes still verify; current bounded content is explicitly different. Old file remains preserved."
        MBV_TryKnownRevision = MBV_CHANGED
        Exit Function
    End If
    If gIndex.Exists(targetKey) Then
        current = MB_IndexFields(targetKey)
        If CStr(current(2)) <> expectedHash Or MB_ByteCount(CStr(current(1))) <> expectedBytes Or _
           StrComp(MBX_IndexedPath(targetKey, current, True), expectedPath, vbTextCompare) <> 0 Then
            reason = "Current revision already points to a different file proof; automatic replacement is blocked."
            GoTo NeedsReview
        End If
        ' An already-current key needs no self-referencing R journal row.
    Else
        MB_RecordCommit "R", revision, expectedPath, expectedBytes, expectedHash, q, mail.Subject, _
                        0, key, MB_NEW_CONTENT_VALIDATION, expectedRoot
        gIdentityAliases = gIdentityAliases + 1
    End If
    MBH_CloseGuard guard
    gRevisionReused = gRevisionReused + 1
    MB_LogEvent "INFO", "KNOWN_SOURCE_REUSED_NO_SAVEAS", q(7), expectedPath, 0, _
        "Current bounded content matches the guarded prior source proof; metadata revision linked without MSG export."
    MBV_TryKnownRevision = MBV_REUSED
    Exit Function
NeedsReview:
    If Len(reason) = 0 Then reason = "Known source content or prior file proof is not fully verifiable."
    gRevisionReview = gRevisionReview + 1
    If gCancel Then MBV_TryKnownRevision = MBV_CANCELLED Else MBV_TryKnownRevision = MBV_REVIEW
Finished:
    MBH_CloseGuard guard
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    If gFatal Then Err.Raise n, "MBV_TryKnownRevision", detail
    reason = "Known source revision requires review: " & detail
    gRevisionReview = gRevisionReview + 1
    If gCancel Then MBV_TryKnownRevision = MBV_CANCELLED Else MBV_TryKnownRevision = MBV_REVIEW
End Function

Private Function MBV_Summary() As String
    MBV_Summary = vbCrLf & "Known-source version comparisons: " & CStr(gRevisionCompared) & _
        "; reused without SaveAs: " & CStr(gRevisionReused) & _
        "; explicitly different version candidates: " & CStr(gRevisionDifferent) & _
        "; review: " & CStr(gRevisionReview)
End Function


' ================= Actual normal-run performance observations ===============
' Observe MB_ProcessOne only. Never issue extra SaveAs, alter the backup decision,
' extrapolate a synthetic acceptance test, or treat REUSE/REVIEW as NEW work.
' The histogram retains EVERY timed observation using fixed memory. Reported
' p50/p95 are containing intervals, not exact percentiles or a confidence band.
Private Sub MBF_Reset()
    Erase gMetricObserved
    Erase gMetricTimed
    Erase gMetricSeconds
    Erase gMetricMinimum
    Erase gMetricMaximum
    Erase gMetricZeros
    Erase gMetricHistogram
    gMetricFault = False: gMetricFaultText = ""
End Sub

Private Function MBF_GroupFor(ByVal outcome As Long, ByVal contentLimited As Boolean) As Long
    Select Case outcome
        Case 1
            If contentLimited Then MBF_GroupFor = 9 Else MBF_GroupFor = 1
        Case 2
            If contentLimited Then MBF_GroupFor = 10 Else MBF_GroupFor = 2
        Case 3 To 8
            MBF_GroupFor = outcome
        Case 0
            ' The queue error handler uses 0 only for an attempt still in flight.
            MBF_GroupFor = 11
        Case Else
            MBF_GroupFor = 12
    End Select
End Function

Private Function MBF_GroupLabel(ByVal group As Long) As String
    Select Case group
        Case 1: MBF_GroupLabel = "NEW_SUCCESS"
        Case 2: MBF_GroupLabel = "REUSE_SUCCESS"
        Case 3: MBF_GroupLabel = "FAILED"
        Case 4: MBF_GroupLabel = "SOURCE_CHANGED"
        Case 5: MBF_GroupLabel = "CANCELLED_ATTEMPT"
        Case 6: MBF_GroupLabel = "POLICY_SKIP"
        Case 7: MBF_GroupLabel = "REVIEW_NO_SAVEAS"
        Case 8: MBF_GroupLabel = "UNMATCHED_NO_SAVEAS"
        Case 9: MBF_GroupLabel = "NEW_LIMITED_VALIDATION"
        Case 10: MBF_GroupLabel = "REUSE_LIMITED_VALIDATION"
        Case 11: MBF_GroupLabel = "ABORTED_ATTEMPT"
        Case Else: MBF_GroupLabel = "UNEXPECTED_RESULT"
    End Select
End Function

Private Function MBF_BinUpper(ByVal bin As Long) As Double
    MBF_BinUpper = MBF_FIRST_UPPER_SECONDS * (2# ^ (bin - 1))
End Function

Private Function MBF_BinFor(ByVal seconds As Double) As Long
    Dim bin As Long, upper As Double
    upper = MBF_FIRST_UPPER_SECONDS
    For bin = 1 To MBF_BIN_COUNT - 1
        If seconds <= upper Then MBF_BinFor = bin: Exit Function
        upper = upper * 2#
    Next bin
    MBF_BinFor = MBF_BIN_COUNT
End Function

Private Sub MBF_Record(ByVal mode As String, ByVal outcome As Long, _
                       ByVal elapsedSeconds As Double, Optional ByVal contentLimited As Boolean = False)
    Dim group As Long, bin As Long
    If mode <> "AUTO" Then Exit Sub
    On Error GoTo Unavailable
    group = MBF_GroupFor(outcome, contentLimited)
    gMetricObserved(group) = gMetricObserved(group) + 1#
    ' A broken/backwards clock is an untimed result, never a zero-second sample.
    If elapsedSeconds < 0# Then Exit Sub
    bin = MBF_BinFor(elapsedSeconds)
    If gMetricTimed(group) = 0# Then
        gMetricMinimum(group) = elapsedSeconds
        gMetricMaximum(group) = elapsedSeconds
    Else
        If elapsedSeconds < gMetricMinimum(group) Then gMetricMinimum(group) = elapsedSeconds
        If elapsedSeconds > gMetricMaximum(group) Then gMetricMaximum(group) = elapsedSeconds
    End If
    gMetricSeconds(group) = gMetricSeconds(group) + elapsedSeconds
    gMetricTimed(group) = gMetricTimed(group) + 1#
    gMetricHistogram(group, bin) = gMetricHistogram(group, bin) + 1#
    If elapsedSeconds = 0# Then gMetricZeros(group) = gMetricZeros(group) + 1#
    Exit Sub
Unavailable:
    ' Telemetry must not turn a durable backup into a failed backup. Suppress all
    ' forecasts after a metrics error; partially updated statistics are not proof.
    gMetricFault = True
    gMetricFaultText = CStr(Err.Number) & " - " & Err.Description
    Err.Clear
End Sub

Private Function MBF_QuantileInterval(ByVal group As Long, ByVal fraction As Double) As String
    Dim target As Double, cumulative As Double, bin As Long
    If gMetricTimed(group) = 0# Then MBF_QuantileInterval = "NOT_MEASURED": Exit Function
    target = -Int(-gMetricTimed(group) * fraction)
    For bin = 1 To MBF_BIN_COUNT
        cumulative = cumulative + gMetricHistogram(group, bin)
        If cumulative >= target Then
            If bin = 1 Then
                MBF_QuantileInterval = "[0, " & Format$(MBF_BinUpper(bin), "0.000000") & "]"
            ElseIf bin = MBF_BIN_COUNT Then
                MBF_QuantileInterval = "(" & Format$(MBF_BinUpper(bin - 1), "0.000000") & ", +infinity)"
            Else
                MBF_QuantileInterval = "(" & Format$(MBF_BinUpper(bin - 1), "0.000000") & ", " & _
                                       Format$(MBF_BinUpper(bin), "0.000000") & "]"
            End If
            Exit Function
        End If
    Next bin
    MBF_QuantileInterval = "UNAVAILABLE_INCONSISTENT_COUNTERS"
End Function

Private Function MBF_GroupSummary(ByVal group As Long) As String
    Dim s As String
    s = MBF_GroupLabel(group) & ": observed=" & Format$(gMetricObserved(group), "0") & _
        "; timed=" & Format$(gMetricTimed(group), "0") & _
        "; untimed=" & Format$(gMetricObserved(group) - gMetricTimed(group), "0")
    If gMetricTimed(group) > 0# Then
        s = s & "; total seconds=" & Format$(gMetricSeconds(group), "0.000000") & _
            "; mean=" & Format$(gMetricSeconds(group) / gMetricTimed(group), "0.000000") & _
            "; min/max=" & Format$(gMetricMinimum(group), "0.000000") & "/" & _
            Format$(gMetricMaximum(group), "0.000000")
        s = s & "; histogram p50=" & MBF_QuantileInterval(group, 0.5) & _
            "; p95=" & MBF_QuantileInterval(group, 0.95) & _
            "; zero-duration=" & Format$(gMetricZeros(group), "0")
    Else
        s = s & "; seconds=NOT_MEASURED"
    End If
    MBF_GroupSummary = s
End Function

Private Function MBF_Forecast(ByVal group As Long, ByVal label As String) As String
    Dim s As String, mean As Double
    s = "100,000 " & label & " indicative processing hours: "
    If gMetricFault Then
        MBF_Forecast = s & "UNAVAILABLE_METRICS_ERROR": Exit Function
    End If
    If gMetricObserved(group) <> gMetricTimed(group) Then
        MBF_Forecast = s & "UNAVAILABLE_UNTIMED_RESULTS": Exit Function
    End If
    If gMetricTimed(group) = 0# Then
        MBF_Forecast = s & "NOT_MEASURED (no eligible actual AUTO observations)": Exit Function
    End If
    If gMetricTimed(group) < MBF_MIN_FORECAST_COUNT Then
        MBF_Forecast = s & "NOT_ESTIMATED (eligible timed observations=" & _
                       Format$(gMetricTimed(group), "0") & "; minimum=" & CStr(MBF_MIN_FORECAST_COUNT) & ")"
        Exit Function
    End If
    mean = gMetricSeconds(group) / gMetricTimed(group)
    If mean <= 0# Then
        MBF_Forecast = s & "NOT_ESTIMATED (clock resolution insufficient)": Exit Function
    End If
    MBF_Forecast = s & Format$(mean * 100000# / 3600#, "0.0000") & _
                   " (eligible timed observations=" & Format$(gMetricTimed(group), "0") & _
                   "; observed mean seconds=" & Format$(mean, "0.000000") & ")"
End Function

Private Function MBF_Summary(ByVal mode As String) As String
    Dim s As String, group As Long
    If mode <> "AUTO" Then Exit Function
    On Error GoTo Unavailable
    s = vbCrLf & "Automatic performance observations: actual AUTO item processing; no extra SaveAs." & vbCrLf
    If gMetricFault Then
        s = s & "METRICS ERROR: counters may be incomplete; all forecasts suppressed. " & gMetricFaultText & vbCrLf
    End If
    For group = 1 To MBF_GROUP_COUNT
        s = s & MBF_GroupSummary(group) & vbCrLf
    Next group
    s = s & MBF_Forecast(1, "NEW") & vbCrLf & MBF_Forecast(2, "REUSE") & vbCrLf
    s = s & "Histogram intervals contain p50/p95 of all timed observations; they are not exact percentiles." & vbCrLf
    s = s & "Forecasts cover only successful NEW or REUSE item processing; limited-validation results, " & _
            "failures, REVIEW and cancelled/aborted attempts are separate above." & vbCrLf
    s = s & "Forecasts exclude inventory, legacy registration, organization, recovery, archive audits, " & _
            "final commits/reports and future retries. No whole-run duration is predicted." & vbCrLf
    s = s & "Observed item costs include the current configured export, candidate checks, validation and staging. " & _
            "Comparable mail sizes, attachments, provider, disk/network and cache are required." & vbCrLf
    s = s & "REUSE cost also depends on the selected content-audit scope. Work can grow nonlinearly as " & _
            "metadata, journals and output grow; a daily success mix is not a 100,000-item workload." & vbCrLf
    s = s & "The minimum observation count is a reporting threshold, not a statistical confidence guarantee. " & _
            "These observations do not replace a real 100,000-item test or every-MAPI fidelity testing."
    MBF_Summary = s
    Exit Function
Unavailable:
    MBF_Summary = vbCrLf & "Automatic performance observations: REPORT_UNAVAILABLE; no forecast is valid."
    Err.Clear
End Function


' =================== Bounded ordinary metadata fidelity ====================
' Importance, Sensitivity and category NAME membership are stable comparison
' fields. Compare the documented PT_MV_UNICODE Keywords property, rather than
' splitting Categories on a guessed locale separator or trimming category names.
' Commas, semicolons, Unicode, case and significant spaces remain exact names.
' UserProperties are explicitly outside this comparator: any nonzero collection
' makes a new MSG BASIC/limited, and prevents an equal-content legacy/R alias.
' Flags/tasks/reminders and arbitrary MAPI properties remain outside this scope.

Private Function MBE_CategorySet(ByVal value As Variant, ByVal displayCategories As String) As Object
    Dim names As Object, lower As Long, upper As Long, dimension As Long, n As Long
    Dim hasSecondDimension As Boolean, i As Long, name As String
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbBinaryCompare
    If MBL_MissingProperty(value, "message category Keywords") Then
        If Len(displayCategories) > 0 Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
            "Categories are present but their exact Keywords names are unavailable."
        Set MBE_CategorySet = names
        Exit Function
    End If
    If Not IsArray(value) Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
        "Category Keywords did not return a string array."
    ' A malformed/unallocated/multidimensional array is unknown, never empty.
    On Error Resume Next
    Err.Clear
    lower = LBound(value, 1): upper = UBound(value, 1): n = Err.Number
    Err.Clear
    dimension = LBound(value, 2): hasSecondDimension = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If n <> 0 Or hasSecondDimension Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
        "Category Keywords array shape is unsupported."
    If CDbl(upper) - CDbl(lower) + 1# > MBE_MAX_CATEGORIES Then _
        Err.Raise MB_ERR + 1124, "MBE_CategorySet", "Category count exceeds the bounded comparison limit."
    For i = lower To upper
        If VarType(value(i)) <> vbString Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
            "Category Keywords contains a non-string name."
        name = CStr(value(i))
        If Len(name) = 0 Or Len(name) >= 256 Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
            "Category name length is outside the documented supported range."
        names(name) = True
    Next i
    If names.Count = 0 Then
        If Len(displayCategories) > 0 Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
            "Nonempty Categories has no exact Keywords names."
    Else
        If Len(displayCategories) = 0 Then Err.Raise MB_ERR + 1124, "MBE_CategorySet", _
            "Category Keywords and the item Categories view disagree."
    End If
    Set MBE_CategorySet = names
End Function

Private Function MBE_ReadCategorySet(ByVal mail As Outlook.MailItem) As Object
    Dim values As Variant, first As Long
    values = MBL_ReadProperties(mail, Array(MBE_CATEGORY_KEYWORDS), 1)
    first = LBound(values)
    Set MBE_ReadCategorySet = MBE_CategorySet(values(first), mail.Categories)
End Function

Private Function MBE_SameCategorySets(ByVal first As Object, ByVal second As Object) As Boolean
    Dim key As Variant
    If first.Count <> second.Count Then Exit Function
    For Each key In first.Keys
        If Not second.Exists(CStr(key)) Then Exit Function
    Next key
    MBE_SameCategorySets = True
End Function

Private Function MBE_MetadataComplete(ByVal source As Outlook.MailItem, ByVal copy As Outlook.MailItem) As Boolean
    Dim firstImportance As Long, secondImportance As Long, firstSensitivity As Long, secondSensitivity As Long
    Dim firstCategories As Object, secondCategories As Object, firstCustom As Long, secondCustom As Long
    firstImportance = CLng(source.Importance): secondImportance = CLng(copy.Importance)
    If firstImportance < 0 Or firstImportance > 2 Or secondImportance < 0 Or secondImportance > 2 Then _
        Err.Raise MB_ERR + 1124, "MBE_MetadataComplete", "Unsupported Importance value."
    If firstImportance <> secondImportance Then _
        Err.Raise MB_ERR + 1120, "MBE_MetadataComplete", "Message Importance mismatch."
    firstSensitivity = CLng(source.Sensitivity): secondSensitivity = CLng(copy.Sensitivity)
    If firstSensitivity < 0 Or firstSensitivity > 3 Or secondSensitivity < 0 Or secondSensitivity > 3 Then _
        Err.Raise MB_ERR + 1124, "MBE_MetadataComplete", "Unsupported Sensitivity value."
    If firstSensitivity <> secondSensitivity Then _
        Err.Raise MB_ERR + 1121, "MBE_MetadataComplete", "Message Sensitivity mismatch."
    Set firstCategories = MBE_ReadCategorySet(source)
    Set secondCategories = MBE_ReadCategorySet(copy)
    If Not MBE_SameCategorySets(firstCategories, secondCategories) Then _
        Err.Raise MB_ERR + 1122, "MBE_MetadataComplete", "Message category name membership mismatch."
    firstCustom = source.UserProperties.Count: secondCustom = copy.UserProperties.Count
    If firstCustom < 0 Or secondCustom < 0 Then _
        Err.Raise MB_ERR + 1124, "MBE_MetadataComplete", "Invalid custom-property count."
    ' Do not infer equal custom values from equal names/counts. Custom formulas,
    ' types and provider-specific fields require a separately declared contract.
    MBE_MetadataComplete = (firstCustom = 0 And secondCustom = 0)
End Function

Private Sub MBE_RequireMetadataComplete(ByVal source As Outlook.MailItem, ByVal copy As Outlook.MailItem)
    If Not MBE_MetadataComplete(source, copy) Then Err.Raise MB_ERR + 1123, "MBE_RequireMetadataComplete", _
        "Custom UserProperties are present; full bounded metadata equality is unproven. Existing MSG is retained for review."
End Sub

Private Function MBE_CategorySelfTest() As Boolean
    Dim first As Object, second As Object, missing As Variant, unicodeName As String
    On Error GoTo Failed
    missing = CVErr(MBL_MAPI_NOT_FOUND)
    Set first = MBE_CategorySet(missing, "")
    Set second = MBE_CategorySet(Array(), "")
    If Not MBE_SameCategorySets(first, second) Then Exit Function
    unicodeName = ChrW(&HD55C) & ChrW(&HAE00)
    Set first = MBE_CategorySet(Array("comma,name", "semi;colon", " padded ", unicodeName), "display A")
    Set second = MBE_CategorySet(Array(unicodeName, " padded ", "semi;colon", "comma,name"), "display B")
    If Not MBE_SameCategorySets(first, second) Then Exit Function
    Set second = MBE_CategorySet(Array(unicodeName, "padded", "semi;colon", "comma,name"), "display B")
    If MBE_SameCategorySets(first, second) Then Exit Function
    If Not MBE_CategoryRejects(missing, "present") Then Exit Function
    If Not MBE_CategoryRejects(Null, "") Then Exit Function
    If Not MBE_CategoryRejects(Empty, "") Then Exit Function
    If Not MBE_CategoryRejects("single", "single") Then Exit Function
    If Not MBE_CategoryRejects(Array(123), "123") Then Exit Function
    If Not MBE_CategoryRejects(Array(""), "") Then Exit Function
    If Not MBE_CategoryRejects(Array("present"), "") Then Exit Function
    If Not MBE_CategoryRejects(CVErr(&H80070005), "") Then Exit Function
    MBE_CategorySelfTest = True
Failed:
End Function

Private Function MBE_CategoryRejects(ByVal value As Variant, ByVal display As String) As Boolean
    Dim names As Object
    On Error GoTo Rejected
    Set names = MBE_CategorySet(value, display)
    Exit Function
Rejected:
    MBE_CategoryRejects = True
End Function

Private Function MBE_MetadataResultCode(ByVal source As Outlook.MailItem, ByVal copy As Outlook.MailItem) As Long
    On Error GoTo DifferentOrUnknown
    MBE_RequireMetadataComplete source, copy
    Exit Function
DifferentOrUnknown:
    MBE_MetadataResultCode = Err.Number - MB_ERR
End Function

Private Function MBE_SyntheticMetadataTest(ByVal source As Outlook.MailItem, ByVal path As String, _
                                          ByRef detail As String) As Boolean
    Dim obj As Object, copy As Outlook.MailItem, custom As Outlook.UserProperty
    Dim oldImportance As Long, oldSensitivity As Long, oldCategories As String, captured As Boolean
    Dim n As Long, reason As String, cleanupIssue As String
    On Error GoTo Failed
    ' This mutation test is private to the already-owned, unsent acceptance MSG.
    ' It never calls Save/Send/SaveAs and never receives a live mailbox item.
    If gMode <> "SELFTEST_MSG" Or gRunId <> "synthetic_msg" Then _
        Err.Raise MB_ERR + 1125, , "Metadata mutation test requires the disposable acceptance context."
    If Not MB_PathWithin(path, gRoot) Then _
        Err.Raise MB_ERR + 1125, , "Metadata test MSG is outside the disposable test root."
    If source.Sent Then Err.Raise MB_ERR + 1125, , "Metadata test source must be unsent."
    If source.UserProperties.Count <> 0 Then _
        Err.Raise MB_ERR + 1125, , "Metadata test source already has unowned custom properties."
    oldImportance = source.Importance: oldSensitivity = source.Sensitivity: oldCategories = source.Categories
    captured = True
    Set obj = Application.Session.OpenSharedItem(path)
    If Not TypeOf obj Is Outlook.MailItem Then Err.Raise MB_ERR + 1125, , "Metadata test copy is not MailItem."
    Set copy = obj: Set obj = Nothing
    If MBE_MetadataResultCode(source, copy) <> 0 Then _
        Err.Raise MB_ERR + 1125, , "Initial stable metadata does not match its owned MSG."
    source.Importance = (oldImportance + 1) Mod 3
    If MBE_MetadataResultCode(source, copy) <> 1120 Then _
        Err.Raise MB_ERR + 1125, , "Importance-only change was not identified."
    source.Importance = oldImportance
    source.Sensitivity = (oldSensitivity + 1) Mod 4
    If MBE_MetadataResultCode(source, copy) <> 1121 Then _
        Err.Raise MB_ERR + 1125, , "Sensitivity-only change was not identified."
    source.Sensitivity = oldSensitivity
    source.Categories = "MB34_Fidelity_" & ChrW(&HD55C) & ChrW(&HAE00)
    If MBE_MetadataResultCode(source, copy) <> 1122 Then _
        Err.Raise MB_ERR + 1125, , "Category-only change was not identified."
    source.Categories = oldCategories
    Set custom = source.UserProperties.Add("MB34FidelityCustom", olText, False)
    custom.Value = "Owned unsaved metadata test"
    If MBE_MetadataComplete(source, copy) Then _
        Err.Raise MB_ERR + 1125, , "Custom properties were incorrectly classified as complete."
    If MBE_MetadataResultCode(source, copy) <> 1123 Then _
        Err.Raise MB_ERR + 1125, , "Custom-property reuse was not held for review."
    custom.Delete: Set custom = Nothing
    If MBE_MetadataResultCode(source, copy) <> 0 Then _
        Err.Raise MB_ERR + 1125, , "Restored in-memory metadata did not return to equality."
    GoTo Cleanup
Failed:
    n = Err.Number: reason = Err.Description
Cleanup:
    On Error Resume Next
    Err.Clear
    If Not custom Is Nothing Then custom.Delete
    If Err.Number <> 0 Then cleanupIssue = " Custom-property cleanup failed: " & Err.Description
    Set custom = Nothing
    Err.Clear
    If captured Then
        source.Importance = oldImportance
        source.Sensitivity = oldSensitivity
        source.Categories = oldCategories
    End If
    If Err.Number <> 0 Then cleanupIssue = cleanupIssue & " Metadata restore failed: " & Err.Description
    Err.Clear
    If Not copy Is Nothing Then copy.Close olDiscard
    If Not obj Is Nothing Then obj.Close olDiscard
    If Err.Number <> 0 Then cleanupIssue = cleanupIssue & " Copy close failed: " & Err.Description
    Set copy = Nothing: Set obj = Nothing
    On Error GoTo 0
    MBE_SyntheticMetadataTest = (n = 0 And Len(cleanupIssue) = 0)
    If MBE_SyntheticMetadataTest Then
        detail = "PASS: owned unsaved Importance/Sensitivity/Unicode-category differences and custom-property REVIEW."
    Else
        detail = "FAIL: " & CStr(n) & " - " & reason & cleanupIssue
    End If
End Function


' A restore plan owns immutable, exactly hashed copies of the still-available
' primary metadata BEFORE the old snapshot is installed. Its durable pointer
' survives interruption. Valid later proof is merged before DONE; a conflict
' never authorizes a new digest, a new baseline or a new MSG export.
Private Function MBR_CaptureOverlay(ByVal recoveryRoot As String, ByVal accountRoot As String, _
                                    ByVal storeId As String, ByVal baseReference As String) As String
    Dim dir As String, primary As String, manifest As String, stage As String, copyPath As String
    Dim source As String, leaf As String, nonce As String, baseDigest As String, digest As String
    Dim bytes As Double, manifestBytes As Double, count As Long, item As Object, stream As Object
    Dim paths As New Collection, key As Variant, fields As Variant, n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    dir = MBH_JoinPath(recoveryRoot, "overlays")
    MBY_NoReparse dir, False: MBH_EnsureDirectory dir
    nonce = Left$(MBH_HashText(gRunId & "|" & gFSO.GetTempName & "|" & CStr(Timer)), 24)
    manifest = MBH_JoinPath(dir, "plan_" & nonce & ".tsv")
    If baseReference = "MSG_ONLY" Then
        baseDigest = ""
    ElseIf Left$(baseReference, 8) = "SALVAGE|" Then
        fields = Split(baseReference, "|")
        If UBound(fields) <> 4 Then Err.Raise MB_ERR + 1140, , "Malformed prior salvage intent."
        baseDigest = CStr(fields(3))
        If Not MB_IsHexKey(baseDigest) Then Err.Raise MB_ERR + 1140, , "Invalid prior salvage digest."
    Else
        If Not MB_PathWithin(baseReference, MBH_JoinPath(recoveryRoot, "snapshots")) Then _
            Err.Raise MB_ERR + 1141, , "Restore base escapes its snapshots directory."
        MBY_NoReparse baseReference, True
        baseDigest = MBH_HashFile(baseReference)
    End If
    stage = MBH_NewStagePath(manifest)
    Set stream = gFSO.CreateTextFile(stage, True, True)
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBOVERLAY1", "HEADER", gRoot, accountRoot, _
        storeId, baseReference, baseDigest, MBC_UtcStamp(MBC_UtcNow())))
    primary = MB_StatePath()
    If gFSO.FolderExists(primary) Then
        For Each item In gFSO.GetFolder(primary).Files
            If MBR_IsStateLeaf(CStr(item.Name)) Then
                If LCase$(CStr(item.Name)) = MBR_MIRROR_LEAF Then _
                    Err.Raise MB_ERR + 1155, , "Reserved recovery-overlay leaf exists in primary state."
                paths.Add Array(CStr(item.Path), CStr(item.Name))
            End If
        Next item
    End If
    source = MBH_JoinPath(recoveryRoot, "stage-intents.tsv")
    If gFSO.FileExists(source) Then paths.Add Array(source, MBR_MIRROR_LEAF)
    For Each key In paths
        If Not MB_ContinueWork() Then Err.Raise MB_ERR + 1158, , "Restore capture cancelled before state replacement."
        source = CStr(key(0)): leaf = CStr(key(1))
        MBY_NoReparse source, True
        guard = MBH_GuardFile(source)
        bytes = CDbl(gFSO.GetFile(source).Size): digest = MBH_HashFile(source)
        count = count + 1
        copyPath = MBH_JoinPath(dir, nonce & "_" & Format$(count, "00000") & ".tsv")
        gFSO.CopyFile source, copyPath, False
        MBM_FlushPath copyPath
        If CDbl(gFSO.GetFile(copyPath).Size) <> bytes Or MBH_HashFile(copyPath) <> digest Then _
            Err.Raise MB_ERR + 1142, , "Pre-restore metadata copy failed exact-byte verification."
        MBH_CloseGuard guard
        MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBOVERLAY1", "FILE", leaf, _
            CStr(gFSO.GetFileName(copyPath)), Format$(bytes, "0"), digest, "", ""))
    Next key
    MB_WriteCheckedLine stream, MB_EncodeFields(Array("MBOVERLAY1", "END", CStr(count), "", "", "", "", ""))
    stream.Close: Set stream = Nothing
    MBM_FlushPath stage
    digest = MBH_HashFile(stage): manifestBytes = CDbl(gFSO.GetFile(stage).Size)
    MBH_CommitStage stage, manifest, digest, manifestBytes: stage = ""
    MBR_CaptureOverlay = "MERGE|" & CStr(gFSO.GetFileName(manifest)) & "|" & _
        Format$(manifestBytes, "0") & "|" & digest
    Exit Function
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    If Not stream Is Nothing Then stream.Close
    If Len(stage) > 0 Then MBH_DeleteStage stage
    ' Completed immutable copies remain available even if a plan could not be
    ' published. No primary state has been replaced at this point.
    On Error GoTo 0
    Err.Raise n, "MBR_CaptureOverlay", detail
End Function

Private Sub MBR_ReadOverlay(ByVal recoveryRoot As String, ByVal accountRoot As String, ByVal storeId As String, _
                            ByVal pointer As String, ByRef files As Object, ByRef baseReference As String, _
                            ByRef baseDigest As String)
    Dim parts As Variant, f As Variant, path As String, line As String, stream As Object
    Dim header As Boolean, footer As Boolean, n As Long, detail As String, leaf As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    parts = Split(pointer, "|")
    If UBound(parts) <> 3 Then Err.Raise MB_ERR + 1143, , "Invalid restore-overlay pointer."
    If CStr(parts(0)) <> "MERGE" Or Not MBR_IsStateLeaf(CStr(parts(1))) Or _
       Not MB_IsHexKey(CStr(parts(3))) Then Err.Raise MB_ERR + 1143, , "Unsafe restore-overlay pointer."
    path = MBH_JoinPath(MBH_JoinPath(recoveryRoot, "overlays"), CStr(parts(1)))
    MBY_NoReparse path, True
    guard = MBH_GuardFile(path)
    If CDbl(gFSO.GetFile(path).Size) <> MB_ByteCount(CStr(parts(2))) Or _
       MBH_HashFile(path) <> CStr(parts(3)) Then Err.Raise MB_ERR + 1144, , "The original restore plan changed or is missing."
    Set files = CreateObject("Scripting.Dictionary"): files.CompareMode = vbTextCompare
    Set stream = gFSO.OpenTextFile(path, 1, False, -1)
    Do Until stream.AtEndOfStream
        line = stream.ReadLine
        If Len(line) > 0 Then
            If footer Then Err.Raise MB_ERR + 1145, , "Restore-overlay data follows its footer."
            f = MB_ReadCheckedLine(line, 8)
            If CStr(f(0)) <> "MBOVERLAY1" Then Err.Raise MB_ERR + 1145, , "Unknown restore-overlay schema."
            Select Case CStr(f(1))
                Case "HEADER"
                    If header Or files.Count > 0 Then Err.Raise MB_ERR + 1145, , "Duplicate/late overlay header."
                    If StrComp(CStr(f(2)), gRoot, vbTextCompare) <> 0 Or _
                       StrComp(CStr(f(3)), accountRoot, vbTextCompare) <> 0 Or _
                       StrComp(CStr(f(4)), storeId, vbTextCompare) <> 0 Then _
                        Err.Raise MB_ERR + 1146, , "Restore-overlay belongs to another root/account/store."
                    If Not MBC_ValidDate(MBC_ParseUtcStamp(CStr(f(7)))) Then _
                        Err.Raise MB_ERR + 1145, , "Restore-overlay timestamp is invalid."
                    baseReference = CStr(f(5)): baseDigest = CStr(f(6)): header = True
                Case "FILE"
                    leaf = CStr(f(2))
                    If Not header Or Not MBR_IsStateLeaf(leaf) Or Not MBR_IsStateLeaf(CStr(f(3))) Or _
                       Not MB_IsHexKey(CStr(f(5))) Then Err.Raise MB_ERR + 1147, , "Unsafe overlay metadata entry."
                    If files.Exists(leaf) Then Err.Raise MB_ERR + 1147, , "Duplicate overlay metadata filename."
                    If MB_ByteCount(CStr(f(4))) < 0 Then Err.Raise MB_ERR + 1147, , "Invalid overlay metadata size."
                    files.Add leaf, MB_EncodeFields(Array(f(3), f(4), f(5)))
                Case "END"
                    If Not header Or MB_ByteCount(CStr(f(2))) <> files.Count Then _
                        Err.Raise MB_ERR + 1148, , "Incomplete restore-overlay footer."
                    footer = True
                Case Else
                    Err.Raise MB_ERR + 1145, , "Unknown restore-overlay record."
            End Select
        End If
    Loop
    stream.Close: Set stream = Nothing
    MBH_CloseGuard guard
    If Not header Or Not footer Then Err.Raise MB_ERR + 1148, , "Incomplete restore-overlay manifest."
    If baseReference = "MSG_ONLY" Then
        If Len(baseDigest) <> 0 Then Err.Raise MB_ERR + 1149, , "Unexpected MSG-only base digest."
    ElseIf Left$(baseReference, 8) = "SALVAGE|" Then
        parts = Split(baseReference, "|")
        If UBound(parts) <> 4 Then Err.Raise MB_ERR + 1149, , "Invalid overlay salvage base."
        If Not MB_IsHexKey(baseDigest) Or CStr(parts(3)) <> baseDigest Then _
            Err.Raise MB_ERR + 1149, , "Overlay salvage digest does not match its prior intent."
    Else
        If Not MB_IsHexKey(baseDigest) Or _
           Not MB_PathWithin(baseReference, MBH_JoinPath(recoveryRoot, "snapshots")) Then _
            Err.Raise MB_ERR + 1149, , "Unsafe overlay snapshot base."
        MBY_NoReparse baseReference, True
        If MBH_HashFile(baseReference) <> baseDigest Then _
            Err.Raise MB_ERR + 1149, , "The original restore snapshot no longer has its captured digest."
    End If
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    MBH_CloseGuard guard
    On Error GoTo 0
    Err.Raise n, "MBR_ReadOverlay", detail
End Sub

Private Sub MBR_PublishOverlayStage(ByRef stage As String, ByVal finalPath As String)
    Dim digest As String, bytes As Double, n As Long, detail As String
    On Error GoTo Failed
    MBM_FlushPath stage
    digest = MBH_HashFile(stage): bytes = CDbl(gFSO.GetFile(stage).Size)
    If gFSO.FileExists(finalPath) Then
        If CDbl(gFSO.GetFile(finalPath).Size) = bytes Then
            If MBH_HashFile(finalPath) = digest Then
                MBH_DeleteStage stage: stage = "": Exit Sub
            End If
        End If
        MBR_PreserveFile finalPath, MB_StatePath("preserved")
    End If
    MBH_CommitStage stage, finalPath, digest, bytes
    stage = ""
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    Err.Raise n, "MBR_PublishOverlayStage", detail
End Sub

Private Function MBR_ParseIndexFields(ByVal line As String) As Variant
    Dim f As Variant, raw As Variant, count As Long
    raw = Split(line, vbTab): count = UBound(raw) - 1
    If count < 1 Then Err.Raise MB_ERR + 15, , "Index row has no checked payload framing."
    ' Establish whether the framing IS checked before interpreting a schema.
    ' An unknown but valid future schema is not a disposable truncated row.
    f = MB_ReadCheckedLine(line, count)
    If CStr(f(0)) = "MBS4" And count = 17 Then
        If CStr(f(1)) <> "N" And CStr(f(1)) <> "T" And CStr(f(1)) <> "R" And CStr(f(1)) <> "L" Then _
            Err.Raise MB_ERR + 1154, , "Unknown checked MBS4 record kind requires review."
    ElseIf CStr(f(0)) = "MBS3" And count = 17 Then
        If CStr(f(1)) <> "N" And CStr(f(1)) <> "T" And CStr(f(1)) <> "R" Then _
            Err.Raise MB_ERR + 1154, , "Unknown checked MBS3 record kind requires review."
    ElseIf CStr(f(0)) = "MBS2" And count = 16 Then
        If CStr(f(1)) <> "M" And CStr(f(1)) <> "A" And CStr(f(1)) <> "L" Then _
            Err.Raise MB_ERR + 1154, , "Unknown checked MBS2 record kind requires review."
    Else
        Err.Raise MB_ERR + 1154, , "Unsupported checked journal schema/field count requires review."
    End If
    MBR_ParseIndexFields = f
End Function

Private Sub MBR_RejectIndexConflict(ByVal f As Variant)
    Dim key As String, prior As Variant, meta As Variant
    key = MBX_IndexKey(CStr(f(1)), CStr(f(2)))
    If Not gIndex.Exists(key) Then Exit Sub
    prior = MB_IndexFields(key): meta = MBX_Meta(key)
    If CStr(prior(2)) <> CStr(f(5)) Or MB_ByteCount(CStr(prior(1))) <> MB_ByteCount(CStr(f(4))) Or _
       CStr(meta(2)) <> CStr(f(6)) Or CStr(meta(0)) <> MBX_StoreToken(CStr(f(7))) Then _
        Err.Raise MB_ERR + 1150, , "Conflicting historical expected bytes/source for one immutable index key: " & key
End Sub

Private Function MBR_TryMergeIndexRow(ByVal line As String, ByVal storeId As String, ByRef reason As String) As Boolean
    Dim f As Variant, n As Long
    On Error GoTo Rejected
    f = MBR_ParseIndexFields(line)
    If Len(storeId) = 0 Then Err.Raise MB_ERR + 1156, , "Journal merge lacks an expected store binding."
    MBS_RequireStore CStr(f(7)), storeId, "immutable restore overlay"
    ' Compare before the real validator can replace an existing cache entry.
    MBR_RejectIndexConflict f
    If Not MBR_TryIndexRow(line, reason) Then _
        Err.Raise MB_ERR + 1154, , "Checked historical proof has unresolved semantics; preserved for review: " & reason
    MBR_TryMergeIndexRow = True
    Exit Function
Rejected:
    n = Err.Number: reason = Err.Description
    If n >= MB_ERR + 15 And n <= MB_ERR + 17 Then Exit Function
    ' A well-framed conflicting expected hash is not disposable corrupt input.
    ' Leave START and both immutable proof sources available for explicit review.
    Err.Raise n, "MBR_TryMergeIndexRow", reason
End Function

Private Sub MBR_MergeJournal(ByVal priorRaw As String, ByVal storeId As String)
    Dim oldIndex As Object, oldExternal As Object, oldMeta As Object, oldStores As Object
    Dim oldBindings As Object, oldSources As Object, oldOwners As Object
    Dim oldBuckets As Object, oldSeen As Object, oldKnown As Object, oldHeads As Object, oldLookup As Object
    Dim inputStream As Object, outputStream As Object, baseRows As New Collection, paths(0 To 1) As String
    Dim source As String, finalPath As String, stage As String, line As String, reason As String, rowHash As String
    Dim pass As Long, added As Long, rejected As Long, cached As Boolean, n As Long, detail As String
    Dim position As Long, matchingPrefix As Boolean, skipPrefixRow As Boolean, visited As Long
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    finalPath = MB_StatePath("journal.tsv")
    paths(0) = finalPath: paths(1) = priorRaw
    Set oldIndex = gIndex: Set oldExternal = gExternalRoots: Set oldMeta = gRecordMeta: Set oldStores = gStoreTokens
    Set oldBindings = gLegacyBindings: Set oldSources = gLegacySources: Set oldOwners = gLegacyOwners
    Set oldBuckets = gIdentityBuckets: Set oldSeen = gIdentitySeen: Set oldKnown = gWindowKnown
    Set oldLookup = gIdentityLookupState
    Set oldHeads = gRevisionHeads: cached = True
    Set gIndex = CreateObject("Scripting.Dictionary")
    Set gExternalRoots = CreateObject("Scripting.Dictionary")
    Set gRecordMeta = CreateObject("Scripting.Dictionary")
    Set gStoreTokens = CreateObject("Scripting.Dictionary"): gStoreTokens.CompareMode = vbTextCompare
    Set gLegacyBindings = CreateObject("Scripting.Dictionary")
    Set gLegacySources = CreateObject("Scripting.Dictionary")
    Set gLegacyOwners = CreateObject("Scripting.Dictionary"): gLegacyOwners.CompareMode = vbTextCompare
    Set gIdentityBuckets = Nothing: Set gIdentitySeen = Nothing: Set gRevisionHeads = Nothing
    Set gWindowKnown = CreateObject("Scripting.Dictionary")
    matchingPrefix = True
    stage = MBH_NewStagePath(finalPath)
    Set outputStream = gFSO.CreateTextFile(stage, True, True)
    For pass = 0 To 1
        source = paths(pass)
        If gFSO.FileExists(source) Then
            MBY_NoReparse source, True
            guard = MBH_GuardFile(source)
            If CDbl(gFSO.GetFile(source).Size) Mod 2 <> 0 Then _
                Err.Raise MB_ERR + 1151, , "An odd-byte journal requires review; its immutable original is preserved."
            Set inputStream = gFSO.OpenTextFile(source, 1, False, -1)
            Do Until inputStream.AtEndOfStream
                line = inputStream.ReadLine
                visited = visited + 1
                If visited Mod 256 = 0 Then
                    If Not MB_ContinueWork() Then Err.Raise MB_ERR + 1158, , "Journal merge cancelled; original restore plan retained."
                End If
                If Len(line) > 0 Then
                    rowHash = MBH_HashText(line)
                    skipPrefixRow = False
                    If pass = 1 Then
                        position = position + 1
                        If matchingPrefix Then
                            If position <= baseRows.Count Then
                                If rowHash = CStr(baseRows(position)) Then skipPrefixRow = True
                            End If
                            If Not skipPrefixRow Then matchingPrefix = False
                        End If
                    End If
                    If Not skipPrefixRow Then
                        If MBR_TryMergeIndexRow(line, storeId, reason) Then
                            outputStream.WriteLine line
                            If pass = 0 Then baseRows.Add rowHash Else added = added + 1
                        ElseIf pass = 0 Then
                            Err.Raise MB_ERR + 1152, , "The restored journal is not a valid semantic base: " & reason
                        Else
                            rejected = rejected + 1
                            MBR_Log "WARNING", "RECOVERY_OVERLAY_ROW_UNPROVEN", CStr(rejected), priorRaw, 0, reason
                        End If
                    End If
                End If
            Loop
            inputStream.Close: Set inputStream = Nothing
            MBH_CloseGuard guard
        End If
    Next pass
    outputStream.Close: Set outputStream = Nothing
    MBR_PublishOverlayStage stage, finalPath
    gRecoveryMergedRows = gRecoveryMergedRows + added
    gRecoveryOverlayRejected = gRecoveryOverlayRejected + rejected
    If rejected > 0 Then gRecoveryRebased = True
    GoTo Cleanup
Failed:
    n = Err.Number: detail = Err.Description
Cleanup:
    On Error Resume Next
    If Not inputStream Is Nothing Then inputStream.Close
    If Not outputStream Is Nothing Then outputStream.Close
    MBH_CloseGuard guard
    If Len(stage) > 0 Then MBH_DeleteStage stage
    If cached Then
        Set gIndex = oldIndex: Set gExternalRoots = oldExternal: Set gRecordMeta = oldMeta: Set gStoreTokens = oldStores
        Set gLegacyBindings = oldBindings: Set gLegacySources = oldSources: Set gLegacyOwners = oldOwners
        Set gIdentityBuckets = oldBuckets: Set gIdentitySeen = oldSeen: Set gWindowKnown = oldKnown
        Set gIdentityLookupState = oldLookup
        Set gRevisionHeads = oldHeads
    End If
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, "MBR_MergeJournal", detail
End Sub

Private Sub MBR_ApplyOverlayOne(ByVal recoveryRoot As String, ByVal leaf As String, ByVal meta As Variant, _
                               ByVal accountRoot As String, ByVal storeId As String, _
                               Optional ByVal stageMirrorRaw As String = "")
    Dim path As String, n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    path = MBH_JoinPath(MBH_JoinPath(recoveryRoot, "overlays"), CStr(meta(0)))
    MBY_NoReparse path, True
    guard = MBH_GuardFile(path)
    If CDbl(gFSO.GetFile(path).Size) <> MB_ByteCount(CStr(meta(1))) Or MBH_HashFile(path) <> CStr(meta(2)) Then _
        Err.Raise MB_ERR + 1153, , "Captured pre-restore proof changed or is missing: " & path
    If LCase$(leaf) = "journal.tsv" Then
        MBR_MergeJournal path, storeId
    Else
        MBR_MergeProofFile leaf, path, accountRoot, storeId, stageMirrorRaw
    End If
    MBH_CloseGuard guard
    gRecoveryOverlayFiles = gRecoveryOverlayFiles + 1
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    Err.Raise n, "MBR_ApplyOverlayOne", detail
End Sub

Private Sub MBR_ApplyOverlay(ByVal recoveryRoot As String, ByVal files As Object, _
                            ByVal accountRoot As String, ByVal storeId As String)
    Dim key As Variant, meta As Variant, leaf As String, mirror As String, n As Long, detail As String
#If VBA7 Then
    Dim mirrorGuard As LongPtr
#Else
    Dim mirrorGuard As Long
#End If
    On Error GoTo Failed
    If files.Exists(MBR_MIRROR_LEAF) Then
        meta = MBM_Unpack(CStr(files(MBR_MIRROR_LEAF)), 3)
        mirror = MBH_JoinPath(MBH_JoinPath(recoveryRoot, "overlays"), CStr(meta(0)))
        MBY_NoReparse mirror, True: mirrorGuard = MBH_GuardFile(mirror)
        If CDbl(gFSO.GetFile(mirror).Size) <> MB_ByteCount(CStr(meta(1))) Or _
           MBH_HashFile(mirror) <> CStr(meta(2)) Then _
            Err.Raise MB_ERR + 1153, , "Captured independent stage proof changed or is missing."
    End If
    ' Ordered move history precedes proof consumers. These helpers only rebuild
    ' metadata; publication recovery and actual file moves happen after reload.
    If files.Exists("moves.tsv") Then
        meta = MBM_Unpack(CStr(files("moves.tsv")), 3)
        MBR_ApplyOverlayOne recoveryRoot, "moves.tsv", meta, accountRoot, storeId
    End If
    If files.Exists("journal.tsv") Then
        meta = MBM_Unpack(CStr(files("journal.tsv")), 3)
        MBR_ApplyOverlayOne recoveryRoot, "journal.tsv", meta, accountRoot, storeId
    End If
    For Each key In files.Keys
        leaf = LCase$(CStr(key))
        If leaf <> "moves.tsv" And leaf <> "journal.tsv" And leaf <> MBR_MIRROR_LEAF Then
            meta = MBM_Unpack(CStr(files(key)), 3)
            If leaf = "stages.tsv" Then
                MBR_ApplyOverlayOne recoveryRoot, CStr(key), meta, accountRoot, storeId, mirror
            Else
                MBR_ApplyOverlayOne recoveryRoot, CStr(key), meta, accountRoot, storeId
            End If
        End If
    Next key
    If Len(mirror) > 0 And Not files.Exists("stages.tsv") Then
        MBR_MergeProofFile "stages.tsv", mirror, accountRoot, storeId
    End If
    MBH_CloseGuard mirrorGuard
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard mirrorGuard
    On Error GoTo 0
    Err.Raise n, "MBR_ApplyOverlay", detail
End Sub

Private Function MBR_HasPreservedState() As Boolean
    Dim dir As Variant, item As Object, dirs As New Collection
    dirs.Add MB_StatePath("preserved")
    dirs.Add MBH_JoinPath(MBH_JoinPath(gRoot, "_state"), "preserved")
    If Len(mbrActiveRecoveryRoot) > 0 Then dirs.Add MBH_JoinPath(mbrActiveRecoveryRoot, "preserved")
    For Each dir In dirs
        If gFSO.FolderExists(CStr(dir)) Then
            MBY_NoReparse CStr(dir), True
            For Each item In gFSO.GetFolder(CStr(dir)).Files
                If LCase$(Right$(CStr(item.Name), 4)) = ".tsv" Then
                    MBR_HasPreservedState = True
                    Exit Function
                End If
            Next item
        End If
    Next dir
End Function

Private Sub MBR_LoadOverlayBase(ByVal recoveryRoot As String, ByVal accountRoot As String, ByVal storeId As String, _
                                ByVal manifest As String, ByVal expectedDigest As String, _
                                ByRef files As Object, ByRef chunks As Object)
    Dim reason As String, n As Long, detail As String
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    MBY_NoReparse manifest, True
    guard = MBH_GuardFile(manifest)
    If MBH_HashFile(manifest) <> expectedDigest Then _
        Err.Raise MB_ERR + 1149, , "Captured restore snapshot changed before its guarded read."
    If Not MBR_ReadSnapshot(manifest, accountRoot, storeId, files, chunks, reason) Then _
        Err.Raise MB_ERR + 854, , "Incomplete recovery needs its original snapshot: " & reason
    If Not MBR_BlocksValid(recoveryRoot, files, chunks, reason) Then _
        Err.Raise MB_ERR + 855, , "Incomplete recovery snapshot is damaged: " & reason
    MBH_CloseGuard guard
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    MBH_CloseGuard guard
    On Error GoTo 0
    Err.Raise n, "MBR_LoadOverlayBase", detail
End Sub


' ========== Preserve current non-index proof across snapshot restoration =====
' The caller guards the immutable pre-restore raw copy against its manifest.
' A conflict stops recovery with both originals preserved. No archive payload is
' read, moved or re-baselined here; only checked operating history is combined.
Private Sub MBR_MergeProofFile(ByVal leaf As String, ByVal priorRaw As String, _
                               ByVal accountRoot As String, ByVal storeId As String, _
                               Optional ByVal stageMirrorRaw As String = "")
    Dim target As String, stage As String, writer As Object, rows As Object, origins As Object, bad As Object
    Dim context As Variant, key As Variant, payload As String, f As Variant, old As Variant
    Dim reason As String, kind As String, n As Long, detail As String, i As Long
    On Error GoTo Failed
    If Not MBR_IsStateLeaf(leaf) Then Err.Raise MB_ERR + 1160, , "Unsafe recovery proof leaf."
    leaf = LCase$(leaf)
    If leaf = "journal.tsv" Then Err.Raise MB_ERR + 1160, , "Index history requires its separate semantic merger."
    target = MB_StatePath(leaf)
    kind = leaf
    If Left$(leaf, 9) = "baseline_" Then kind = "BASELINE"
    If Left$(leaf, 13) = "source_audit_" Then kind = "SOURCE"
    If Left$(leaf, 14) = "archive_audit_" Then kind = "ARCHIVE"
    If kind <> "moves.tsv" And kind <> "stages.tsv" And kind <> "account-roots.tsv" And _
       kind <> "BASELINE" And kind <> "SOURCE" And kind <> "ARCHIVE" Then
        If Not MBR_CheckedStateValid(priorRaw, reason) Then
            MBR_Log "WARNING", "RECOVERY_CURRENT_SETTINGS_RETAINED_RAW", leaf, priorRaw, 0, reason
            Exit Sub
        End If
        kind = "OTHER"
    End If
    stage = MBH_NewStagePath(target)
    Set writer = gFSO.CreateTextFile(stage, True, True)
    If kind = "moves.tsv" Then
        MBR_MergeMoveProof target, priorRaw, writer
    Else
        Set rows = CreateObject("Scripting.Dictionary"): rows.CompareMode = vbTextCompare
        Set origins = CreateObject("Scripting.Dictionary"): origins.CompareMode = vbTextCompare
        Set bad = CreateObject("Scripting.Dictionary"): bad.CompareMode = vbTextCompare
        If kind <> "OTHER" Then MBR_CollectProof target, leaf, kind, rows, origins, bad, context, 1, accountRoot, storeId
        MBR_CollectProof priorRaw, leaf, kind, rows, origins, bad, context, 2, accountRoot, storeId
        If kind = "stages.tsv" Then
            If Len(stageMirrorRaw) > 0 Then
                If StrComp(stageMirrorRaw, priorRaw, vbTextCompare) = 0 Then _
                    Err.Raise MB_ERR + 1161, , "One stage copy cannot claim two independent proof origins."
                MBR_CollectProof stageMirrorRaw, leaf, kind, rows, origins, bad, context, 4, accountRoot, storeId
            End If
            For Each key In bad.Keys
                f = MBM_Unpack(CStr(bad(key)), 3)
                If Not rows.Exists(CStr(f(0))) Then Err.Raise MB_ERR + 1161, , "Damaged stage has no independent complete proof."
                If CStr(rows(CStr(f(0)))) <> CStr(f(1)) Or _
                   (CLng(origins(CStr(f(0)))) And (7 Xor CLng(f(2)))) = 0 Then _
                    Err.Raise MB_ERR + 1161, , "Damaged stage payload is not exactly proved by another independent copy."
            Next key
        ElseIf kind = "account-roots.tsv" Then
            MBR_CheckMergedRoots rows, accountRoot, storeId
        End If
        If kind = "OTHER" Then
            ' Ordinary append logs retain their exact record order, independent
            ' of any Dictionary enumeration behavior.
            For i = 1 To rows.Count
                MB_WriteCheckedLine writer, CStr(rows(CStr(i)))
            Next i
        Else
            For Each key In rows.Keys
                MB_WriteCheckedLine writer, CStr(rows(key))
            Next key
        End If
        If kind = "SOURCE" Then
            If IsEmpty(context) Then Err.Raise MB_ERR + 1162, , "Source audit has no complete bound context."
            MB_WriteCheckedLine writer, MB_EncodeFields(Array("MBSOURCEAUDIT1", context(0), context(1), context(2), _
                "", "CURSOR", "", MBC_UtcStamp(MBC_UtcNow()), gRunId))
        ElseIf kind = "ARCHIVE" Then
            If IsEmpty(context) Then Err.Raise MB_ERR + 1162, , "Archive audit has no complete bound context."
            MB_WriteCheckedLine writer, MB_EncodeFields(Array("MBARCHIVEAUDIT1", context(0), context(1), context(2), _
                "CURSOR", "", "", "", "", "", "", "", "", "CURSOR", "", MBC_UtcStamp(MBC_UtcNow()), gRunId))
        End If
    End If
    writer.Close: Set writer = Nothing
    MBM_FlushPath stage
    MBR_PublishOverlayStage stage, target
    MBR_Log "INFO", "RECOVERY_CURRENT_PROOF_MERGED", leaf, target, 0, _
        "Pre-restore proof retained; unresolved audits restart with an empty cursor."
    Exit Sub
Failed:
    n = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not writer Is Nothing Then writer.Close
    If Len(stage) > 0 Then MBH_DeleteStage stage
    On Error GoTo 0
    Err.Raise n, "MBR_MergeProofFile", detail
End Sub

Private Function MBR_NextProofLine(ByVal stream As Object) As String
    If stream Is Nothing Then Exit Function
    Do Until stream.AtEndOfStream
        MBR_NextProofLine = stream.ReadLine
        If Len(MBR_NextProofLine) > 0 Then Exit Function
    Loop
    MBR_NextProofLine = ""
End Function

Private Sub MBR_MergeMoveProof(ByVal restored As String, ByVal currentRaw As String, ByVal writer As Object)
    Dim leftStream As Object, rightStream As Object, a As String, b As String, chosen As Variant
    Dim fa As Variant, fb As Variant, oldOps As Object, oldRedirects As Object, cached As Boolean
    Dim n As Long, detail As String, rowNumber As Long
#If VBA7 Then
    Dim leftGuard As LongPtr, rightGuard As LongPtr
#Else
    Dim leftGuard As Long, rightGuard As Long
#End If
    On Error GoTo Failed
    Set oldOps = mbmOps: Set oldRedirects = mbmRedirects: cached = True
    Set mbmOps = CreateObject("Scripting.Dictionary"): mbmOps.CompareMode = vbBinaryCompare
    Set mbmRedirects = CreateObject("Scripting.Dictionary"): mbmRedirects.CompareMode = vbTextCompare
    If gFSO.FileExists(restored) Then
        MBY_NoReparse restored, True: leftGuard = MBH_GuardFile(restored)
        If CDbl(gFSO.GetFile(restored).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 1163, , "Restored move history is truncated."
        Set leftStream = gFSO.OpenTextFile(restored, 1, False, -1)
    End If
    MBY_NoReparse currentRaw, True: rightGuard = MBH_GuardFile(currentRaw)
    If CDbl(gFSO.GetFile(currentRaw).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 1163, , "Current move history is truncated."
    Set rightStream = gFSO.OpenTextFile(currentRaw, 1, False, -1)
    Do
        a = MBR_NextProofLine(leftStream): b = MBR_NextProofLine(rightStream)
        If Len(a) = 0 And Len(b) = 0 Then Exit Do
        rowNumber = rowNumber + 1
        If rowNumber Mod 256 = 0 Then
            If Not MB_ContinueWork() Then Err.Raise MB_ERR + 1181, , "Recovery proof merge cancelled; original proofs are preserved."
        End If
        If Len(a) > 0 Then fa = MB_ReadCheckedLine(a, 11): a = MB_EncodeFields(fa)
        If Len(b) > 0 Then fb = MB_ReadCheckedLine(b, 11): b = MB_EncodeFields(fb)
        If Len(a) > 0 And Len(b) > 0 Then
            If a <> b Then Err.Raise MB_ERR + 1164, , "Move histories diverge; an unordered union would invent transitions."
        End If
        If Len(b) > 0 Then chosen = fb Else chosen = fa
        ' Replay exactly one longest ordered extension under isolated real caches.
        ' Do not call MBM_LoadMoves/MBM_Recover: merging never performs a move.
        MBM_ValidateMoveRow chosen
        MBM_CacheMove chosen
        MB_WriteCheckedLine writer, MB_EncodeFields(chosen)
    Loop
    GoTo Cleanup
Failed:
    n = Err.Number: detail = Err.Description
Cleanup:
    On Error Resume Next
    If Not leftStream Is Nothing Then leftStream.Close
    If Not rightStream Is Nothing Then rightStream.Close
    MBH_CloseGuard leftGuard: MBH_CloseGuard rightGuard
    If cached Then Set mbmOps = oldOps: Set mbmRedirects = oldRedirects
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, "MBR_MergeMoveProof", detail
    If gFatal Then Err.Raise MB_ERR + 1165, , "Move proof guard cleanup failed."
End Sub

Private Sub MBR_ProofContext(ByVal f As Variant, ByVal leaf As String, ByVal kind As String, ByRef context As Variant)
    Dim expectedLeaf As String, account As String, owner As String
    If StrComp(MBH_NormalizePath(CStr(f(1))), gRoot, vbTextCompare) <> 0 Or Len(CStr(f(2))) = 0 Then _
        Err.Raise MB_ERR + 1166, , "Proof belongs to another archive or has no store."
    account = MBH_NormalizePath(CStr(f(3))): owner = CStr(f(2))
    If MB_PathWithin(account, gRoot) Or MBL_IsFileSystemRoot(account) Then _
        Err.Raise MB_ERR + 1166, , "Invalid proof account binding."
    If kind = "BASELINE" Then
        expectedLeaf = "baseline_" & Left$(MBH_HashText(MB_EncodeFields(Array(gRoot, account, owner))), 32) & ".tsv"
    Else
        expectedLeaf = LCase$(kind) & "_audit_" & Left$(MBH_HashText(MB_EncodeFields(Array(owner, LCase$(account)))), 32) & ".tsv"
    End If
    If StrComp(leaf, expectedLeaf, vbTextCompare) <> 0 Then _
        Err.Raise MB_ERR + 1166, , "Proof filename disagrees with its store/account binding."
    If IsEmpty(context) Then
        context = Array(gRoot, owner, account)
    ElseIf StrComp(CStr(context(1)), owner, vbBinaryCompare) <> 0 Or _
           StrComp(CStr(context(2)), account, vbTextCompare) <> 0 Then
        Err.Raise MB_ERR + 1166, , "Proof copies belong to different account contexts."
    End If
End Sub

Private Function MBR_StageProofIdentity(ByVal f As Variant) As String
    Dim stage As String, final As String, held As String, phase As String, bytes As Double, p As String
    MBR_StageProofIdentity = MBSR_Identity(f)
    stage = MBH_NormalizePath(CStr(f(2))): final = MBH_NormalizePath(CStr(f(3)))
    p = CStr(gFSO.GetParentFolderName(final))
    final = MBY_CanonicalFilePath(final, p)
    stage = MBY_CanonicalFilePath(stage, p)
    If StrComp(CStr(gFSO.GetParentFolderName(stage)), p, vbTextCompare) <> 0 Or _
       Left$(LCase$(CStr(gFSO.GetFileName(stage))), 4) <> "~mb_" Or _
       StrComp(stage, final, vbTextCompare) = 0 Or MB_PathWithin(final, gRoot) Then _
        Err.Raise MB_ERR + 1167, , "Stage proof has unsafe output/stage paths."
    phase = CStr(f(4)): bytes = MB_ByteCount(CStr(f(6))): held = CStr(f(7))
    Select Case phase
        Case "RESERVE", "ABORTED"
            If Len(CStr(f(5))) > 0 Or bytes <> 0 Or Len(held) > 0 Then _
                Err.Raise MB_ERR + 1168, , "Unpublished empty phase contains contradictory bytes."
        Case "PREPARED", "PUBLISHED", "RECOVERED", "CANDIDATE", "HOLD_PREPARE", "HELD"
            If Not MB_IsHexKey(CStr(f(5))) Then Err.Raise MB_ERR + 1168, , "Stage proof has no expected digest."
            If phase = "HOLD_PREPARE" Or phase = "HELD" Then
                held = MBY_CanonicalFilePath(held, p)
                If StrComp(held, stage & ".held", vbTextCompare) <> 0 Then _
                    Err.Raise MB_ERR + 1168, , "Stage preservation suffix is invalid."
            ElseIf Len(held) > 0 Then
                Err.Raise MB_ERR + 1168, , "Publication phase unexpectedly includes a held path."
            End If
        Case Else
            Err.Raise MB_ERR + 1168, , "Unknown stage phase cannot authorize recovery."
    End Select
End Function

Private Sub MBR_ValidateAuditProof(ByVal f As Variant, ByVal kind As String, ByRef terminal As Boolean)
    Dim stamp As Date, ignored As Double, p As String
    If kind = "SOURCE" Then
        If CStr(f(0)) <> "MBSOURCEAUDIT1" Then Err.Raise MB_ERR + 1169, , "Wrong source-audit schema."
        stamp = MBC_ParseUtcStamp(CStr(f(7)))
        If Len(CStr(f(8))) = 0 Then Err.Raise MB_ERR + 1169, , "Source audit has no run identity."
        Select Case CStr(f(5))
            Case "CURSOR"
                If Len(CStr(f(4))) > 0 Then Err.Raise MB_ERR + 1169, , "Source cursor contains a message identity."
                If Len(CStr(f(6))) > 0 Then
                    If Not MB_IsHexKey(CStr(f(6))) Then Err.Raise MB_ERR + 1169, , "Invalid source cursor."
                End If
                terminal = True
            Case "REVIEW", "OK", "ABSENT"
                If Not MB_IsHexKey(CStr(f(4))) Then Err.Raise MB_ERR + 1169, , "Invalid source-audit identity."
            Case Else
                Err.Raise MB_ERR + 1169, , "Unknown source-audit status."
        End Select
    Else
        If CStr(f(0)) <> "MBARCHIVEAUDIT1" Then Err.Raise MB_ERR + 1170, , "Wrong archive-audit schema."
        stamp = MBC_ParseUtcStamp(CStr(f(15)))
        If Len(CStr(f(16))) = 0 Then Err.Raise MB_ERR + 1170, , "Archive audit has no run identity."
        If CStr(f(4)) = "CURSOR" Then
            If CStr(f(13)) <> "CURSOR" Then Err.Raise MB_ERR + 1170, , "Incomplete archive cursor."
            If Len(CStr(f(5))) > 0 Then
                If Not MB_IsHexKey(CStr(f(5))) Then Err.Raise MB_ERR + 1170, , "Invalid archive cursor."
            End If
            terminal = True
        ElseIf CStr(f(4)) = "RESULT" Then
            If CStr(f(13)) <> "REVIEW" And CStr(f(13)) <> "OK" Then _
                Err.Raise MB_ERR + 1170, , "Unknown archive-audit status."
            If Not MB_IsHexKey(CStr(f(5))) Or Not MB_IsHexKey(CStr(f(8))) Then _
                Err.Raise MB_ERR + 1170, , "Invalid archive expected digest/identity."
            ignored = MB_ByteCount(CStr(f(7)))
            p = MBY_CanonicalFilePath(CStr(f(6)), CStr(f(9)))
            If CStr(f(10)) <> "0" And CStr(f(10)) <> "1" Then Err.Raise MB_ERR + 1170, , "Invalid archived type flag."
            If MBA_ProofId(p, CStr(f(7)), CStr(f(8))) <> CStr(f(5)) Then _
                Err.Raise MB_ERR + 1170, , "Archive audit changed its expected-byte identity."
            Select Case CStr(f(12))
                Case "COMMITTED_INDEX", "MSG_ONLY_REBASELINED", "VERIFIED_PUBLICATION"
                Case Else
                    Err.Raise MB_ERR + 1170, , "Unknown archived proof provenance."
            End Select
        Else
            Err.Raise MB_ERR + 1170, , "Unknown archive-audit row type."
        End If
    End If
    If Not MBC_ValidDate(stamp) Then Err.Raise MB_ERR + 1171, , "Audit observation has no valid UTC timestamp."
End Sub

Private Sub MBR_CollectProof(ByVal path As String, ByVal leaf As String, ByVal kind As String, _
                              ByVal rows As Object, ByVal origins As Object, ByVal bad As Object, _
                              ByRef context As Variant, ByVal origin As Long, _
                              ByVal accountRoot As String, ByVal storeId As String)
    Dim reader As Object, line As String, f As Variant, prior As Variant, raw As Variant
    Dim payload As String, identity As String, key As String, p As String, stamp As Date
    Dim terminal As Boolean, count As Long, n As Long, detail As String, localBindings As Object, rowNumber As Long
#If VBA7 Then
    Dim guard As LongPtr
#Else
    Dim guard As Long
#End If
    On Error GoTo Failed
    If Not gFSO.FileExists(path) Then Exit Sub
    MBY_NoReparse path, True: guard = MBH_GuardFile(path)
    If CDbl(gFSO.GetFile(path).Size) Mod 2 <> 0 Then Err.Raise MB_ERR + 1172, , "Proof history is truncated."
    Set reader = gFSO.OpenTextFile(path, 1, False, -1)
    Set localBindings = CreateObject("Scripting.Dictionary"): localBindings.CompareMode = vbTextCompare
    Do Until reader.AtEndOfStream
        line = reader.ReadLine
        rowNumber = rowNumber + 1
        If rowNumber Mod 256 = 0 Then
            If Not MB_ContinueWork() Then Err.Raise MB_ERR + 1181, , "Recovery proof merge cancelled; original proofs are preserved."
        End If
        If Len(line) > 0 Then
            If terminal Then Err.Raise MB_ERR + 1173, , "Data follows an audit snapshot terminator."
            If kind = "stages.tsv" Then
                If Not MBSR_CheckedRecord(line, payload, identity) Then
                    If Not MBSR_UnprovenPayload(line, payload, identity) Then _
                        Err.Raise MB_ERR + 1161, , "Truncated/unknown stage cannot be merged safely."
                    key = CStr(origin) & "|" & identity
                    If bad.Exists(key) Then
                        prior = MBM_Unpack(CStr(bad(key)), 3)
                        If CStr(prior(1)) <> payload Then Err.Raise MB_ERR + 1174, , "Conflicting damaged stage payloads."
                    End If
                    bad(key) = MB_EncodeFields(Array(identity, payload, CStr(origin)))
                Else
                    f = MB_ReadCheckedLine(line, 10): identity = MBR_StageProofIdentity(f)
                    key = "STAGE_PATH|" & LCase$(MBH_NormalizePath(CStr(f(2))))
                    If origins.Exists(key) Then
                        If StrComp(CStr(origins(key)), MBH_NormalizePath(CStr(f(3))), vbTextCompare) <> 0 Then _
                            Err.Raise MB_ERR + 1174, , "A stage identity names conflicting destinations."
                    End If
                    origins(key) = MBH_NormalizePath(CStr(f(3)))
                    If rows.Exists(identity) Then
                        If CStr(rows(identity)) <> payload Then Err.Raise MB_ERR + 1174, , "Same stage/sequence has conflicting checked payloads."
                        origins(identity) = CLng(origins(identity)) Or origin
                    Else
                        rows(identity) = payload: origins(identity) = origin
                    End If
                End If
            Else
                Select Case kind
                    Case "BASELINE", "SOURCE": count = 9
                    Case "ARCHIVE": count = 17
                    Case "account-roots.tsv": count = 4
                    Case Else
                        raw = Split(line, vbTab): count = UBound(raw) - 1
                        If count < 1 Then Err.Raise MB_ERR + 1175, , "Unchecked state row."
                End Select
                f = MB_ReadCheckedLine(line, count): payload = MB_EncodeFields(f)
                Select Case kind
                    Case "BASELINE"
                        MBR_ProofContext f, leaf, kind, context
                        If CStr(f(0)) <> "MBBASE1" Or CStr(f(8)) <> "MSG_ONLY_REBASELINED" Or _
                           Not MB_IsHexKey(CStr(f(6))) Or MB_ByteCount(CStr(f(5))) <= 0 Then _
                            Err.Raise MB_ERR + 1176, , "Invalid immutable MSG baseline."
                        stamp = MBC_ParseUtcStamp(CStr(f(7)))
                        If Not MBC_ValidDate(stamp) Then Err.Raise MB_ERR + 1176, , "Baseline has no valid UTC capture time."
                        p = MBY_CanonicalFilePath(CStr(f(4)), CStr(f(3))): key = LCase$(p)
                        If rows.Exists(key) Then
                            prior = MBM_Unpack(CStr(rows(key)), 9)
                            If CStr(prior(6)) <> CStr(f(6)) Or MB_ByteCount(CStr(prior(5))) <> MB_ByteCount(CStr(f(5))) Then _
                                Err.Raise MB_ERR + 1177, , "Baseline path has conflicting historical expected bytes."
                        Else
                            rows(key) = payload
                        End If
                    Case "SOURCE", "ARCHIVE"
                        MBR_ProofContext f, leaf, kind, context
                        MBR_ValidateAuditProof f, kind, terminal
                        If kind = "SOURCE" Then
                            If CStr(f(5)) = "REVIEW" Then rows(CStr(f(4))) = payload
                        ElseIf CStr(f(13)) = "REVIEW" Then
                            key = CStr(f(5))
                            If rows.Exists(key) Then
                                prior = MBM_Unpack(CStr(rows(key)), 17)
                                If CStr(prior(12)) <> "COMMITTED_INDEX" Then rows(key) = payload
                            Else
                                rows(key) = payload
                            End If
                        End If
                        ' Unknown OK/ABSENT observations cannot remove REVIEW
                        ' obligations recovered from either compact snapshot.
                    Case "account-roots.tsv"
                        If CStr(f(0)) <> "MBAC3" Or Len(CStr(f(1))) = 0 Or _
                           StrComp(MBH_NormalizePath(CStr(f(3))), MBY_StateRoot(), vbTextCompare) <> 0 Then _
                            Err.Raise MB_ERR + 1178, , "Account-map binding is invalid."
                        key = LCase$(CStr(f(1))): p = MBH_NormalizePath(CStr(f(2)))
                        If Len(p) > 80 Or MB_PathWithin(p, MBY_StateRoot()) Or MBL_IsFileSystemRoot(p) Then _
                            Err.Raise MB_ERR + 1178, , "Account-map path is unsafe."
                        ' Per-file append history may legitimately rebind a store.
                        localBindings(key) = MB_EncodeFields(Array("MBAC3", key, p, MBY_StateRoot()))
                    Case Else
                        rows(CStr(rows.Count + 1)) = payload
                End Select
            End If
        End If
    Loop
    If kind = "SOURCE" Or kind = "ARCHIVE" Then
        If Not terminal Then Err.Raise MB_ERR + 1173, , "Audit snapshot has no terminal cursor."
    End If
    If kind = "account-roots.tsv" Then
        For Each raw In localBindings.Keys
            f = MBM_Unpack(CStr(localBindings(raw)), 4)
            If rows.Exists(CStr(raw)) Then
                prior = MBM_Unpack(CStr(rows(raw)), 4)
                If StrComp(CStr(prior(2)), CStr(f(2)), vbTextCompare) <> 0 Then
                    ' Current later binding must independently agree with its
                    ' selected owner or marker; conflicting unknown roots stop.
                    If StrComp(CStr(f(1)), storeId, vbTextCompare) <> 0 Or _
                       StrComp(CStr(f(2)), accountRoot, vbTextCompare) <> 0 Then MBY_CheckMarker CStr(f(2)), CStr(f(1)), True
                End If
            End If
            rows(CStr(raw)) = CStr(localBindings(raw))
        Next raw
    End If
    GoTo Cleanup
Failed:
    n = Err.Number: detail = Err.Description
Cleanup:
    On Error Resume Next
    If Not reader Is Nothing Then reader.Close
    MBH_CloseGuard guard
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, "MBR_CollectProof", detail
    If gFatal Then Err.Raise MB_ERR + 1179, , "Proof guard cleanup failed."
End Sub

Private Sub MBR_CheckMergedRoots(ByVal rows As Object, ByVal accountRoot As String, ByVal storeId As String)
    Dim a As Variant, b As Variant, fa As Variant, fb As Variant, path As String
    For Each a In rows.Keys
        fa = MBM_Unpack(CStr(rows(a)), 4): path = CStr(fa(2))
        MBY_NoReparse path, False
        If StrComp(CStr(fa(1)), storeId, vbTextCompare) = 0 Then
            If StrComp(path, accountRoot, vbTextCompare) <> 0 Then _
                Err.Raise MB_ERR + 1180, , "Recovered current-store mapping conflicts with selected account ownership."
        Else
            MBY_CheckMarker path, CStr(fa(1)), True
        End If
        For Each b In rows.Keys
            If StrComp(CStr(a), CStr(b), vbTextCompare) <> 0 Then
                fb = MBM_Unpack(CStr(rows(b)), 4)
                If MB_PathWithin(path, CStr(fb(2))) Or MB_PathWithin(CStr(fb(2)), path) Then _
                    Err.Raise MB_ERR + 1180, , "Recovered mappings overlap different stores."
            End If
        Next b
    Next a
End Sub
