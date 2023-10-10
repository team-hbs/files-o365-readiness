SELECT Source.Id, SamAccountName, Source.BatchNumber, ADHomeDirectory, OneDriveUrl,Validated, SiteOwner, FileCountCrawl, FileSizeCrawl, ErrorCount, PathLengthCount, OfficeErrorCount, NoAccessCount
FROM Source
LEFT OUTER JOIN ScanJob
ON Source.Id = ScanJob.OwnerId
WHERE Source.BatchNumber in (
141,142,143
)

ORDER BY BatchNumber ASC