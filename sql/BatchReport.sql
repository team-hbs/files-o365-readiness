SELECT Source.Id, SamAccountName, Source.BatchNumber, ADHomeDirectory, OneDriveUrl, DestinationFolder, Validated, FileCountCrawl, FileSizeCrawl, ErrorCount, OfficeErrorCount,ApiCompletedCount, ApiNeedToMigrateCount, ApiErrorCount, ApiReportPath
FROM Source
LEFT OUTER JOIN MigrationJob
ON Source.Id = MigrationJob.OwnerId
WHERE BatchNumber IN

(
3010,3020,3030, 3040
)


ORDER BY BatchNumber ASC