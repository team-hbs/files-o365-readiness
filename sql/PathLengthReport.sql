SELECT *
FROM ScanFile,Source
WHERE  Source.BatchNumber In (170)
AND PathLength > 255
AND Source.Id = ScanFile.OwnerId
ORDER BY PathLength ASC