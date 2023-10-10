SELECT Source.Id, SamAccountName, Source.BatchNumber, Server, RunDate, CutoffDate
FROM Source
LEFT OUTER JOIN Batch
ON Source.BatchNumber = Batch.BatchNumber
WHERE Source.BatchNumber IN (141,142,143)