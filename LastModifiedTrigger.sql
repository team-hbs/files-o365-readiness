CREATE TRIGGER trg_Source_UpdateLastModified
ON dbo.Source
AFTER UPDATE
AS
DECLARE @PreHash nvarchar(256)
DECLARE @BatchNumber int
DECLARE @ADHomeDirectory nvarchar(256)
DECLARE @FileCount int
DECLARE @Hash nvarchar(256)
DECLARE @Status nvarchar(256)

SELECT @BatchNumber = BatchNumber FROM inserted
SELECT @ADHomeDirectory = ADHomeDirectory FROM inserted
SELECT @FileCount = FileCount FROM inserted
SELECT @Status = [Status] FROM inserted

SET @PreHash = CONCAT(@BatchNumber,@ADHomeDirectory,@FileCount,@Status)
SET @Hash = convert(varchar(256),HASHBYTES('SHA2_256',convert(varchar(256),@PreHash)),2)

UPDATE dbo.Source
SET LastModified = CURRENT_TIMESTAMP,
    [Hash] = @Hash,
	PreHash = @PreHash
WHERE Id IN (SELECT DISTINCT Id FROM inserted);