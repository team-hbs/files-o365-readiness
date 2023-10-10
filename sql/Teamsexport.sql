SELECT SourceTeamName, CONCAT(',','''',SourceSiteUrl, ''''), SourceEmail, SourceTeamName,SourceTeamId, CONCAT(',','''',SourceTeamId, '''')
FROM Source
LEFT OUTER JOIN MigrationJob
ON Source.Id = MigrationJob.OwnerId
WHERE BatchNumber IN

(
3010,3020,3030, 3040
)
AND MigrationType = 'TeamsToTeams'



