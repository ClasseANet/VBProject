select    E.* 
--,cast(StartDateTime) & month(StartDateTime)&day(StartDateTime), year(TimeStamp)&month(TimeStamp)&day(TimeStamp)
from OEVENTOAGENDA E 
where convert(varchar(10), TimeStamp,103)='28/02/2012'
and convert(varchar(10), StartDateTime,103)<>convert(varchar(10), TimeStamp,103)
ORDER BY StartDateTime

