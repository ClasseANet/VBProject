--select * from rbatida  where idfuncionario =7
--select * from rfuncionario
set dateformat 'dmy'


select distinct b.idbatida [ID0],b2.idbatida, f.nome, b.dtbatida [Entrada], b2.dtbatida [Saida]
, datediff(hh,  case when datepart(hh, b.dtbatida) >= 9 then  b.dtbatida else cast(convert(varchar(10),b.dtbatida, 103)+' 09:00' as datetime) end, b2.dtbatida) [HH]
, datepart(hh, b.dtbatida)
, case when datepart(hh, b.dtbatida) >= 9 then  b.dtbatida else cast(convert(varchar(10),b.dtbatida, 103)+' 09:00' as datetime) end
, (Select Max(T.HHFIM) From OATENDIMENTO T Where day(b.dtbatida)=day(t.dtatend) and month(b.dtbatida)=month(t.dtatend) and year(b.dtbatida)=year(t.dtatend))
from rbatida b 
Left join rbatida b2 on b.idloja=b2.idloja and b.idfuncionario=b2.idfuncionario 
Left join rfuncionario f on b.idloja=f.idloja and b.idfuncionario=f.idfuncionario 
where f.nome like '%cam%'
and month(b.dtbatida) =1
and year(b.dtbatida)= 2013
and b.dtbatida=(Select min(a.dtbatida) from rbatida a where a.idfuncionario=b.idfuncionario 
					and day(a.dtbatida)=day(b.dtbatida)
					and month(a.dtbatida)=month(b.dtbatida) 
					and year(a.dtbatida)=year(b.dtbatida)
				)
and b2.dtbatida=(Select max(a.dtbatida) from rbatida a where a.idfuncionario=b.idfuncionario 
					and day(a.dtbatida)=day(b.dtbatida)
					and month(a.dtbatida)=month(b.dtbatida) 
					and year(a.dtbatida)=year(b.dtbatida)
				)
order by ID0, b2.idbatida