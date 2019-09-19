set dateformat 'dmy'

select c.idcliente, c.nome
, (Select Max(a.startdatetime) 
	from oeventoagenda a 
	where a.idcliente=c.idcliente) [DT]
from ocliente c 
where ativo =1 or ativo is null
order by 3