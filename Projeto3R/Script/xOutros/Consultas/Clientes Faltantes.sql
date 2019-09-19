select a.idcliente, c.nome, max(a.dtatend) [DTATEND], max(a.idatendimento) [IDATENDIMENTO]
, (select max(e.StartDateTime) 
	from oeventoagenda e 
	where e.idcliente=a.idcliente 
	group by e.idcliente) as [AGENDA]
from osessao s 
join oatendimento a on a.idatendimento=s.idatendimento
join ocliente c on a.idcliente=c.idcliente
where s.idtpservico<>1
and c.idcliente not in (2, 113, 116, 117)
group by a.idcliente, c.nome  
having month(max(a.dtatend))<=9 
--and month(max(a.dtatend))>7
and max(a.dtatend)>=(select max(e.StartDateTime) from oeventoagenda e where e.idcliente=a.idcliente group by e.idcliente) 
 order by  c.nome, [DTATEND]

