select e.startdatetime, a.dscarea, c.nome
from oeventoagenda e 
join oservicoevt s on s.idevento=e.idevento
join oarea a on s.idarea=a.idarea
join ocliente c on c.idcliente=e.idcliente
where day(e.startdatetime) = 17
and month(e.startdatetime) = 11
and year(e.startdatetime) = 2011
and e.flgcancelado = 0
order by c.nome, e.startdatetime

