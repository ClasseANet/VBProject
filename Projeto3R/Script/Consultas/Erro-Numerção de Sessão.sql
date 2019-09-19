
select a.idcliente, s.idarea, s.numsessao, s.idtpservico , count(*)
from osessao s 
join oatendimento a on a.idatendimento=s.idatendimento
where s.idtpservico >1
group by a.idcliente, s.idarea, s.numsessao, s.idtpservico 
having count(*)>1
order by a.idcliente, s.idarea, s.numsessao 

--select idatendimento from oatendimento where idcliente=447
select * from osessao where idatendimento in (1020,
1034,
1255)
order by idarea , idatendimento