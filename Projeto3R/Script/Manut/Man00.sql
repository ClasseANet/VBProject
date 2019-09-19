--DELETE FROM FLAN WHERE IDLAN IN (17,18,37,53,54,62,75,164,176,183,201,212)
select * from FLAN F 
where rtrim(cast(F.IDVENDA as char)) + '-' + rtrim(cast(F.IDPGTO as char))
in (
SELECT rtrim(cast(F2.IDVENDA as char)) + '-' + rtrim(cast(F2.IDPGTO as char))
FROM FLAN  F2
GROUP BY F2.IDVENDA, F2.IDPGTO
HAVING COUNT(*) >1
)
ORDER BY IDLAN


/***************
Tempo de Sessão
*/
update osessao set osessao.temposessao=(select (datediff(n,t.hhini, t.hhfim)*s.disparos)/(select sum(s2.disparos) from osessao s2 where s2.idatendimento=s.idatendimento group by s2.idatendimento) 
from osessao s join oarea a on s.idarea=a.idarea
join oatendimento t on s.idatendimento=t.idatendimento
where s.idatendimento=osessao.idatendimento
and s.idsessao=osessao.idsessao)
where osessao.disparos>0
/***************
