--**********************
-- Clientes sem Eventos
Select C.IDCLIENTE, C.NOME, E.IDEVENTO 
from OCLIENTE C LEFT JOIN OEVENTOAGENDA E ON E.IDCLIENTE =C.IDCLIENTE
WHERE E.IDEVENTO IS NULL
ORDER BY C.IDCLIENTE

--**********************
-- Clientes Sem Atendimentos e sem Eventos a serem realizados
Select DISTINCT C.IDCLIENTE, c.NOME
from OCLIENTE C LEFT JOIN OATENDIMENTO A ON A.IDCLIENTE=C.IDCLIENTE
WHERE A.IDATENDIMENTO IS NULL
AND C.IDCLIENTE NOT IN (
					Select DISTINCT C2.IDCLIENTE
from (OCLIENTE C2 JOIN OEVENTOAGENDA E2 ON C2.IDCLIENTE=E2.IDCLIENTE AND E2.FLGCANCELADO=0)
LEFT JOIN OATENDIMENTO A2 ON A2.IDCLIENTE=C2.IDCLIENTE
WHERE A2.IDATENDIMENTO IS NULL)
--and C.ATIVO=1
select * from ocliente where ativo=1
/*
Update OCLIENTE set ATIVO=0
Where OCLIENTE.IDCLIENTE in (Select DISTINCT C.IDCLIENTE
from OCLIENTE C LEFT JOIN OATENDIMENTO A ON A.IDCLIENTE=C.IDCLIENTE
WHERE A.IDATENDIMENTO IS NULL
AND C.IDCLIENTE NOT IN (
					Select DISTINCT C2.IDCLIENTE
from (OCLIENTE C2 JOIN OEVENTOAGENDA E2 ON C2.IDCLIENTE=E2.IDCLIENTE AND E2.FLGCANCELADO=0)
LEFT JOIN OATENDIMENTO A2 ON A2.IDCLIENTE=C2.IDCLIENTE
WHERE A2.IDATENDIMENTO IS NULL)
and C.ATIVO=1)

*/