/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='11', ARQATU='Rev11.sql';
/****************************************************************************
****************************************************************************/

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
Data do Lançamento
*/
--UPDATE FLAN SET DTVENCIMENTO = DTEMISSAO WHERE IDCONTA=1
--GO
--UPDATE FLAN SET DTBAIXA = DTEMISSAO WHERE IDCONTA=1
--GO



