select s.idtpservico, s.idtptratamento, s.idarea, s.numsessao ,a.idcliente, count(*)
from osessao  s join oatendimento a on s.idloja=a.idloja and s.idatendimento=a.idatendimento
group by s.idtpservico, s.idtptratamento, s.idarea,  s.numsessao ,a.idcliente
order by 6 desc

--select idcliente, idatendimento from oatendimento where idcliente in (520,126,208,121,508)
--order by idcliente
