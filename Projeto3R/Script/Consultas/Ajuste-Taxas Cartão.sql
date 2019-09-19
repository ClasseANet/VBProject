update FLAN set FLAN.VALOR=FLAN.VALOR*0.972
where FLAN.IDLAN in (select  f.idlan  
from flan f join cvenda v on f.idvenda=v.idvenda
join cpgtosvenda p on p.idvenda=v.idvenda
where f.idconta =2 
and f.tptransa <> 'T'
and p.idformapgto=3
and f.nparcela=1
and year(f.dtcadastro)=2012)
--select * from cformapgto
--0,972
--0,967