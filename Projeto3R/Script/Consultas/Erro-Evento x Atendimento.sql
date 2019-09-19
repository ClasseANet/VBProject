select e.idevento, e.idcliente, a.idatendimento, a.idcliente
from oeventoagenda e join oatendimento a on e.idevento=a.idevento
where e.idcliente<>a.idcliente
order by startdatetime

--select * from osessao where idatendimento=728
--update osessao set numsessao=2 where idatendimento=728;
--select * from oatendimento where idatendimento=358
--select * from oatendimento_venda where idatendimento=728
--select * from oeventoagenda where idevento =1127
--select * from oeventoagenda where idevento =310
--select * from ocliente where idcliente = 224
--select * from cvenda where idvenda=505
--select * from frecibo where idvenda=505
--update oeventoagenda set idcliente=224 where idevento=693;
--update oatendimento set idcliente=131 where idatendimento=728;
--update cvenda set idcliente=131 where idvenda=505;
--update oeventoagenda set subject='Nicole Velasco - 22-8136-85661' where idevento=1127;

/*
select * 
from oatendimento a
where a.dtatend>'2011-08-14'
and a.dtatend<'2011-08-16'
order by a.dtatend

update oeventoagenda set idcliente=74 where idevento=704;
update oeventoagenda set idcliente=227 where idevento=624;
update oeventoagenda set idcliente=61 where idevento=617;
update oeventoagenda set idcliente=210 where idevento=584;
update oeventoagenda set idcliente=246 where idevento=669;
update oeventoagenda set idcliente=11 where idevento=326;
update oeventoagenda set idcliente=203 where idevento=505;
update oeventoagenda set idcliente=204 where idevento=506;
update oeventoagenda set idcliente=143 where idevento=572;

*/