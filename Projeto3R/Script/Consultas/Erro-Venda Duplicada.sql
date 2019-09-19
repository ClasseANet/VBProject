/*
select o.name from syscolumns c join sysobjects o On o.id=c.id where c.name='idvenda'

update cvenda set idvenda=idvenda-1 where idvenda>=551

delete From CITENSVENDA where  idvenda=998
delete From CPGTOSVENDA where  idvenda=998
delete From OATENDIMENTO_VENDA where  idvenda=998
delete From FRECIBO where  idvenda=998
delete From FLAN where  idvenda=998
delete From FNOTAFISCAL where  idvenda=998
delete From SMOVEST where  idvenda=998
delete From CCUPOM_VENDA where  idvenda=998

delete From cvenda where idvenda=998
*/
select * from FRECIBO where  idvenda=998
select * from CVENDA where  idvenda=550
select * from CITENSVENDA where  idvenda=550
select * from CPGTOSVENDA where  idvenda=550
select * from OATENDIMENTO_VENDA where  idvenda=550
select * from FLAN where  idvenda=550
select * from FNOTAFISCAL where  idvenda=550
select * from SMOVEST where  idvenda=550
select * from CCUPOM_VENDA where  idvenda=550
