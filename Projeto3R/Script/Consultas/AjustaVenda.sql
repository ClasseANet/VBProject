declare @Id as int;
set @Id=315

delete from oatendimento_venda where idvenda =@Id
delete from cpgtosvenda where idvenda =@Id
delete from frecibo where idvenda =@Id
delete from flan where idvenda =@Id
delete from citensvenda where idvenda =@Id
delete from cvenda where idvenda =@Id
delete from cvenda where idvenda =@Id
select * from cvenda where idvenda >=@Id-2
select * from frecibo