Use G3R;
declare @IdVenda int
declare @NewId int
select @IdVenda=999

--delete from FRECIBO				Where idvenda>=@IdVenda
--delete from FLAN				Where idvenda=@IdVenda
--delete from CPGTOSVENDA			Where idvenda=@IdVenda
--delete from OATENDIMENTO_VENDA	Where idvenda=@IdVenda
--delete from CITENSVENDA			Where idvenda=@IdVenda
--delete from CVENDA				Where idvenda=@IdVenda
Select * From CVENDA where idvenda> 433
--rollback
--commit
--begin transaction 
--update FLAN set idvenda=idvenda-1 Where idvenda>=435
--update CVENDA set idvenda=idvenda-1 Where idvenda>=435
--update CITENSVENDA set idvenda=idvenda-1 Where idvenda>=435
--update CPGTOSVENDA set idvenda=idvenda-1 Where idvenda>=435
--update FRECIBO set idvenda=idvenda-1 Where idvenda>=435
--update OATENDIMENTO_VENDA set idvend=idvenda-1 Where idvenda>=435
--delete from FLAN				Where idvenda=@IdVenda
--delete from CPGTOSVENDA			Where idvenda=@IdVenda
--delete from OATENDIMENTO_VENDA	Where idvenda=@IdVenda
--delete from CITENSVENDA			Where idvenda=@IdVenda
--delete from CVENDA				Where idvenda=@IdVenda
