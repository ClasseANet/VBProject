SELECT C1.IDCLIENTE, C1.NOME, C1.TEL1, C1.TEL2, C2.IDCLIENTE, C2.NOME, C2.TEL1, C2.TEL2
FROM OCLIENTE C1 JOIN OCLIENTE C2 ON (replace(replace(replace(C1.TEL1, '-',''),'(',''),')','')=replace(replace(replace(C2.TEL1, '-',''),'(',''),')','') AND C1.TEL1<>'' AND C2.TEL1<>'')  OR (replace(replace(replace(C1.TEL1, '-',''),'(',''),')','')=replace(replace(replace(C2.TEL2, '-',''),'(',''),')','') AND C1.TEL1<>'' AND C2.TEL2<>'')
WHERE C1.IDCLIENTE<>C2.IDCLIENTE
and substring(C1.NOME,1,2)=substring(C2.NOME,1,2)

/*
Declare @IdCerto int;declare @IDErrado int
select @IdCerto=365
Select @IDErrado=564

Update OEVENTOAGENDA SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)
Update OATENDIMENTO SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)
Update CVENDA SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)
Update OTAREFAEVT SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)

Update OCLASSE_CLIENTE SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)
Update OTRATAMENTOCLI SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)
Update OPENDENCIACLI SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)
Update SETORESCLI SET IDCLIENTE=@IdCerto WHERE IDCLIENTE IN (@IDErrado)

Delete OCLIENTE WHERE IDCLIENTE IN (@IDErrado)
*/


/*
SELECT C1.IDCLIENTE, C1.NOME, C1.TEL1, C1.TEL2, C2.IDCLIENTE, C2.NOME, C2.TEL1, C2.TEL2
FROM OCLIENTE C1 JOIN OCLIENTE C2 ON C1.NOME=C2.NOME
WHERE C1.IDCLIENTE<>C2.IDCLIENTE
*/

--select o.name from syscolumns c join sysobjects o On o.id=c.id where c.name='IDCLIENTE'
--select * from sysobjects where id=110623437