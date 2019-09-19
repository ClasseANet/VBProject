SELECT C1.IDCLIENTE, C1.NOME, C1.TEL1, C1.TEL2, C2.IDCLIENTE, C2.NOME, C2.TEL1, C2.TEL2
FROM OCLIENTE C1 JOIN OCLIENTE C2 ON (replace(replace(replace(C1.TEL1, '-',''),'(',''),')','')=replace(replace(replace(C2.TEL1, '-',''),'(',''),')','') AND C1.TEL1<>'' AND C2.TEL1<>'')  OR (replace(replace(replace(C1.TEL1, '-',''),'(',''),')','')=replace(replace(replace(C2.TEL2, '-',''),'(',''),')','') AND C1.TEL1<>'' AND C2.TEL2<>'')
WHERE C1.IDCLIENTE<>C2.IDCLIENTE
and substring(C1.NOME,1,2)=substring(C2.NOME,1,2)

declare @IDCerto int;declare @IdErrado int
select @IDCerto=1012;select @IdErrado=1022

Update OEVENTOAGENDA SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)
Update OATENDIMENTO SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)
Update CVENDA SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)
Update OTAREFAEVT SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)

Update OCLASSE_CLIENTE SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)
Update OTRATAMENTOCLI SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)
Update OPENDENCIACLI SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)
Update SETORESCLI SET IDCLIENTE=@IDCerto WHERE IDCLIENTE IN (@IdErrado)

Delete OCLIENTE WHERE IDCLIENTE IN (@IdErrado)
/*
SELECT C1.IDCLIENTE, C1.NOME, C1.TEL1, C1.TEL2, C2.IDCLIENTE, C2.NOME, C2.TEL1, C2.TEL2
FROM OCLIENTE C1 JOIN OCLIENTE C2 ON C1.NOME=C2.NOME
WHERE C1.IDCLIENTE<>C2.IDCLIENTE
*/

--select o.name from syscolumns c join sysobjects o On o.id=c.id where c.name='IDCLIENTE'
--select * from sysobjects where id=110623437