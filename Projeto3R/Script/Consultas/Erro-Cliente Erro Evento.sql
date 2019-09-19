SELECT distinct E.IDEVENTO, c.idcliente, C.NOME, E.SUBJECT, E.STARTDATETIME , e.flgcancelado
FROM OEVENTOAGENDA E JOIN OCLIENTE C ON C.IDCLIENTE=E.IDCLIENTE
WHERE e.ISMEETING = 1
AND SUBSTRING(C.NOME,1,5) <>  substring(e.SUBJECT,1,5)  
ORDER BY   STARTDATETIME

SELECT distinct E.IDEVENTO, c.idcliente, C.NOME, E.SUBJECT, c2.idcliente, E.STARTDATETIME , e.flgcancelado
FROM OEVENTOAGENDA E JOIN OCLIENTE C ON C.IDCLIENTE=E.IDCLIENTE
join ocliente c2 on  SUBSTRING(C2.NOME,1,12)=substring(e.SUBJECT,1,12)
WHERE e.ISMEETING = 1
AND SUBSTRING(C.NOME,1,5) <>  substring(e.SUBJECT,1,5)  
--and e.flgcancelado =1
ORDER BY   STARTDATETIME


--and e.flgcancelado =1
--update oeventoagenda set idcliente =681 where idevento=2775

--select idcliente, * from oeventoagenda where idevento =1711
--select idcliente, * from oatendimento where idevento =1711
--select idcliente, * from ocliente  where idcliente in (266, 317)


--BEGIN TRANSACTION
--update oeventoagenda set subject = (Select NOME + ' - ' + CASE WHEN TEL1='' THEN TEL2 ELSE TEL1 END
--From OCLIENTE where OCLIENTE.IDCLIENTE=oeventoagenda.IDCLIENTE)
--WHERE  IDEVENTO IN (441, 1667)
--and oeventoagenda.ISMEETING = 1
--and oeventoagenda.idcliente>0

--rollback
--COMMIT
--select * from oeventoagenda