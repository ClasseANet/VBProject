select  a.dtatend [Data]
, t.dsctratamento [Tratamento], c.nome [Cliente]
--, f.nome [Funcionária] 
From oatendimento a
join ocliente c on a.idloja=c.idloja and a.idcliente=c.idcliente
join osessao s on a.idloja=s.idloja and a.idatendimento=s.idatendimento
--join rfuncionario f on a.idloja=f.idloja and a.idfuncionario=f.idfuncionario
join OTPTRATAMENTO t on s.idloja=t.idloja and s.idtptratamento=t.idtptratamento
Where  not t.IDTPTRATAMENTO in (1,2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,22,23,25,26)
--year(a.dtatend)=2017
--and month(a.dtatend)>0
 
group by a.idloja, a.idcliente, a.dtatend, c.nome, t.dsctratamento 
having a.dtatend= (select max(a2.dtatend) From oatendimento a2 Where a.idloja=a2.idloja and a.idcliente=a2.idcliente)
order by a.dtatend
