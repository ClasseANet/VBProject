select a.idatendimento [Id.], a.dtatend [Data], c.nome [Cliente], count(distinct s.idsessao) [Áreas]
, datediff("mi", a.hhini, a.hhfim) [Tempo Sessão(Min)], f.nome [Funcionária] 
From oatendimento a
join ocliente c on a.idloja=c.idloja and a.idcliente=c.idcliente
join rfuncionario f on a.idloja=f.idloja and a.idfuncionario=f.idfuncionario
join osessao s on a.idloja=s.idloja and a.idatendimento=s.idatendimento
Where year(a.dtatend)=2012
and month(a.dtatend)=10
group by a.idatendimento, a.dtatend, c.nome,a.hhini, a.hhfim, f.nome 
order by f.nome
