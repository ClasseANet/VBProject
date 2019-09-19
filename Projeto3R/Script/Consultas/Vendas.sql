Use G3R;
declare @Mes int;declare @Ano int
declare @VlSessao int;
declare @ComProd int;declare @ComServ int
select @Ano=2012;select @Mes=5
select @VlSessao=55
select @ComProd=2;Select @ComServ=1
/*1====================================================================*/
Select Cast(@Mes as varchar)+'/'+Cast(@Ano as varchar) [Periodo]
/*2====================================================================*/
Select I.NMPROD [Produto], Sum(I.QTDVENDA) as [Qtd]
, Sum(I.QTDVENDA*I.VLUNIT)-case I.IDPROD when 1 then Sum(V.VLDESC) else 0 end
[Valor Real] 
, Comissão=case I.IDPROD when 1 then @ComServ*(Sum(I.QTDVENDA*I.VLUNIT)-Sum(V.VLDESC))/@VlSessao else @ComProd*Sum(I.QTDVENDA) end
From CITENSVENDA I Join CVENDA V On I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA
Join OCLIENTE C On V.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0
Where  Month(V.DTVENDA)=@Mes
And Year(V.DTVENDA)=@Ano
And C.ISENTO=0
Group by Year(V.DTVENDA), Month(V.DTVENDA), I.IDPROD, I.NMPROD
Order by Year(V.DTVENDA), Month(V.DTVENDA), I.NMPROD
/*3====================================================================*/
Select Count(distinct IDCLIENTE) [Clientes], SUM(Total) [Total Calc.]
, Sum(Comissão) [Comissão], Sum(Desconto) [Desconto]
, SUM(Total)-Sum(Comissão)-Sum(Desconto) [Total Real]
, SUM(Total)/Count(distinct IDCLIENTE) [Ticket]
From (Select Year(V.DTVENDA) [Ano], Month(V.DTVENDA) [Mes], I.NMPROD
		, V.IDCLIENTE, Sum(I.QTDVENDA*I.VLUNIT) Total
		, Comissão=case I.IDPROD when 1 then @ComServ*(Sum(I.QTDVENDA*I.VLUNIT)-Sum(V.VLDESC))/@VlSessao else @ComProd*Sum(I.QTDVENDA) end
		, Desconto=case I.IDPROD when 1 then Sum(V.VLDESC) else 0 end
	From CITENSVENDA I Join CVENDA V On I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA
	Join OCLIENTE C On V.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0
	Where  Month(V.DTVENDA)=@Mes
	And Year(V.DTVENDA)=@Ano
	Group by Year(V.DTVENDA), Month(V.DTVENDA), I.IDPROD, I.NMPROD, V.IDCLIENTE
	) as Temp
Group by Ano, Mes
/*4====================================================================*/
Select Count(*) [Sessões Prev.],  Count(*)*@VlSessao [Valor Previsto]
, Count(*)*@ComServ [Comissão]
From OEVENTOAGENDA E 
join OSERVICOEVT S On E.IDLOJA=S.IDLOJA and E.IDEVENTO=S.IDEVENTO
Join OCLIENTE C On E.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0
Where E.STARTDATETIME >= getdate()
And Month(E.STARTDATETIME) = @Mes
And Year(E.STARTDATETIME) = @Ano
And S.IDTPSERVICO <>1
And E.FLGCANCELADO<>1
Group By Year(E.STARTDATETIME), Month(E.STARTDATETIME)
Order by Year(E.STARTDATETIME), Month(E.STARTDATETIME)
/*5====================================================================*/
Select count(distinct IDCLIENTE) [Clientes], SUM(SESSAO) [Sessões]
, SUM(Valor) [Total Previsto], Sum(Comissão) [Comissão]
, SUM(Valor)/count(distinct IDCLIENTE) [Ticket Prev]
From (
Select  V.IDCLIENTE
, SESSAO=sum(case I.IDPROD when 1 then I.QTDVENDA else 0 end)
, Sum(I.QTDVENDA*I.VLUNIT)-SUM(V.VLDESC) Valor
, Comissão=case I.IDPROD when 1 then @ComServ*(Sum(I.QTDVENDA*I.VLUNIT)-SUM(V.VLDESC))/@VlSessao else @ComProd*Sum(I.QTDVENDA) end
From CITENSVENDA I Join CVENDA V On I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA
Join OCLIENTE C On V.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0
Where  Month(V.DTVENDA)=@Mes
And  Year(V.DTVENDA)=@Ano
Group by Year(V.DTVENDA), Month(V.DTVENDA), I.IDPROD, I.NMPROD, V.IDCLIENTE
--) as Temp
Union all
Select E.IDCLIENTE , Count(*) [SESSAO],  Count(*)*@VlSessao [Valor]
, Count(*)*@ComServ [Comissão]
From OEVENTOAGENDA E 
join OSERVICOEVT S On E.IDLOJA=S.IDLOJA and E.IDEVENTO=S.IDEVENTO
Join OCLIENTE C On E.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0
Where E.STARTDATETIME >= getdate()
And Month(E.STARTDATETIME) = @Mes
And Year(E.STARTDATETIME) = @Ano
And S.IDTPSERVICO <>1
And E.FLGCANCELADO<>1
Group By E.IDCLIENTE
) as Temp
