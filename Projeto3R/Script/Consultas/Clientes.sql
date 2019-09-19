Use G3R;
declare @Mes int;declare @Ano int
declare @VlSessao int;
declare @ComProd int;declare @ComServ int
select @Ano=2011;select @Mes=10
select @VlSessao=55
select @ComProd=2;Select @ComServ=1

--Select IDCLIENTE, BAIRRO, DTCADASTRO, ATIVO  FROM OCLIENTE O

Select Year(DTCADASTRO) Ano, Month(DTCADASTRO) Mes
, Count(Distinct O.IDCLIENTE) [Total], Sum(O.ATIVO) [Ativos]
FROM OCLIENTE O 
Group By Year(DTCADASTRO), Month(DTCADASTRO)

Select Distinct Count(Distinct O.IDCLIENTE) [Total], Sum(O.ATIVO) [Ativos]
, (Select Count(Distinct E.IDCLIENTE)
FROM OCLIENTE O2 Left Join OEVENTOAGENDA E On O2.IDLOJA=E.IDLOJA And O2.IDCLIENTE=E.IDCLIENTE) [Com Evento]
, (Select Count(Distinct V.IDCLIENTE)
FROM OCLIENTE O2 Left Join CVENDA V On O2.IDLOJA=V.IDLOJA And O2.IDCLIENTE=V.IDCLIENTE) [Com Venda]
FROM OCLIENTE O 

Select 
--Cast(Year(V.DTVENDA) as varchar)+'/'+Cast(Month(V.DTVENDA) as varchar) [Mes]
--, 
Count(Distinct C.IDCLIENTE) [Clientes], cast(C.IDTPCONHEC as varchar)+'-'+T.NMCONHEC [Conhec.], sum(V.VLVENDA-V.VLDESC) [Valor]
--, C.DSCTPCONHEC
From CVENDA V Join OCLIENTE C On C.IDLOJA=V.IDLOJA And C.IDCLIENTE=V.IDCLIENTE
Join OTPCONHEC T On T.IDTPCONHEC=C.IDTPCONHEC 
Group By 
--Year(V.DTVENDA), Month(V.DTVENDA),
 C.IDTPCONHEC, T.NMCONHEC
--, C.DSCTPCONHEC
Order By 1 Desc

--Select * from OCLIENTE Where IDTPCONHEC=0
--Select C.* from OCLIENTE C Join CVENDA V On C.IDLOJA=V.IDLOJA And C.IDCLIENTE=V.IDCLIENTE Where C.IDTPCONHEC=10








