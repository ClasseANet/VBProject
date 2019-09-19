USE G3R;

Select convert(varchar, A.DTATEND,103) [Data], E.DSCSERVICO as [Produto], COUNT(*) [Qtd.]
, case E.DSCSERVICO when 'Sessão' then COUNT(*)*55 else 0 end [Valor]
From OSESSAO S 
Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO
Join OTPSERVICO E On S.IDTPSERVICO=E.IDTPSERVICO
Where A.DTATEND>=convert(varchar, Getdate(),103)
GROUP BY convert(varchar, A.DTATEND,103), E.DSCSERVICO
union 
Select convert(varchar, A.DTATEND,103) [Data -1], S.NMPROD as [Produto], Count(*) [Qtd.]
, COUNT(*)*25 [Valor]
From OATENDIMENTO_PRODUTO P
Join OATENDIMENTO A On P.IDLOJA=A.IDLOJA And P.IDATENDIMENTO=A.IDATENDIMENTO
Join SPRODUTO S On S.IDPROD=P.IDPROD
Where A.DTATEND>=convert(varchar, Getdate(),103)
Group by convert(varchar, A.DTATEND,103), S.NMPROD

Select (Select sum(VLVENDA) From CVENDA Where DTVENDA>=convert(varchar, Getdate(),103)) [VENDAS]
,(Select SUM(VLPGTO) [PAGAMENTOS] From CPGTOSVENDA Where DTPGTO>=convert(varchar, Getdate(),103)) [PAGAMENTOS]
,(Select Sum(VALOR) [LANÇAMENTOS] From FLAN Where DTEMISSAO>=convert(varchar, Getdate(),103) And FLGDELETE = 0)[LANÇAMENTOS]


Select F.FORMAPGTO [Forma], SUM(P.VLPGTO) [Valor]
From CPGTOSVENDA P 
JOIN CFORMAPGTO F On P.IDFORMAPGTO=F.IDFORMAPGTO
Where DTPGTO>=convert(varchar, Getdate(),103)
GROUP BY F.FORMAPGTO
Compute Sum(SUM(P.VLPGTO))

