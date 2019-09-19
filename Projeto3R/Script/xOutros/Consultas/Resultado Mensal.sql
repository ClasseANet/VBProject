USE G3R;

Select Month(A.DTATEND) [M�s], E.DSCSERVICO, COUNT(*) [Qtd.]
, case E.DSCSERVICO when 'Sess�o' then COUNT(*)*55 else 0 end [Valor]
From OSESSAO S 
Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO
Join OTPSERVICO E On S.IDTPSERVICO=E.IDTPSERVICO
Where month(A.DTATEND)=month(getdate())
And year(A.DTATEND)=year(getdate())
GROUP BY Month(A.DTATEND), E.DSCSERVICO
union
Select month(A.DTATEND) [M�s], S.NMPROD, Count(*) [Qtd.]
, COUNT(*)*25 [Valor]
From OATENDIMENTO_PRODUTO P
Join OATENDIMENTO A On P.IDLOJA=A.IDLOJA And P.IDATENDIMENTO=A.IDATENDIMENTO
Join SPRODUTO S On S.IDPROD=P.IDPROD
Where month(A.DTATEND)=month(getdate())
And year(A.DTATEND)=year(getdate())
Group by month(A.DTATEND), S.NMPROD

Select month(P.DTPGTO) [M�s],F.FORMAPGTO [Forma], SUM(P.VLPGTO) [Valor]
From CPGTOSVENDA P 
JOIN CFORMAPGTO F On P.IDFORMAPGTO=F.IDFORMAPGTO
Where month(P.DTPGTO)>=month(Getdate())
GROUP BY month(P.DTPGTO), F.FORMAPGTO
Compute Sum(SUM(P.VLPGTO))
