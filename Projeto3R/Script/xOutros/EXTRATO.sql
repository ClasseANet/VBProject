--set dateformat 'dmy';
--Select FF.NLAN,
Select FF.NDOC, FF.DTBAIXA, FF.DTEMISSAO, FF.DTVENCIMENTO, FF.IDLAN, Pagamento, Deposito
, (
	Select sum(Deposito)- sum(Pagamento)
	From (
	   Select Row_Number() Over (Order By DTBAIXA, DTEMISSAO, DTVENCIMENTO, IDLAN ) [NLAN], *
	   From (Select DTBAIXA, DTEMISSAO, DTVENCIMENTO, IDLAN
			   , [Pagamento]= Case FS.TPLAN WHEN 'D' THEN FS.VALOR ELSE 0 End
			   , [Deposito] = Case FS.TPLAN WHEN 'C' THEN FS.VALOR ELSE 0 End	   
			   From FLAN FS
			   Where FS.FLGDELETE=0 
			   And Not FS.DTBAIXA is Null 
			   And FS.IDLOJA = 1 
			   And FS.IDCONTA= 1 			   
			) TbRow 
		) TbSum
	Where TbSum.NLAN <= FF.NLAN
	 ) [Saldo] 

 From (
Select Row_Number() Over (Order By DTBAIXA, DTEMISSAO, DTVENCIMENTO, IDLAN ) [NLAN]
, FLAN2.NDOC, FLAN2.DTBAIXA, FLAN2.DTEMISSAO, FLAN2.DTVENCIMENTO, FLAN2.IDLAN
, Pagamento, Deposito
FROM (
		Select FS.NDOC, FS.DTBAIXA, FS.DTEMISSAO, FS.DTVENCIMENTO, FS.IDLAN
		, [Pagamento]= Case FS.TPLAN WHEN 'D' THEN FS.VALOR ELSE NULL End
		, [Deposito] = Case FS.TPLAN WHEN 'C' THEN FS.VALOR ELSE NULL End
		From FLAN FS
		Where FS.FLGDELETE=0 
		And Not FS.DTBAIXA is Null 
		And FS.IDLOJA = 1 
		And FS.IDCONTA= 1 
) FLAN2 
) FF order by FF.NLAN