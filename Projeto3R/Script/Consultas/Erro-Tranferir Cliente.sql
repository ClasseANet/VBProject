Use G3R;
declare @IDErro int;
declare @IDCerto int;

select @IDErro=797;
select @IDCerto=796
/*=====================================================================*/
UPDATE OEVENTOAGENDA SET IDCLIENTE=@IDCerto WHERE IDCLIENTE=@IDErro
UPDATE OATENDIMENTO SET IDCLIENTE=@IDCerto WHERE IDCLIENTE=@IDErro
UPDATE CVENDA SET IDCLIENTE=@IDCerto WHERE IDCLIENTE=@IDErro
