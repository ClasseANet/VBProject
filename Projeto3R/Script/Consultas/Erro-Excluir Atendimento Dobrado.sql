Use G3R;
declare @IDErro int;
declare @IDCerto int;

select @IDErro=3601;
/*=====================================================================*/
DELETE OATENDIMENTO WHERE IDATENDIMENTO=@IDErro
DELETE OSESSAO WHERE IDATENDIMENTO=@IDErro

