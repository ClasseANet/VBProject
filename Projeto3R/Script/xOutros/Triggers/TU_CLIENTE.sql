IF EXISTS (SELECT name FROM sysobjects WHERE name = 'TU_CLIENTE' AND type = 'TR')  DROP TRIGGER TU_CLIENTE
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create TRIGGER TU_CLIENTE ON OEVENTOAGENDA FOR UPDATE AS
  Declare @Sub nvarchar(255)
  Declare @NewSub nvarchar(255)
  Declare @Sessao int
  Declare @IDCLI int
  Declare @NewIDCLI int
  Declare @Msg as varchar(255)
  Declare @Pos as integer

  If Update(IDCLIENTE) 
  Begin

    Select @Sessao=Inserted.IsMeeting, @NewIDCLI=Inserted.IDCLIENTE  From Inserted
    Select @Sub=Deleted.Subject, @IDCLI=Deleted.IDCLIENTE From Deleted
    
    --Select @NewSub=C.NOME + ' - ' + Case When C.TEL2='' Then C.TEL1 When C.TEL2 Is Null Then C.TEL1 Else C.TEL2 End 
    Select @NewSub=C.NOME 
    From OCLIENTE C, Inserted 
    Where C.IDCLIENTE=Inserted.IDCLIENTE

    Set @Pos=CharIndex(' - ', @Sub)
    if @Pos<>0
    begin
      Set @Sub=SubString(@Sub, 1, @Pos-1)
    End 

    If( @Sessao=1 And @IDCLI<>0)
	Begin
      If (@Sub<>@NewSub And @IDCLI<>@NewIDCLI)
      Begin
		Set @Msg='Cliente Não pode ser Trocado'  + char(13)
        Set @Msg=@Msg + 'Atual: ' + cast(@IDCLI as varchar(50)) + '-' + @Sub  + char(13)
		Set @Msg=@Msg + 'Novo: ' + cast(@NewIDCLI as varchar(50))+ '-' + @NewSub
        RAISERROR (@Msg,16, 1)
        ROLLBACK TRANSACTION
      End
    End
  End
 Go
