
/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='37', ARQATU='Rev37.sql';
/****************************************************************************
****************************************************************************/

IF  EXISTS (SELECT * FROM sys.triggers WHERE object_id = OBJECT_ID(N'[dbo].[TU_CLIENTE]'))
DROP TRIGGER [dbo].[TU_CLIENTE]
go

/****** Object:  Trigger [dbo].[TU_CLIENTE]    Script Date: 04/07/2012 19:07:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[TU_CLIENTE] ON [dbo].[OEVENTOAGENDA] FOR UPDATE AS
  Declare @Sub nvarchar(255)
  Declare @SubName nvarchar(255)
  Declare @NewName nvarchar(255)
  Declare @Sessao int
  Declare @OldName nvarchar(255)
  Declare @OldIDCLI int
  Declare @NewIDCLI int
  Declare @Msg varchar(255)
  Declare @Pos int
  Declare @Tam int
  Declare @IDCLIAtend int

  If Update(IDCLIENTE) 
  Begin

    Select @Sessao=Inserted.IsMeeting, @NewIDCLI=Inserted.IDCLIENTE  From Inserted;
    Select @Sub=Deleted.Subject, @OldIDCLI=Deleted.IDCLIENTE From Deleted;
    
    --Select @NewName=C.NOME + ' - ' + Case When C.TEL2='' Then C.TEL1 When C.TEL2 Is Null Then C.TEL1 Else C.TEL2 End 
    Select @NewName=C.NOME From OCLIENTE C, Inserted Where C.IDCLIENTE=Inserted.IDCLIENTE;
	Select @OldName=C.NOME From OCLIENTE C, Deleted  Where C.IDCLIENTE=Deleted.IDCLIENTE;

    Set @Pos=CharIndex(' - ', @Sub)
    if @Pos<>0
    begin
      Set @SubName=RTrim(LTrim(SubString(@Sub, 1, @Pos-1)))
    End 
	Set @NewName=RTrim(LTrim(@NewName))
	--Set @Tam = iif(Len(@NewName) < Len(@SubName), Len(@NewName), Len(@SubName))
	If Len(@NewName) < Len(@SubName)
       Set @Tam = Len(@NewName)
    Else
       Set @Tam = Len(@SubName)
  

	Set @SubName=SubString(@SubName, 1, @Tam);
    Set @NewName=SubString(@NewName, 1, @Tam);
    
    Select @IDCLIAtend=IsNull(A.IDCLIENTE,0) From OATENDIMENTO A , Deleted  Where A.IDEVENTO=Deleted.IDEVENTO;

    --If( @Sessao=1 And @OldIDCLI<>0)
    If( @Sessao=1 And @OldIDCLI<>0 And (@IDCLIAtend=0 OR @IDCLIAtend<>@NewIDCLI))
	Begin
      If (@SubName<>@NewName And @OldIDCLI<>@NewIDCLI)
      Begin
		Set @Msg='Cliente Não pode ser Trocado'  + char(13)
        Set @Msg=@Msg + 'Título       : '+ @Sub  + char(13)
		Set @Msg=@Msg + 'Cliente Atual: '+ cast(@OldIDCLI as varchar(50)) + '-' + @OldName  + char(13)
		Set @Msg=@Msg + 'Cliente Novo : '+ cast(@NewIDCLI as varchar(50)) + '-' + @NewName
        RAISERROR (@Msg,16, 1)
        ROLLBACK TRANSACTION
      End
    End
  End
Go  
  
UPDATE OSESSAO SET IDMANIPULO =1 WHERE IDMANIPULO =0
GO

ALTER TABLE [dbo].[CITENSVENDA]  DROP CONSTRAINT [FK_CITENSVENDA_CVENDA]
GO 
ALTER TABLE [dbo].[CITENSVENDA]  WITH CHECK ADD  CONSTRAINT [FK_CITENSVENDA_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
ON UPDATE CASCADE
GO

ALTER TABLE [dbo].[CPGTOSVENDA]  DROP CONSTRAINT [FK_CPGTOSVENDA_CVENDA]
GO 
ALTER TABLE [dbo].[CPGTOSVENDA]  WITH CHECK ADD  CONSTRAINT [FK_CPGTOSVENDA_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
ON UPDATE CASCADE
GO

IF Not EXISTS (SELECT * FROM sys.columns WHERE name='IDFINGER' and object_id = OBJECT_ID(N'[dbo].[RFUNCIONARIO]'))
ALTER TABLE [dbo].[RFUNCIONARIO] ADD IDFINGER Int 
go

ALTER TABLE FITEMFATURA ALTER COLUMN IDPROD [decimal](10, 0)
GO

ALTER TABLE FITEMFATURA
	ADD CONSTRAINT  R_158 FOREIGN KEY (IDPROD) REFERENCES SPRODUTO(IDPROD)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

