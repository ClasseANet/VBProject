SELECT O.NAME [TABELA], C.NAME [COLUNA], C.COLORDER 
FROM SYSCOLUMNS C JOIN SYSOBJECTS O ON C.ID=O.ID
WHERE O.NAME IN ('OSALA')
--AND C.NAME = 'OBS' --('OTPMAQ','OMAQUINA','OTPMANUT','OMAQMANUT','OTPMANIPULO','OMANIPULO','OLAMPADA','OMAQDISPAROS','OAREA','OTPSERVICO','OTPTRATAMENTO','OAREA_TRATAMENTO','ODIRECAO','OREACAOADV')
ORDER BY O.NAME, C.COLORDER 