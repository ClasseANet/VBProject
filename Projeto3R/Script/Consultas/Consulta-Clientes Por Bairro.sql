Select distinct cidade, bairro,  count(*) [Qtd.]
from ocliente 
--where ativo =1
group by cidade, bairro 
order by [Qtd.] desc , cidade, bairro 
--order by cidade, bairro 
--update OCLIENTE set cidade='Uberlandia' where cidade in ('Uberlândia', 'UBERLNADIA', 'UBERLANDDIA')
--update OCLIENTE set cidade='ITUIUTABA' where cidade ='ITUITABA'
--update OCLIENTE set bairro='Parque das Américas' where bairro ='Parque da Américas'
--update OCLIENTE set bairro='CANAÃ' where bairro ='CANAA'
--update OCLIENTE set bairro='MARACANÃ' where bairro ='MARACANA'
--update OCLIENTE set bairro='MORADA DA COLINA' where bairro ='MORADA COLINAA'
--update OCLIENTE set bairro='NOVA UBERLANDIA' where bairro ='NOVA UBERLÂNDIA'
--update OCLIENTE set bairro='OSVALDO' where bairro ='OSWALDO'
--update OCLIENTE set bairro='ROOSEVELT' where bairro in('ROOSELLT','ROOSVELT','ROUSSEVELT')
--update OCLIENTE set bairro='PACAEMBU' where bairro ='PACAMBU'
--update OCLIENTE set bairro='SANTA MARIA' where bairro ='STA MARIA'
--update OCLIENTE set bairro='SANTA MONICA' where bairro ='STA MONICA'
--update OCLIENTE set bairro='TALISMÃ' where bairro ='TALISMA'
--update OCLIENTE set bairro='VIGILATO PEREIRA' where bairro ='VIJILATO PEREIRA'


