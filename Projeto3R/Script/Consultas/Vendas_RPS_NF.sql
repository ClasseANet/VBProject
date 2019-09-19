select R.IDRECIBO, '' [NF], v.idvenda, v.dtvenda, '' [Data]
, v.vlvenda-v.vldesc [valor], '' [QTD], '' [VL], c.registro, c.nome 
, c.cep, c.endereço, c.email,  c.tel1, c.tel2 
, c.bairro, c.cidade, c.estado
from cvenda v join ocliente c on c.idcliente=v.idcliente
Left Join FRECIBO R On R.IDLOJA=V.IDLOJA And R.IDVENDA=V.IDVENDA
where month(dtvenda)=9
order by dtvenda


--SELECT * FROM frecibo