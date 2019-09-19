Select E.*  From OEVENTOAGENDA E
 Join OATENDIMENTO A ON E.IDLOJA=A.IDLOJA And E.IDEVENTO=A.IDEVENTO 
AND A.IDFUNCIONARIO=4
 Where (E.RecurrenceState = 0   Or E.RecurrenceState = 1) 
And Year(E.StartDateTime)  <= 2011 
And Month(E.StartDateTime) <= 9 
And Day(E.StartDateTime)   <= 19 
And Year(E.EndDateTime)  >= 2011 
And Month(E.EndDateTime) >= 9 
And Day(E.EndDateTime)   >= 19
