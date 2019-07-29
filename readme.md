# Gerador de Certificados

Esta é uma aplicação em python que visa facilitar a geração de certificados. 

A aplicação pressupõe que o usuário possui um modelo de certificado (.docx) 
e uma lista de informações (.csv), sendo que para cada linha da lista será 
gerado um certificado.

Para que isso seja possível, cada coluna da linha deve ter um campo correspondente 
no modelo. Esses campos correspondentes são representados pelos termos Valor_1, 
Valor_2, ... , Valor_n, para n colunas de cada linha de informação.
