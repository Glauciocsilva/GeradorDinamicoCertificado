# Gerador Dinâmico de Certificados em PDF

Este projeto tem como objetivo gerar certificados em PDF para participantes de um curso de forma dinâmica, utilizando um template em PPTX e uma lista de nomes em um arquivo Excel.

## Funcionalidades

- Manipulação de um template PPTX para incluir informações específicas do curso, local e data de emissão.
- Leitura de uma lista de participantes de um arquivo Excel.
- Geração de arquivos PPTX personalizados para cada participante.
- Conversão dos arquivos PPTX para PDF.
- Organização dos arquivos PDF em uma pasta específica e exclusão dos arquivos PPTX após a conversão.

## Pré-requisitos

- Python 3.x
- Bibliotecas Python: `pandas`, `pptx`, `pdfkit`
- Um arquivo Excel com a lista de nomes dos participantes
- Um template PPTX configurado corretamente

## Estrutura do Projeto

- `certificados.py`: Script principal para gerar os certificados.
- `template_certificado.pptx`: Template PPTX do certificado.
- `nomes_participantes.xlsx`: Arquivo Excel com a lista de nomes dos participantes.

## Observações Importantes
Certifique-se de que o template PPTX (template_certificado.pptx) esteja configurado corretamente conforme o desejado.
O script irá gerar uma pasta com os arquivos PPTX personalizados, convertê-los para PDF e, em seguida, excluir os arquivos PPTX, deixando apenas os PDFs na pasta especificada.

## Contato
Para mais informações, entre em contato:

Nome: Glaucio Cesar da Silva
Email: glauciocsilva@gmail.com
LinkedIn: https://www.linkedin.com/in/glauciocsilva/
