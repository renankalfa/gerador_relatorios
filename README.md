# Geradores Automáticos de Relatórios

Este repositório tem como objetivo apresentar duas soluções desenvolvidas por mim, no Ministério da Mulher, que envolvem a automatização da geração de arquivos Words a partir de planilhas.

 ![geradores de relatórios](https://user-images.githubusercontent.com/97196457/214443656-bcec2f68-64d8-4cf5-be3f-55579bf74707.png)

1. **Gerador de Subsídios**: programa com interface gráfica que gera um documento Word de uma ou várias unidades federativas;
2. **Gerador de Relatórios de Feminicídios**: script Python que gera relatórios com gráficos e métricas de feminicídios para todas as unidades federativas.

#

## Gerador de Subsídios
**Problema de negócio**: demanda semanal/diária de um arquivo Word com os dados das parcerias vigentes sob responsabilidades do departamento.

**Como era feita**: manualmente era copiado e colado para um arquivo Word informações de cada parceira de uma determinada UF.

**Solução criada**: programa com interface gráfica que gera automaticamente e de maneira rápida esse arquivo.

![gerador](https://user-images.githubusercontent.com/97196457/214552842-c86079a3-559f-49ae-891b-7049b30fe83a.png)

Exemplo de arquivo gerado:

![documento gerado](https://user-images.githubusercontent.com/97196457/214554411-e267bf59-5bf9-4996-bcdc-bebbd1fafe90.png)

#### Ferramentas utilizadas

- Jupyter Notebook/PyCharm: locais para escrita e testes do script Python.
- Python (versão 3.9.6): linguagem de programação usada para escrita dos scripts.
- Bibliotecas Python
   - Tkinter (versão 8.5): criação da interface gráfica.
   - Docx (versão 0.8.11): criação do arquivo Word.
   - Pandas (versão 1.5.3): carregamento e manipulação de dados.

#

## Gerador de Relatórios de Feminicídios
**Problema de negócio**: falta de um material físico com dados sobre a violência contra a mulher para viagens e consulta em reuniões.

**Solução criada**: geração de relatórios de feminicídios para todas as UFs por meio de um script Python.

Os arquivos gerados podem ser encontrados [clicando aqui](https://drive.google.com/drive/folders/10Somv9JTl1QycLr0KsT_6X9L8DT40FDf?usp=sharing). A imagem abaixo retrata um exemplo de arquivo gerado:

![relatorio feminicidio](https://user-images.githubusercontent.com/97196457/214555699-b390133b-e43b-4d46-9592-2991729be9ac.png)

#### Ferramentas utilizadas

- Jupyter Notebook/PyCharm: locais para escrita e testes do script Python.
- Python (versão 3.9.6): linguagem de programação usada para escrita dos scripts.
- Bibliotecas Python
    - Docx (versão 0.8.11): criação do arquivo Word.
    - Pandas (versão 1.5.3): carregamento e manipulação de dados.

#

## Próximos passos

- Refatoração dos códigos;
- Automação da geração de subsídios (sem a necessidade de intervenção humana).

#

<a href="#top">Back to top</a>
