# Robo de Busca de Artigos Academicos

Aplicativo web em Streamlit para buscar artigos academicos por palavras-chave, limpar abstracts e gerar uma planilha Excel pronta para download.

## O que o app faz

- Busca artigos em fontes abertas:
  - OpenAlex
  - Semantic Scholar
  - Crossref
  - Europe PMC
  - DOAJ
  - arXiv
- Junta metadados quando o mesmo artigo aparece em mais de uma fonte.
- Mantem apenas registros com abstract valido.
- Gera uma planilha Excel no mesmo padrao do corpus de revisao bibliografica.
- Disponibiliza o arquivo `.xlsx` para baixar no navegador.
- Preenche um fluxo PRISMA editavel com base nos resultados da planilha.
- Permite baixar o PRISMA em PowerPoint editavel e SVG vetorial.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app_web_robo_artigos.py
```

## Publicar no Streamlit Community Cloud

1. Suba este repositorio para o GitHub.
2. Acesse `https://share.streamlit.io/`.
3. Clique em `New app`.
4. Selecione este repositorio.
5. Em `Main file path`, informe:

```text
app_web_robo_artigos.py
```

6. Clique em `Deploy`.

## Como o usuario usa

1. Informa ate 5 palavras-chave ou consultas.
2. Se precisar, usa o painel lateral para traduzir termos em portugues para ingles.
3. Opcionalmente informa o periodo por ano.
4. Clica em `Gerar planilha Excel pronta`.
5. Revisa/edita os numeros do PRISMA.
6. Baixa a planilha final e o PRISMA em alta qualidade.

Os parametros tecnicos de busca ficam configurados internamente para manter a experiencia simples.

## Observacoes

O app usa APIs e paginas publicas de bases academicas abertas. Alguns resultados podem variar conforme disponibilidade das fontes, limite de requisicoes e tempo de resposta de cada servico.
