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
2. Opcionalmente informa o periodo por ano.
3. Clica em `Gerar planilha Excel pronta`.
4. Baixa a planilha final.

Os parametros tecnicos de busca ficam configurados internamente para manter a experiencia simples.

## Observacoes

O app usa APIs e paginas publicas de bases academicas abertas. Alguns resultados podem variar conforme disponibilidade das fontes, limite de requisicoes e tempo de resposta de cada servico.
