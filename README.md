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

## Configuracao recomendada

- `Artigos por consulta`: 25
- `Agentes paralelos`: 8
- `Limite final`: 500
- `Tempo limite`: 20 minutos

Para buscas muito amplas, reduza o limite final ou nao marque `Somar redes predefinidas`.

## Observacoes

O app usa APIs e paginas publicas de bases academicas abertas. Alguns resultados podem variar conforme disponibilidade das fontes, limite de requisicoes e tempo de resposta de cada servico.
