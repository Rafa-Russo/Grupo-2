# Relatório de progresso — Notebooks 03, 04 e 05
## Projeto: crime e preço de imóveis na cidade de São Paulo

> **Como usar este documento.** Ele foi escrito para ser lido isoladamente (por exemplo, dentro do NotebookLM). Não é preciso abrir os notebooks para entendê-lo. Os gráficos dos notebooks não podem ser lidos como texto, então cada figura relevante é descrita em palavras. Quando um número vem do texto de um notebook e não de uma saída executada, isso está dito. A **seção 8** lista as divergências encontradas entre o texto explicativo e as saídas executadas. Leia-a antes de citar qualquer número.

---

## 1. Resumo executivo

**Pergunta de pesquisa.** Bairros mais violentos têm metro quadrado mais barato em São Paulo?

**Resposta curta.** Sim, mas não pelo motivo intuitivo. O que se associa ao preço é a **composição** do crime (a fração das ocorrências que é violenta), e não a **quantidade** de crime.

| Achado | Notebook | Evidência principal |
|---|---|---|
| Fração de crimes violentos no bairro tem efeito negativo e significativo sobre o R$/m². | 03 | Coeficiente OLS −0,922 (p ≈ 7×10⁻¹⁶). +10 pontos percentuais de crimes violentos equivalem a −8,8% no R$/m². |
| A quantidade de crimes violentos não tem efeito detectável. | 03 | Coeficiente −0,029, p = 0,16, intervalo de confiança cruza o zero. |
| Incluir crime melhora a previsão de preço. | 03 | R² no teste sobe de 0,377 para 0,451 (+0,074). O RMSE cai 6,15%. |
| O crime captura 97% do ganho que a identidade do bairro capturaria. | 03 | O crime é, em grande parte, um marcador de localização. **O resultado é preditivo e descritivo, não causal.** |
| Os bairros se organizam em 3 perfis estáveis, mas pouco separados. | 04 | Silhueta 0,276 (estrutura fraca); estabilidade de ~97% sob reamostragem. |
| Dos quatro "tipos" imaginados de bairro (caro/barato × muita/pouca violência), só "caro + pouca violência" é um tipo real. | 04 | 86% do cluster 2 (k=4) cai nesse quadrante. "Caro + muita violência" não forma cluster. |
| O resultado não é um artefato da geografia inferida. | 05 | Com geografia exata por distrito, os sinais se mantêm. Spearman de share_violento: −0,545 (bairro) e −0,813 (distrito). |
| Os coeficientes não são comparáveis entre bairro e distrito. | 05 | A agregação infla a correlação (efeito ecológico). O efeito verdadeiro está entre −8,8% e −16,6% por 10 p.p. |
| Lição metodológica central: dado faltante não falta ao acaso. | 05 | A geolocalização dos crimes varia de 2,8% a 97,8% entre bairros. A coordenada foi usada como chave, não como amostra. |

**Ressalva principal do projeto.** Não há controle de renda nem de população. Por isso nenhum resultado pode ser lido como "o crime causa a desvalorização".

---

## 2. Contexto necessário (notebooks 01 e 02)

Os notebooks 03 a 05 consomem arquivos produzidos pelos dois primeiros. Este resumo basta para entender o resto.

### 2.1 Fontes de dados

| Fonte | Conteúdo | Granularidade |
|---|---|---|
| SSP-SP, `SPDadosCriminais_2022..2026.xlsx` | Boletins de ocorrência | 1 linha por ocorrência, estado inteiro (só ~40% é da capital) |
| Guias de ITBI pagas (`GUIAS*.xlsx`) | Transações de imóveis | 1 linha por transação, só a capital |
| GeoSampa (WFS da Prefeitura) | Malha de 96 distritos e de setores fiscais | Polígonos |
| IBGE, tabela 1737 | IPCA (número-índice e variação mensal) | Mensal |

### 2.2 O problema da chave geográfica

As duas bases não têm chave geográfica em comum:
- No ITBI, o campo `Bairro` vem preenchido em apenas ~61% das linhas.
- Nos crimes, o bairro vem em 99,7% e a lat/lon em ~81%.

A solução foi construir o bairro de cada transação por inferência, em três camadas, nesta ordem de prioridade:
1. O `Bairro` do próprio ITBI (38,8%).
2. Propagação por CEP (cobertura acumulada de 92,6%).
3. Dicionário de ruas, derivado dos crimes e desambiguado por distrito (cobertura combinada de 97,0%).

Um erro conhecido, já corrigido: a rua `MELO ALVES` existe duas vezes na cidade. O dicionário mandou 255 apartamentos dos Jardins para o Jardim Monte Verde. Por isso o CEP passou a ter prioridade sobre o nome da rua.

### 2.3 Chave exata do ITBI

Os 3 primeiros dígitos do número `SQL` do imóvel são o **setor fiscal** do IPTU. Esse código existe nos dois lados, no ITBI e na malha do GeoSampa, então a ligação `SQL → setor fiscal → distrito` é exata, sem inferência. Cada setor recebe o distrito com que tem maior sobreposição de área. Esse caminho é a base do notebook 05.

### 2.4 Tratamento do preço (notebook 02)

- O preço é deflacionado pelo IPCA para **reais constantes do mês de referência, 2026-08**. Fator = índice do mês de referência ÷ índice do mês da transação.
- Sem a deflação, a variável `ano` absorveria uma valorização que não existiu. O notebook 03 informa que o preço nominal subiu 22,6% no período, enquanto o real caiu 1,2%.
- Filtros: 2022–2026; só proporção transmitida de 100% (o filtro mais caro, ~metade da base); sem vaga de garagem; só meses cobertos pelo IPCA.
- Resultado: **312.248 transações em 431 bairros**.

### 2.5 As quatro medidas de crime (notebook 02)

Os crimes são agrupados em 4 grupos: trânsito/culposo, patrimonial sem violência (furto), violento (roubo, homicídio, lesão corporal, estupro, extorsão, sequestro, latrocínio) e outros. A regra de trânsito/culposo é aplicada primeiro; sem isso, "lesão corporal culposa" seria contada como violência.

Para cada bairro, as taxas anuais usam só os anos completos **2022–2025**. Foram definidas quatro medidas:

| Medida | O que mede | Fraqueza |
|---|---|---|
| Crimes totais por ano | Volume absoluto | Contaminada por fluxo de pessoas |
| Crimes violentos por ano | Volume do que assusta | Idem, em menor grau |
| **% de crimes violentos** (`share_violento`) | **Composição** | Ignora o nível |
| Violentos por mil transações | Ajustado pelo estoque imobiliário | Aproximação grosseira de população |

Distinção conceitual que organiza todo o projeto:
- **Exposição** é o volume de ocorrências. Cresce onde há muita gente circulando (furtos em polos comerciais).
- **Risco** é a proporção de ocorrências graves.

Bairros com menos de 30 transações ou menos de 30 ocorrências são descartados por ruído.

### 2.6 Resultado do notebook 02 que motiva os seguintes

| Medida de crime | Spearman ρ com R$/m² mediano do bairro |
|---|---|
| Crimes totais por ano | +0,146 (levemente positiva) |
| Crimes violentos por ano | +0,036 (nenhuma relação) |
| **% de crimes violentos** | **−0,548** (p ≈ 10⁻³⁵) |
| Violentos por mil transações | −0,296 |

Duas leituras apoiam o resultado:
- O R$/m² mediano cai cerca de 31% do quintil menos violento (por % violento) ao mais violento.
- Remover os 7 bairros do centro comercial quase não muda o ρ (−0,548 para −0,554). Logo, a correlação não é artefato de fluxo.

Esses números vêm do texto do notebook 02 (430 bairros). O cache atual tem 431 bairros, então os valores podem diferir ligeiramente (ver seção 8).

---

## 3. Notebook 03 — Modelagem hedônica

### 3.1 Objetivo

O notebook 02 achou uma correlação forte, mas correlação simples não responde a duas perguntas. O notebook 03 as responde:
1. **O sinal sobrevive depois de descontar as características físicas do imóvel?** Quem responde é o modelo hedônico, que explica o preço pelos atributos do bem.
2. **Quanto do sinal é crime e quanto é só "em que parte da cidade o imóvel está"?** Quem responde é a seção 6, que compara contra um teto de "identidade do bairro".

### 3.2 Variáveis

**Alvo:** `log(preço por m²)`. O log torna simétrica a cauda longa dos preços e permite ler o coeficiente como variação percentual.

**Atributos do imóvel:**
- Numéricos: `area_construida`, `area_terreno`, `testada`, `fracao_ideal`, `idade`, `ano`.
- Categóricos: `uso_desc`, `padrao_desc`.

**Variáveis de crime** (duas das quatro medidas):
- `log_violentos` = log₁₀(1 + crimes violentos por ano no bairro). É o **nível**.
- `share_violento` = fração de crimes violentos no bairro. É a **composição**.

**Por que só duas das quatro.** A correlação entre `log_crimes_total` e `log_violentos` é 0,98. Usar as duas juntas numa regressão espalharia o efeito de forma arbitrária. A correlação entre `log_violentos` e `share_violento` é só −0,10, então as duas medem coisas distintas.

**Estatísticas descritivas (312.248 transações):**

| Variável | Média | Mediana | Mínimo | Máximo |
|---|---|---|---|---|
| area_construida (m²) | 133,3 | 107 | 2 | 6.425 |
| area_terreno (m²) | 5.166,8 | 1.655 | 24 | 233.403 |
| idade (anos) | 26,0 | 22 | 0 | 107 |
| ano | 2023,96 | 2024 | 2022 | 2026 |
| log_violentos | 2,61 | 2,97 | 0,76 | 3,65 |
| share_violento | 0,33 | 0,31 | 0,10 | 0,67 |
| alvo (log R$/m²) | 8,49 | 8,53 | 3,62 | 9,83 |

### 3.3 Divisão dos dados

Divisão **aleatória** (não temporal) em 70% treino, 15% validação e 15% teste, com `random_state=42`. A escolha aleatória é deliberada: a hipótese é transversal, isto é, compara bairros no mesmo período.

| Partição | Transações | Alvo médio | R$/m² mediano |
|---|---|---|---|
| Treino | 218.573 | 8,489 | R$ 5.059 |
| Validação | 46.837 | 8,485 | R$ 5.045 |
| Teste | 46.838 | 8,487 | R$ 5.055 |

Papéis: o treino ajusta os parâmetros; a validação escolhe entre modelos; o teste mede o desempenho final **uma única vez**. As três partições têm perfil de preço praticamente idêntico, como esperado numa amostra aleatória grande.

### 3.4 Pipeline e quatro modelos

O pré-processamento fica dentro de um `Pipeline` do scikit-learn e nunca é aplicado antes da divisão. Assim evita-se vazamento de dados: imputador e `StandardScaler` aprendem só no treino.

- **Numéricas:** imputação pela mediana, depois padronização.
- **Categóricas** (`uso`, `padrão`): `OneHotEncoder` com `min_frequency=50` e `handle_unknown="ignore"`.
- **`bairro`** (431 categorias; só no conjunto 3 da seção 6): `TargetEncoder`, que substitui o bairro pela média do alvo, com validação cruzada interna de 5 dobras e semente fixa.

| Modelo | Papel | Configuração |
|---|---|---|
| Ridge | Baseline linear (piso) | alpha = 1,0 |
| Random Forest | Ensemble de árvores (bagging) | 200 árvores, profundidade máxima 20, folha mínima 5 |
| HistGradientBoosting | Ensemble de árvores (boosting) | 400 iterações, learning rate 0,08 |
| SVM (SVR, kernel RBF) | Família baseada em margem | Treina numa amostra de 20 mil linhas |

A SVM treina em 20 mil transações porque o custo do SVR cresce de forma aproximadamente quadrática com o número de linhas. O pré-processador dela vê o treino inteiro, para manter a mesma codificação dos outros modelos. Só o regressor final usa a amostra. Os outros três usam as 218 mil linhas de treino.

### 3.5 Resultados da escolha de modelo (treino e validação)

Métricas: RMSE e MAE em escala log (menor é melhor), R² (maior é melhor) e **erro mediano**, que é o erro relativo típico numa transação, calculado como mediana de |exp(previsto − real) − 1|.

| Modelo | Partição | RMSE (log) | MAE (log) | R² | Erro mediano |
|---|---|---|---|---|---|
| Ridge | treino | 0,5392 | 0,3743 | 0,0886 | 27,3% |
| Ridge | validação | 0,5429 | 0,3755 | 0,0879 | 27,3% |
| **Random Forest** | treino | 0,3468 | 0,2010 | 0,6231 | 12,8% |
| **Random Forest** | validação | 0,4241 | 0,2462 | **0,4433** | 15,3% |
| SVM | treino | 0,4972 | 0,3270 | 0,2251 | 22,8% |
| SVM | validação | 0,5037 | 0,3289 | 0,2148 | 22,8% |
| HistGradientBoosting | treino | 0,4157 | 0,2571 | 0,4583 | 17,5% |
| HistGradientBoosting | validação | 0,4351 | 0,2645 | 0,4141 | 17,7% |

**Leitura:**
- O **Ridge** fica longe dos demais (R² ≈ 0,09). A relação entre preço e atributos não é linear.
- A **SVM** treinada em 20 mil linhas também fica baixa (R² ≈ 0,21).
- O **Random Forest vence** na validação (R² = 0,4433).
- O Random Forest tem o maior descolamento entre treino e validação: 0,62 contra 0,44. Por isso o R² de treino não serve como estimativa de desempenho.

### 3.6 Tentativa de reduzir o overfitting do Random Forest

Foram testadas 10 configurações de regularização, avaliadas **só na validação** (o teste continua intocado).

| Tentativa | R² treino | R² validação | Gap |
|---|---|---|---|
| atual (folha 5, prof. 20) | 0,6231 | 0,4433 | 0,1798 |
| folha mínima 20 | 0,5210 | 0,4327 | 0,0883 |
| folha mínima 50 | 0,4604 | 0,4082 | 0,0521 |
| profundidade 10 | 0,3908 | 0,3552 | 0,0356 |
| max_features 0,5 | 0,6061 | 0,4493 | 0,1568 |
| folha 20 + max_features 0,5 | 0,5102 | 0,4323 | 0,0779 |
| max_features 0,3 | 0,5702 | 0,4395 | 0,1307 |
| max_features 0,7 | 0,6160 | 0,4486 | 0,1674 |
| max_features 0,5 + folha 2 | 0,6735 | 0,4418 | 0,2317 |
| max_features 0,5 + sem limite de profundidade | 0,6348 | 0,4503 | 0,1844 |

**Conclusão da seção:**
- Regularizar fecha o gap (para ~0,04–0,05 com folha mínima 50 ou profundidade 10), mas **derruba o R² de validação**. O overfitting aparente é, em boa parte, capacidade útil do modelo.
- O único ganho é `max_features` entre 0,5 e 0,7, cerca de +0,006 a +0,007 de R². Isso é pouco acima do ruído entre configurações.
- **Decisão:** manter os hiperparâmetros originais, para não mexer nos números do resto do notebook.

### 3.7 Desempenho final no teste

Avaliado uma única vez, com o Random Forest original:

| Partição | RMSE (log) | MAE (log) | R² | Erro mediano |
|---|---|---|---|---|
| Treino | 0,3468 | 0,2010 | 0,6231 | 12,8% |
| Validação | 0,4241 | 0,2462 | 0,4433 | 15,3% |
| **Teste** | **0,4185** | **0,2448** | **0,4511** | **15,4%** |

O teste fica muito próximo da validação (0,4511 contra 0,4433). Isso indica que a escolha do modelo não contaminou a estimativa final.

Figuras descritas em palavras:
- *Previsto × real no teste:* nuvem de 6 mil pontos em escala log-log em torno da diagonal.
- *Resíduos:* histograma centrado em zero, com uma cauda de **superestimação**.

### 3.8 Seção 6 — O crime agrega informação? (evidência central)

Mesmo pipeline, mesma divisão, mesmas sementes. Só muda o conjunto de variáveis, usando o Random Forest vencedor:

| Conjunto | RMSE (log) | MAE (log) | R² (teste) | Erro mediano |
|---|---|---|---|---|
| 1 · só atributos do imóvel (piso) | 0,4459 | 0,2710 | 0,3768 | 17,1% |
| 2 · imóvel + crime | 0,4185 | 0,2448 | 0,4511 | 15,4% |
| 3 · imóvel + identidade do bairro (teto) | 0,4177 | 0,2446 | 0,4531 | 15,5% |

Quantificação:

| Indicador | Valor |
|---|---|
| Ganho de R² das 2 variáveis de crime (conjunto 2 − conjunto 1) | **+0,0743** |
| Ganho de R² da identidade do bairro (conjunto 3 − conjunto 1), o "teto" | +0,0763 |
| Razão crime ÷ teto | **97%** |
| Redução do RMSE no teste pelo crime | **6,15%** |

**Por que o conjunto 3 existe.** Ele responde à objeção mais séria: as variáveis de crime não seriam só um jeito disfarçado de dizer em que parte da cidade o imóvel está? Em parte, sim. Duas variáveis de crime recuperam 97% do que os 431 rótulos de bairro recuperam.

**Faca de dois fios:**
- *Lado impressionante:* dois números resumem quase toda a informação de localização que importa para o preço.
- *Lado cauteloso:* em São Paulo, localização é renda, infraestrutura, transporte, escolaridade e história urbana. O ganho de R² **não é o efeito causal do crime**.

O ganho do conjunto 2 sobre o 1 vale como resultado preditivo de qualquer forma: crime prevê preço. O conjunto 3 só impede interpretá-lo como efeito do crime em si.

### 3.9 Interpretação do modelo vencedor

O Random Forest não tem coeficientes. Foram usadas duas técnicas:
- **Importância por permutação** mede quanto o R² cai ao embaralhar uma variável. Foi usada uma subamostra de 4 mil linhas da validação, com 5 repetições.
- **Dependência parcial** mostra a forma do efeito médio.

Resultados lidos do texto do notebook (os valores exatos estão só nos gráficos):
- `share_violento` **domina**. Embaralhá-la derruba o R² em ~0,46, quase o dobro de qualquer atributo do imóvel. A dependência parcial desce de ~8,5 para ~8,2 (em log R$/m²) conforme a fração violenta sobe. A curva é ruidosa no começo, onde há poucos bairros.
- `log_violentos` fica em terceiro lugar (~0,19), mas **sem forma monotônica**: sobe e desce ao longo do eixo.
- Interpretação: com apenas 431 valores distintos, a floresta usa `log_violentos` como **identificador de bairro** (o teto de localização da seção 6), não como efeito de crime.

### 3.10 Teste estatístico formal (OLS no nível de bairro)

Pergunta: o coeficiente é distinguível de zero?

A regressão é no nível de **bairro**, e não de transação. Isso é deliberado. Trezentas mil transações repetiriam a mesma informação de crime milhares de vezes e estreitariam artificialmente o intervalo de confiança.

Fórmula: `log(preço mediano do bairro) ~ log_violentos + share_violento + log(área mediana) + idade mediana + log(nº de transações)`. O número de transações aproxima o porte do bairro.

| Termo | Coeficiente | Erro-padrão | t | p | IC 95% |
|---|---|---|---|---|---|
| Intercepto | 6,8275 | 0,197 | 34,70 | 0,000 | [6,441; 7,214] |
| log_violentos | −0,0287 | 0,021 | −1,40 | 0,162 | [−0,069; 0,012] |
| **share_violento** | **−0,9220** | 0,110 | −8,40 | 6,7×10⁻¹⁶ | **[−1,138; −0,706]** |
| log(área mediana) | 0,3333 | 0,036 | 9,15 | 0,000 | [0,262; 0,405] |
| idade mediana | 0,0014 | 0,001 | 1,37 | 0,172 | [−0,001; 0,003] |
| log(nº transações) | 0,0630 | 0,011 | 5,66 | 0,000 | [0,041; 0,085] |

R² = 0,416, com 431 bairros.

**Efeitos práticos** (mantidos os controles):
- +10 pontos percentuais no % de crimes violentos → **−8,8%** no R$/m² do bairro. O cálculo é exp(−0,922 × 0,10) − 1.
- Dobrar a contagem de crimes violentos → −0,9%, **não significativo**.

O gráfico do intervalo de confiança mostra o IC de `share_violento` longe da linha vermelha do zero. O de `log_violentos` cruza o zero. Isso bate com a dependência parcial irregular da seção 3.9.

### 3.11 Conclusão do notebook 03

**A hipótese é confirmada, com uma correção na formulação:** não é a quantidade de crime que se precifica, é a violência dele.

| Evidência | Resultado |
|---|---|
| Correlação com contagem de crimes violentos (431 bairros) | ρ = +0,04, nenhuma relação |
| Correlação com % de crimes violentos | ρ = −0,55 |
| R$/m² real, quintil menos violento → mais violento | R$ 5.045 → R$ 3.487 (**−31%**) |
| Ganho de R² no teste ao incluir crime | **+0,074** (0,377 → 0,451) |
| Redução do RMSE no teste | **−6,2%** |
| OLS, coeficiente de `share_violento` | **−0,922**, IC 95% [−1,14; −0,71] |
| OLS, coeficiente de `log_violentos` | −0,029 (p = 0,16), não significativo |
| Efeito de +10 p.p. no % violento | **−8,8%** |

O modelo erra 15,4% numa transação típica. O notebook considera isso razoável para um modelo sem informação de conservação, vista, andar ou reforma.

**Limitações e o que destravaria cada uma:**

| Limitação | O que resolveria |
|---|---|
| Sem taxa de crime por habitante, só por transação | População por setor censitário **de São Paulo** (a do repositório é do Rio) |
| Sem controle de renda, o confundidor mais provável | Renda por setor censitário do Censo 2022 |
| Sem argumento causal | Janela temporal mais longa, para observar preço depois de mudança no crime |
| Geografia por CEP e nome de rua, não por coordenada do imóvel | Geocodificação do endereço do ITBI |

A primeira linha é a mais importante. Com população, `share_violento` poderia virar taxa de homicídio ou roubo por 100 mil habitantes. Com renda no modelo, a pergunta vira "crime prevê preço **a renda constante**?", que é a pergunta que interessa.

---

## 4. Notebook 04 — Clusterização de bairros

### 4.1 Objetivo

Os notebooks anteriores mediram **uma relação**. Este pergunta outra coisa: **que perfis de bairro a cidade tem?**

A intuição a testar é que existiriam **quatro tipos** de bairro: caro com pouca violência, caro com muita, barato com pouca e barato com muita. Cortando nas medianas, os quatro quadrantes de fato têm bairros dentro (contagens atuais na seção 4.8). Mas "ter bairros dentro de um quadrante" não é o mesmo que "ser um tipo". Os quadrantes vêm de dois cortes arbitrários, e um cluster só existe se os bairros se parecem entre si e se distinguem dos outros.

**O que clusterização é e não é.** Ela não testa hipótese nem prova nada. É descrição, uma forma de organizar 431 linhas em poucos grupos legíveis. Não há número "certo" de clusters, e outras variáveis produziriam outros grupos. O valor está em os grupos serem nomeáveis e reconhecíveis por quem conhece a cidade.

### 4.2 Variáveis e padronização

Seis variáveis cobrem três dimensões de um bairro:

| Variável | Dimensão |
|---|---|
| `preco_m2_mediano` | Preço |
| `share_violento` | **Composição** do crime |
| `log_crimes_total` | **Exposição** (volume absoluto) |
| `log_violentos_por_1k` | Violência normalizada pelo estoque imobiliário |
| `area_mediana` | Tipo de imóvel |
| `idade_mediana` | Idade do estoque |

Compor e exposição entram **separadas** de propósito. Isso permite ao algoritmo distinguir "bairro movimentado onde se furta muito" de "bairro tranquilo onde o pouco que acontece é violento".

**Padronizar é obrigatório.** O KMeans usa distância euclidiana. Sem padronizar, o preço (de ~983 a ~14.892) esmagaria `share_violento` (de 0,10 a 0,67). O `StandardScaler` deixa todas as variáveis com média 0 e desvio 1, e cada uma pesa igual.

Amplitude das variáveis (431 bairros, nenhum nulo):

| Variável | Mínimo | Mediana | Máximo |
|---|---|---|---|
| preco_m2_mediano | 982,67 | 4.073,63 | 14.891,68 |
| share_violento | 0,10 | 0,40 | 0,67 |
| log_crimes_total | 1,49 | 2,14 | 4,08 |
| log_violentos_por_1k | 0,85 | 2,42 | 4,59 |
| area_mediana | 49 | 103 | 472 |
| idade_mediana | 2 | 24 | 60 |

### 4.3 Quantos clusters? (k)

Dois indicadores, KMeans com `n_init=20` e `random_state=42`, k de 2 a 10:

- **Silhueta** mede se cada bairro está mais perto do próprio grupo que do vizinho. Vai de −1 a 1. Tem máximo, então serve para escolher.
- **Inércia** é a soma das distâncias ao centroide. Cai sempre, então só serve para procurar o "joelho".

Os dois aparecem em painéis separados (não em um gráfico de dois eixos y), para que ninguém "escolha" o joelho mexendo nos limites dos eixos.

| k | Silhueta | Inércia |
|---|---|---|
| 2 | 0,250 | 2060,5 |
| **3** | **0,276** | 1636,8 |
| 4 | 0,255 | 1409,2 |
| 5 | 0,217 | 1240,6 |
| 6 | 0,225 | 1132,8 |
| 7 | 0,219 | 1051,9 |
| 8 | 0,228 | 984,9 |
| 9 | 0,198 | 926,8 |
| 10 | 0,199 | 866,2 |

A melhor silhueta está em **k = 3**.

**Referência de interpretação usada no notebook:** acima de 0,5, estrutura clara; entre 0,25 e 0,5, estrutura fraca mas real; abaixo de 0,25, os grupos são mais um corte conveniente do que tipos naturais. Com 0,276, k = 3 fica na faixa "fraca mas real", pouco acima do limite.

### 4.4 Validação cruzada: agrupamento hierárquico de Ward

O algoritmo de Ward funde repetidamente os dois grupos mais parecidos. O dendrograma mostra a distância de cada fusão. Um **salto grande** na altura indica que a fusão juntou coisas muito diferentes, então o número de grupos imediatamente antes dela era significativo.

**Resultado executado:**
- k com maior salto no dendrograma: **4**.
- k com melhor silhueta no KMeans: **3**.

Os dois métodos **não coincidem** na saída executada. Ward aponta 4 e a silhueta aponta 3. O texto de conclusão do notebook diz que concordam em k = 3, o que contradiz esse resultado (ver seção 8). Foi por isso que o notebook ajustou os dois agrupamentos, k = 3 e k = 4, e os comparou.

### 4.5 Ajuste e rotulagem

Os rótulos do KMeans são arbitrários (o "cluster 0" não significa nada e muda entre execuções). Por isso os clusters são **renumerados por preço mediano crescente**. Assim o cluster 0 é sempre o de menor preço e o maior número é o de maior preço.

Tamanho dos clusters:

| Cluster | k = 3 | k = 4 |
|---|---|---|
| 0 | 247 | 79 |
| 1 | 109 | 217 |
| 2 | 75 | 111 |
| 3 | — | 24 |

### 4.6 O que cada cluster é (k = 3), pelas saídas executadas

O heatmap mostra a média padronizada de cada variável por cluster: quantos desvios-padrão acima ou abaixo da média da cidade. A rampa é divergente (azul abaixo, cinza na média, vermelho acima), porque aqui o zero significa "igual à cidade".

Etiquetas automáticas, que listam as duas variáveis em que cada cluster mais se destaca:
- Cluster 0: volume de crimes **baixo** · % de crimes violentos **alto**.
- Cluster 1: volume de crimes **alto** · violentos por mil transações **alto**.
- Cluster 2: preço do m² **alto** · área construída mediana **alta**.

Perfil em unidades originais:

| Cluster | Bairros | R$/m² mediano | % violento | Crimes/ano (mediana) | Violentos/1k | Área (m²) | Idade (anos) | Transações |
|---|---|---|---|---|---|---|---|---|
| 0 | 247 | R$ 3.875 | 43,9% | 74 | 207 | 100 | 24 | 67.893 |
| 1 | 109 | R$ 4.056 | 36,2% | 2.590 | 1.401 | 97 | 22 | 159.721 |
| 2 | 75 | R$ 6.655 | 28,4% | 227 | 135 | 143 | 23 | 84.634 |

**Bairros mais típicos** (os mais próximos do centroide), conforme o texto do notebook:
- Cluster 0: Vila Butantã, Vila Bancária, Vila Palmeiras, Parque Peruche, Vila Buenos Aires.
- Cluster 1: Cachoeirinha, Vila Maria, Freguesia do Ó, Penha, Limão, Cangaíba.
- Cluster 2: Vila Romana, Pompeia, Alto da Lapa, Paraíso, Vila Clementino, Brooklin Paulista.

**Nomes em prosa:**
- **Cluster 0 — bairro residencial miúdo.** Pouquíssimo crime em volume absoluto (~74 por ano, o menor dos três) e a **maior fração violenta (~44%)**. São bairros pequenos, essencialmente residenciais, onde pouca coisa acontece, mas o que acontece é proporcionalmente mais grave. É o tipo mais comum da cidade.
- **Cluster 1 — distrito populoso de alto fluxo.** Cerca de 2.590 ocorrências por ano, ~35 vezes mais que o cluster 0, com preço quase igual ao dele. São os grandes distritos populosos, com comércio de rua, terminais de ônibus e muita circulação.
- **Cluster 2 — miolo valorizado.** Preço ~60–70% acima dos outros dois, a **menor fração violenta (~28–29%)** e os maiores imóveis (143 m², contra ~97–100 m²). Volume de crime intermediário.

### 4.7 O que a comparação entre os clusters 0 e 1 mostra

Os dois têm preço quase igual (R$ 3.875 e R$ 4.056) e volume de crime que difere **cerca de 35 vezes**. É a demonstração mais direta do achado dos notebooks anteriores: **contar ocorrências não diz nada sobre preço**. Se a contagem derrubasse o preço, o cluster 1 teria de ser drasticamente mais barato. Não é.

O que ordena o preço dos três grupos é a **fração violenta**, que cai monotonicamente enquanto o preço sobe: **43,9% → 36,2% → 28,4%**.

O algoritmo foi quem organizou a cidade assim, com seis variáveis e sem saber que preço e violência eram o interesse. Ele separou os bairros não caros por **tamanho e fluxo**, e não por violência.

### 4.8 Os quatro quadrantes contra os clusters

Cortes nas medianas: preço = **R$ 4.074** e violência = **39,5%**.

**Os quatro quadrantes por conta própria:**

| Quadrante | Bairros | R$/m² mediano | Crimes/ano | Área (m²) | Transações (mediana) |
|---|---|---|---|---|---|
| barato + muita violência | 155 | R$ 3.514 | 96 | 98 | 147 |
| barato + pouca violência | 60 | R$ 3.705 | 216 | 99 | 201 |
| caro + muita violência | 61 | R$ 4.490 | 66 | 115 | 239 |
| caro + pouca violência | 155 | R$ 5.087 | 289 | 117 | 513 |

Como o corte é na mediana, os quadrantes "esperados" (barato+muita e caro+pouca) são grandes e os "cruzados" são pequenos.

**Clusters (k = 4) contra quadrantes** (contagem de bairros):

| Cluster k=4 | barato + muita | barato + pouca | caro + muita | caro + pouca | Total |
|---|---|---|---|---|---|
| 0 | 33 | 20 | 7 | 19 | 79 |
| 1 | 121 | 27 | 44 | 25 | 217 |
| 2 | 0 | 12 | 4 | 95 | 111 |
| 3 | 1 | 1 | 6 | 16 | 24 |

**Concentração de cada cluster no quadrante dominante** (saída executada):
- Cluster 0: 42% em "barato + muita violência".
- Cluster 1: 56% em "barato + muita violência".
- Cluster 2: **86%** em "caro + pouca violência".
- Cluster 3: 67% em "caro + pouca violência".

### 4.9 Veredito sobre a intuição dos quatro tipos

> **Dos quatro tipos imaginados, só um é um tipo de verdade.**

- **"Caro + pouca violência" é um tipo real.** O cluster 2 (k = 4) concentra 86% dos seus bairros nesse quadrante, e o grupo tem identidade própria (imóveis grandes, miolo da cidade, estoque valorizado).
- **Os outros três não são tipos.** Clusters com 42–56% de concentração estão espalhados: quase metade do cluster mora em outro quadrante.
- **"Caro + muita violência"** nunca forma a maioria de nenhum cluster. A maior célula (44 bairros) está no cluster 1, mas isso é só ~20% desse cluster. Os 61 bairros desse quadrante são heterogêneos, unidos apenas por estarem do mesmo lado de dois cortes arbitrários.

**Por quê.** Entre os bairros não caros, o eixo que mais os separa é **tamanho e fluxo**, não violência. Dois bairros podem ter a mesma fração violenta e ser completamente diferentes: a Vila Bancária tem ~74 ocorrências por ano e a Penha tem ~2.590. O algoritmo, livre para escolher, separa por essa diferença, porque ela é muito maior.

**Compatibilidade com os notebooks 02 e 03:** violência **prevê preço**, mas **não define tipo de bairro**. A primeira afirmação é sobre a relação entre duas variáveis. A segunda é sobre a estrutura do conjunto. Não se contradizem.

### 4.10 Visualizações e mapas

**Gráficos descritos em palavras:**
- *Dispersão preço × % violento:* cada bolha é um bairro, com tamanho proporcional ao número de transações, colorida pelo cluster. O centroide recebe rótulo numérico dentro de um círculo, para que a identidade não dependa só de cor.
- *Projeção PCA em 2D:* as duas primeiras componentes somam apenas **58%** da variância das seis variáveis. A separação visual entre clusters é parcial.

**Mapas interativos (Plotly), sobre a malha oficial dos 96 distritos do GeoSampa.** A geometria foi simplificada a 20 metros, o que reduz o GeoJSON de 4 MB para 221 KB sem diferença visual nessa escala. Não se usa camada de mapa externa, então funcionam offline.

- **Atribuição transação → distrito exata**, sem inferência: SQL → setor fiscal → polígono → distrito. Atingiu **100,00% das transações**, em **94 dos 96 distritos**.
- As 312 mil transações têm só **~5.002 coordenadas distintas**, porque vêm de medianas de rua e centros de bairro. Por isso o mapa desenha 96 distritos preenchidos e 431 bairros como pontos, e não 312 mil marcadores.
- **Mapa 1:** cada distrito é pintado pelo cluster **predominante** das suas transações, e os bairros ficam por cima como pontos coloridos pelo próprio cluster. Onde preenchimento e pontos discordam, o distrito é internamente heterogêneo, e isso é informação, não erro.
- **Mapa 2:** cada distrito é pintado pela mediana do R$/m² real.

**Cinco distritos mais caros** (mediana de R$/m²):

| Distrito | R$/m² | Transações | Bairros | Cluster predominante |
|---|---|---|---|---|
| Jardim Paulista | 8.879 | 6.803 | 16 | 2 |
| Pinheiros | 8.810 | 3.662 | 11 | 1 |
| Moema | 8.345 | 8.991 | 19 | 2 |
| Itaim Bibi | 8.307 | 12.586 | 26 | 2 |
| Alto de Pinheiros | 8.217 | 4.238 | 18 | 2 |

Pinheiros aparece no cluster 1 (alto fluxo), o que é coerente com o fato de o cluster 1 reunir áreas de grande circulação.

### 4.11 Estabilidade: o quanto confiar nos grupos

A clusterização sempre devolve clusters, mesmo em dados sem estrutura. Teste: reajustar o KMeans (k = 3) em **20 subamostras de 80%** dos bairros e medir com que frequência dois bairros que ficaram juntos continuam juntos. A métrica não depende da numeração dos rótulos.

| Grupo de pares | Resultado (saída executada) |
|---|---|
| Pares que o agrupamento final colocou **juntos** | continuam juntos em **97,2%** das reamostragens |
| Pares que o agrupamento final colocou **separados** | aparecem juntos em apenas **1,9%** |

### 4.12 Conclusão do notebook 04

**Estáveis, mas não bem separados.** A partição é reproduzível (estabilidade ~97%), mas a silhueta de 0,276 indica estrutura fraca. Não há vazios entre os grupos. São Paulo é um **contínuo**, e os clusters são cortes reproduzíveis nesse contínuo, não ilhas. Tratá-los como categorias naturais com fronteiras nítidas seria superinterpretar.

**Os dois tipos extras que k = 4 revela** (valores do texto do notebook; ver a ressalva da seção 8 sobre tamanhos):
- **"Casarões de bairro-jardim":** poucos bairros, R$ ~6.900/m², área mediana de ~259 m², idade de ~36 anos. Exemplos: Jardim Guedala, Alto de Pinheiros, Jardim Panorama. Só ~5.900 transações em cinco anos, porque é um estoque que quase não gira.
- **"Vertical recente de classe média":** muitos bairros, R$ ~5.000/m², idade mediana de ~17 anos, ~28% de crimes violentos. Exemplos: Penha de França, Pompeia, Parada Inglesa. É o cluster com mais transações (~193 mil), as torres novas onde o mercado efetivamente acontece.

Vale olhar k = 4 por isso, mesmo que a silhueta prefira k = 3.

**Limitações declaradas no notebook:**
1. A clusterização não testa hipótese: não tem p-valor, nem conjunto de teste, nem resposta certa para k.
2. Outras variáveis dariam outros grupos. Com renda e população por setor censitário, os clusters provavelmente passariam a separar por nível socioeconômico, o confundidor que não foi controlado.
3. `share_violento` é instável em bairros com pouco crime. O corte de 30 ocorrências mínimas herdado do notebook 02 reduz o problema, mas um bairro com 40 ocorrências por ano ainda tem margem de erro visível.
4. As seis variáveis pesam igual por construção do `StandardScaler`. Isso é uma escolha. Dar mais peso ao preço geraria clusters mais parecidos com os quadrantes da intuição inicial.
5. O preenchimento por distrito é uma agregação de agregação: o cluster é do bairro, e o distrito recebe o cluster predominante. Por isso os pontos dos bairros ficam sobrepostos.

---

## 5. Notebook 05 — Checagem: a geografia inferida distorceu os resultados?

### 5.1 Objetivo

Os notebooks 02 a 04 usam o **bairro** como unidade, e o bairro de cada transação foi inferido (campo do IPTU, CEP e dicionário de nomes de rua). Essa camada já produziu um erro detectado (o caso Melo Alves). A pergunta deste notebook é: **a inferência distorceu as conclusões?**

Método: refazer a análise numa geografia **sem inferência do lado do imóvel** (SQL → setor fiscal → distrito) e comparar com a original. Se a inferência tivesse inventado o resultado, ele desapareceria na geografia exata.

### 5.2 Imóveis → distrito, exato

Cada transação tem o `SQL`, e os 3 primeiros dígitos são o setor fiscal. O notebook 01 já resolveu `setor fiscal → distrito` por maior sobreposição de área contra a malha do GeoSampa. Basta juntar.

Resultado: **100,00% das 312.248 transações** recebem distrito, em **94 dos 96 distritos**. Dois distritos não têm nenhuma transação.

### 5.3 Crimes → distrito, e a armadilha

O lado do crime não tem setor fiscal, mas tem latitude e longitude em ~81% das ocorrências. O caminho óbvio seria ponto-em-polígono direto: 1,79 milhão de pontos contra 96 polígonos. **Esse caminho está errado.**

**Por que está errado.** A disponibilidade de coordenada **não é aleatória**. Medida em 619 bairros com pelo menos 200 ocorrências:

| Estatística | % das ocorrências com coordenada |
|---|---|
| Mínimo | 2,8% |
| Percentil 5 | 36,1% |
| Mediana | 80,6% |
| Média | 75,0% |
| Percentil 95 | 92,5% |
| Máximo | 97,8% |

Os cinco bairros com pior geolocalização são periféricos:

| Bairro | % com coordenada |
|---|---|
| Guaianazes | 2,8% |
| Água Funda | 5,3% |
| Jardim São Januário | 5,5% |
| Jardim Avenida | 6,9% |
| Residencial Taipas | 6,9% |

Usar só o que está geolocalizado representaria Guaianazes por ~3% das suas ocorrências e a Vila Azevedo por ~97%. Como periferia e centro diferem exatamente na variável de interesse, isso **enviesaria o resultado na direção da hipótese**.

Há também **viés por tipo de crime**: ocorrências violentas são menos geolocalizadas (79,2%) que as patrimoniais (81,4%). Figuras descritas em palavras: um histograma da taxa de geolocalização por bairro, largo, com a média da cidade marcada; e um gráfico de barras comparando a composição dos crimes com e sem coordenada.

### 5.4 A correção: coordenada para mapear, nome para contar

Separam-se duas tarefas que o caminho ingênuo confunde:

| Tarefa | O que exige | Fonte |
|---|---|---|
| Descobrir **em que distrito** um bairro fica | Poucos pontos bastam | Coordenadas (ponto-em-polígono) |
| **Contar** quantas ocorrências o bairro teve | Todas as ocorrências | Nome do bairro (99,7% preenchido) |

Guaianazes tem só 2,8% de ocorrências geolocalizadas, mas isso ainda são 64 pontos, o bastante para estabelecer que ela fica no distrito de Guaianases. Estabelecido o distrito, contam-se **todas** as ocorrências do bairro, geolocalizadas ou não.

**Parâmetros:** mínimo de 20 pontos por bairro; o distrito atribuído é o modal (onde caem mais pontos do bairro); a **pureza** do mapeamento é reportada (pontos no distrito modal ÷ pontos totais), porque nomes de bairro não respeitam divisa de distrito.

**Resultados:**

| Indicador | Valor |
|---|---|
| Bairros mapeados para um distrito | 1.453 |
| Crimes com distrito | **99,23%** (contra 81,0% pelo caminho ingênuo) |
| Distritos alcançados | 96 de 96 |
| Pureza média do mapeamento, ponderada por ocorrência | **90,4%** |

É a diferença entre usar a coordenada como **amostra** (enviesada) e como **chave** (não enviesada).

### 5.5 Agregação por distrito

As mesmas quatro medidas de crime do notebook 02, agora por distrito. A definição de `grupo_crime` é idêntica, senão a comparação mediria a mudança de definição e não a de geografia. Para cada distrito calculam-se:
- A média anual de ocorrências em 2022–2025.
- O preço mediano, a área mediana, a idade mediana e o número de transações.

Filtro: pelo menos 30 transações e 30 ocorrências. Passam **93 distritos**, que cobrem **312.247 transações**.

Mais caros por distrito (ocorrências por ano em média 2022–2025):

| Distrito | R$/m² | Transações | Crimes/ano | Violentos/ano | % violento |
|---|---|---|---|---|---|
| Jardim Paulista | 8.879 | 6.803 | 7.410 | 1.787 | 24% |
| Pinheiros | 8.810 | 3.662 | 12.279 | 3.344 | 27% |
| Moema | 8.345 | 8.991 | 4.721 | 1.332 | 28% |
| Itaim Bibi | 8.307 | 12.586 | 10.372 | 3.225 | 31% |
| Alto de Pinheiros | 8.217 | 4.238 | 1.652 | 521 | 32% |
| Consolação | 7.171 | 7.576 | 10.538 | 2.784 | 26% |

Note que Pinheiros e Itaim Bibi combinam preços altos com volumes enormes de crime. São polos de circulação.

### 5.6 A comparação: bairro (inferido) contra distrito (exato)

Correlações entre cada medida de crime e o R$/m² mediano (saída executada):

| Medida de crime | Geografia | Pearson r | Spearman ρ | p (Spearman) | Unidades |
|---|---|---|---|---|---|
| Crimes totais/ano (log) | bairro (inferido) | +0,127 | +0,150 | 1,8×10⁻³ | 431 |
| | distrito (exato) | +0,406 | +0,377 | 2,0×10⁻⁴ | 93 |
| Crimes violentos/ano (log) | bairro (inferido) | +0,045 | +0,037 | 0,44 | 431 |
| | distrito (exato) | +0,054 | −0,002 | 0,99 | 93 |
| **% de crimes violentos** | bairro (inferido) | −0,451 | **−0,545** | 9,8×10⁻³⁵ | 431 |
| | distrito (exato) | −0,681 | **−0,813** | 4,1×10⁻²³ | 93 |
| Violentos por mil transações (log) | bairro (inferido) | −0,226 | −0,283 | 2,3×10⁻⁹ | 431 |
| | distrito (exato) | −0,566 | −0,703 | 4,0×10⁻¹⁵ | 93 |

Descrição da figura: dois gráficos de dispersão lado a lado (% violento × R$/m²), um por geografia, com reta de regressão. Ambos mostram a mesma inclinação negativa.

### 5.7 Teste formal nas duas geografias

Mesma regressão OLS do notebook 03, com os mesmos controles:

| Geografia | Unidades | R² | Coef. `share_violento` | IC 95% | p-valor | Efeito de +10 p.p. |
|---|---|---|---|---|---|---|
| Bairro (inferido) | 431 | 0,416 | −0,922 | [−1,138; −0,706] | 6,7×10⁻¹⁶ | −8,8% |
| Distrito (exato) | 93 | 0,750 | −1,816 | [−2,276; −1,355] | 1,1×10⁻¹¹ | −16,6% |

Os dois intervalos de confiança ficam inteiros do lado negativo, longe do zero. O resultado do bairro (−0,922, R² 0,416, 431 bairros) é idêntico ao da seção 3.10, o que confirma a consistência entre os notebooks 03 e 05.

### 5.8 Veredito

#### A inferência não inventou nada

Os quatro sinais e a ordenação entre as medidas se preservam nas duas geografias:
- Crimes totais: positivo e fraco no bairro, positivo e moderado no distrito.
- Crimes violentos em quantidade: sem relação.
- **% violento: forte e negativo.**
- Violentos por mil transações: negativo.

**Contagem de crime não prevê preço (se algo, prevê preço mais alto). Composição prevê, e forte.** Se a camada de nome de rua e CEP tivesse fabricado o resultado, ele teria desaparecido aqui. Não desapareceu.

#### Mas os números do distrito não são "os verdadeiros"

Na geografia exata tudo fica mais forte (ρ de −0,545 para −0,813; R² de 0,416 para 0,750; coeficiente de −0,922 para −1,816). É tentador concluir que o bairro subestimava o efeito real. Dois motivos diferentes produzem esse aumento, e só um deles é mérito:

1. **Menos erro de atribuição.** É mérito real. Cada transação mal etiquetada atenua a correlação, e remover a inferência recupera sinal genuíno.
2. **Agregação.** É artefato. Os 93 distritos têm cerca de 3.400 transações cada, e suas medianas são mais estáveis que as de 431 bairros com cerca de 730. Agregar **sempre** eleva coeficientes de correlação, melhorando ou não a medição. É o **efeito ecológico** clássico: comparar ρ entre níveis de agregação diferentes é comparar coisas distintas.

Como os dois motivos empurram na mesma direção, **não é possível separá-los** com os dados disponíveis. A leitura defensável:
- Os **sinais e a ordenação** das quatro medidas são comparáveis e idênticos. Isso valida a análise original.
- As **magnitudes não são comparáveis** entre os níveis. O −8,8% do bairro é provavelmente atenuado, e o −16,6% do distrito é provavelmente inflado pela agregação. O efeito real está entre os dois, e não dá para dizer mais perto de qual.

#### O que ficou mais nítido

A relação **positiva** entre volume de crime e preço aparece mais clara no distrito (ρ de +0,15 para +0,38). Pinheiros registra ~12.279 ocorrências por ano e custa R$ 8.810/m². Itaim Bibi tem ~10.372 e R$ 8.307. São os distritos de maior circulação da cidade. Isso reforça a distinção entre **exposição** e **risco**.

#### Lição metodológica

> **Dado faltante raramente falta ao acaso.** Usar o subconjunto completo como se fosse uma amostra aleatória é o erro mais fácil de cometer e o mais difícil de notar.

O desenho natural da checagem (ponto-em-polígono sobre a coordenada) confirmaria a hipótese por motivo falso. A correção foi usar a coordenada como **chave** para descobrir o distrito do bairro e o **nome** do bairro para contar. Isso levou a cobertura de 81% para 99,23%, sem viés regional.

### 5.9 Limitações do notebook 05

- **A pureza do mapeamento é 90,4%, não 100%.** Nomes de bairro não respeitam divisa de distrito. "Guaianazes" tem pureza de 51,6%, espalhada entre Guaianases e Lajeado. Cerca de 10% das ocorrências ficam contadas no distrito vizinho ao correto.
- **O lado do imóvel é exato; o lado do crime não é.** O crime ainda depende do nome do bairro para ser contado.
- **`setor fiscal → distrito` usa maior sobreposição de área**, e não a localização exata do lote. Um setor que cruza a divisa de distrito tem todas as transações atribuídas ao distrito dominante.
- **93 distritos é pouco para o modelo hedônico.** Esta checagem compara correlação e OLS agregado, não o modelo de 312 mil transações do notebook 03, que continua valendo no nível da transação.

---

## 6. Como as peças se encaixam

1. **Notebook 02** mostra a correlação: a composição do crime (% violento) tem ρ ≈ −0,55 com o preço, e a contagem não tem relação.
2. **Notebook 03** mostra que o sinal sobrevive ao controle dos atributos do imóvel (OLS) e melhora a previsão (+0,074 de R²). Também mostra que o crime captura 97% da informação de localização do bairro, então o sinal é preditivo, não causal.
3. **Notebook 04** mostra que violência prevê preço, mas não define tipos de bairro. O que separa bairros é tamanho e fluxo, e só "caro + pouca violência" é um tipo real.
4. **Notebook 05** mostra que a conclusão não é produto da geografia inferida. Os sinais se mantêm numa geografia exata, mas as magnitudes não são comparáveis entre níveis de agregação.

O confundidor não controlado em todas as etapas é a **renda**. A segunda lacuna é a **população**, que impede transformar contagens em taxas por habitante.

---

## 7. Notas do redator deste relatório (observações próprias, não dos notebooks)

Estas observações não aparecem nos notebooks. São leituras do redator ao consolidar o material.

1. **A divisão aleatória e a "identidade do bairro".** As variáveis de crime são constantes dentro de cada bairro. Como a divisão é aleatória, transações do mesmo bairro aparecem no treino e no teste. Isso ajuda a explicar por que o crime captura 97% do teto: o modelo reconhece bairros. O R² de teste mede a capacidade de prever preços **em bairros já vistos**, e não em bairros novos. Uma validação por bairros inteiros deixados de fora (por exemplo, `GroupKFold`) mediria a generalização espacial. Isso não foi feito.
2. **O HistGradientBoosting abaixo do Random Forest** (R² de validação 0,414 contra 0,443) é incomum. Os hiperparâmetros do boosting não foram ajustados; só os do Random Forest receberam tentativas de regularização.
3. **O OLS dos notebooks 03 e 05 usa erros-padrão convencionais** em unidades geográficas vizinhas, que podem ter dependência espacial. Os intervalos de confiança podem estar levemente otimistas. Mesmo assim, o p-valor de ~10⁻¹⁶ deixa folga ampla.
4. **Os pontos fortes metodológicos** do conjunto: split com teste usado uma vez; pré-processamento dentro do pipeline; comparação piso/teto de variáveis; checagem de robustez da geografia; e a descoberta do viés de geolocalização.

---

## 8. Divergências entre texto e saídas executadas

Ao cruzar o texto explicativo (markdown) dos notebooks 04 e 05 com as saídas executadas, aparecem inconsistências. **Este relatório usa as saídas executadas como fonte primária.**

**Causa provável.** Os arquivos `bairros.parquet` e `base_modelagem.parquet` foram regravados em 05/10 às 21:29. Os notebooks 04 e 05 foram salvos às 19:25–19:26 desse mesmo dia. O texto vem de uma execução ainda anterior, com 430 bairros. A execução mais recente tem 431. O notebook 03 foi salvo em 08/10 e está consistente com o cache atual. **Recomenda-se reexecutar os notebooks 04 e 05 e revisar o texto antes de qualquer apresentação.**

### 8.1 Contradição de conclusão (a mais importante)

| Item | Texto do notebook 04 | Saída executada |
|---|---|---|
| Silhueta e Ward concordam em k = 3? | "Concordam em k = 3" | Silhueta aponta k = 3, **Ward aponta k = 4** |

### 8.2 Notebook 04: números do texto contra as saídas

| Item | Texto | Saída executada |
|---|---|---|
| Nº de bairros | 430 | 431 |
| Tamanho k = 3 (cl. 0 / 1 / 2) | 244 / 109 / 77 | **247 / 109 / 75** |
| R$/m² do cluster 1 | R$ 4.017 | R$ 4.056 |
| R$/m² do cluster 2 | R$ 6.589 | R$ 6.655 |
| % violento (cl. 0 / 1 / 2) | 44% / 37% / 29% | 43,9% / 36,2% / 28,4% |
| Crimes/ano (cl. 1 / 2) | 2.586 / 236 | 2.590 / 227 |
| Silhueta em k = 3 | 0,274 | 0,276 |
| Estabilidade dos pares juntos | 98,0% | 97,2% |
| Tamanho k = 4 (cl. 2 / 3) | 121 / 22 bairros | 111 / 24 bairros (cl. 2 = 111; cl. 3 = 24) |
| Quadrante "caro + muita violência" | 60 bairros | 61 bairros |
| Concentração por quadrante (k = 3: 90% / 50% / 39%) | Aparece no texto | **Não há célula que gere esses números.** O código só cruza k = 4 |
| Concentração k = 4 (cl. 2 / 3 / 1 / 0) | 82% / 68% / 58% / 42% | 86% / 67% / 56% / 42% |
| Maior célula em "caro + muita violência" | "7 bairros" | **44 bairros** (cl. 1). Isso é ~20% do cluster, então a conclusão de que não é maioria de nenhum cluster se mantém |
| Transações no mapa | 312.322 | 312.248 |

A afirmação de que o cluster 2 de k = 3 tem "90% de concentração" não pode ser verificada com as saídas atuais. Use com cautela: o dado verificável é 86% no cluster 2 de k = 4.

### 8.3 Notebook 05: números do texto contra as saídas

| Item | Texto (veredito) | Saída executada |
|---|---|---|
| ρ crimes totais (bairro) | +0,146 | +0,150 |
| ρ crimes violentos (bairro) | +0,036 | +0,037 |
| ρ % violento (bairro) | −0,548 | −0,545 |
| ρ violentos por mil (bairro) | −0,296 | −0,283 |
| Unidades (bairro) | 430 | 431 |
| R² do OLS no bairro | 0,411 | 0,416 |
| Coeficiente OLS no bairro | −0,912 | −0,922 |
| Efeito de +10 p.p. no bairro | −8,7% | −8,8% |
| Coeficiente OLS no distrito | −1,815 | −1,816 |

As diferenças são pequenas e **não alteram nenhuma conclusão qualitativa**. Os números do bairro no texto (−0,548; −0,296; 430 bairros) coincidem com os do texto do notebook 02, que também vem da execução mais antiga.

### 8.4 Números que dependem de gráficos

Os valores da importância por permutação (~0,46 para `share_violento`; ~0,19 para `log_violentos`) e da dependência parcial (~8,5 → ~8,2) vêm só do texto do notebook 03 e não podem ser verificados em saída de texto. O quintil "R$ 5.045 → R$ 3.487" também vem do texto do notebook 03.

---

## 9. Glossário

| Termo | Significado |
|---|---|
| **ITBI** | Imposto sobre transmissão de bens imóveis. As guias pagas fornecem as transações. |
| **SQL** | Número de cadastro do imóvel (setor + quadra + lote + dígito). Não confundir com a linguagem. |
| **Setor fiscal** | Os 3 primeiros dígitos do SQL. Unidade da Prefeitura, com 309 setores. |
| **GeoSampa** | Portal de dados geoespaciais da Prefeitura de São Paulo. |
| **Modelo hedônico** | Regressão que explica o preço de um bem pelos seus atributos. |
| **Reais constantes** | Preços corrigidos pela inflação (IPCA) para um mês de referência (2026-08). |
| **Composição** (do crime) | Fração das ocorrências que é violenta (`share_violento`). |
| **Exposição** | Volume absoluto de ocorrências. Cresce com a circulação de pessoas. |
| **R²** | Parcela da variância do alvo explicada pelo modelo. |
| **RMSE / MAE** | Erros médios (quadrático e absoluto). Aqui estão em escala log. |
| **Erro mediano** | Erro relativo típico: mediana de \|exp(previsto − real) − 1\|. |
| **Overfitting** | Modelo que vai muito melhor no treino do que em dados novos. |
| **Vazamento de dados** | Informação de validação ou teste influenciando o treino. |
| **Permutação (importância)** | Queda do R² ao embaralhar uma variável. |
| **Dependência parcial** | Forma média do efeito de uma variável sobre a previsão. |
| **Spearman ρ** | Correlação de ordens. Robusta a outliers. |
| **Pearson r** | Correlação linear. |
| **OLS** | Regressão linear por mínimos quadrados. |
| **KMeans** | Algoritmo de agrupamento por distância aos centroides. |
| **Silhueta** | Qualidade de separação dos clusters (−1 a 1). |
| **Inércia** | Soma das distâncias ao centroide. Cai sempre com mais clusters. |
| **Dendrograma de Ward** | Agrupamento hierárquico que funde os grupos mais parecidos. |
| **PCA** | Projeção que preserva a máxima variância, usada para visualizar. |
| **Efeito ecológico** | Correlações maiores quando se agregam unidades. Não vale como efeito individual. |
| **Pureza** (do mapeamento) | Fração dos pontos de um bairro que cai no distrito modal. |
| **Ponto-em-polígono** (sjoin) | Operação espacial que descobre em qual polígono cai um ponto. |

---

## 10. Reprodutibilidade

- **Arquivos:** `exploration/Russo/imoveis_sp/03_modelagem.ipynb`, `04_clusterizacao.ipynb`, `05_checagem_distrito.ipynb`.
- **Entradas (em `data/SP/cache/`):** `base_modelagem.parquet`, `bairros.parquet`, `mes_referencia.txt`, `crimes_sp.parquet`, `distritos_sp.geojson`, `setor_para_distrito.parquet`.
- **Sementes:** `random_state = 42` em todos os divisores, modelos, `KFold` do `TargetEncoder`, amostra da SVM, permutação, PCA e KMeans.
- **Ordem de execução:** 01 → 02 → 03 → 04 → 05. Os notebooks 03 a 05 só leem o cache, sem recalcular a preparação.
- **Bibliotecas principais:** scikit-learn, statsmodels, scipy, geopandas, plotly, seaborn.
- **Mapas (notebook 04):** o renderizador `plotly_mimetype` mantém o notebook em ~1,5 MB. Para exportar HTML estático, troque para `plotly_mimetype+notebook`, que embute o plotly.js e passa de 8 MB.
