# TCC_UspEsalq
Repositório com scipts utilizados para o TCC do MBA em Data Science &amp; Analytics

# Precificação Hedônica de Imóveis na Grande Florianópolis

Repositório do TCC de MBA em Data Science e Analytics — USP/ESALQ (2025).

## Objetivo
Estimar os determinantes do valor de imóveis residenciais nos municípios de 
Florianópolis, São José, Palhoça e Biguaçu, comparando regressão por MQO com 
os algoritmos Random Forest e XGBoost. A interpretação dos modelos ensemble 
foi realizada via SHAP values.

## Dados
Coletados via web scraping em março de 2026 de dois portais imobiliários 
(Crédito Real Imóveis e Gralha Imóveis). O dataset sanitizado não está 
disponível neste repositório por restrições de uso das plataformas de origem.

## Estrutura
📁 scrapers/
scraper_credito_real.py
scraper_gralha.py
📁 notebooks/
eda.ipynb
regressao_hedonica.ipynb
modelos_ml.ipynb
comparacao_segmentos.ipynb
📁 outputs/
eda_output/          → gráficos e CSVs da análise exploratória
regressao_output/    → coeficientes, VIF, diagnósticos do MQO
ml_output/           → métricas, SHAP values, comparação de segmentos

## Requisitos
Python 3.12. Instale as dependências com:
pip install pandas numpy matplotlib seaborn scikit-learn xgboost shap
statsmodels rapidfuzz requests beautifulsoup4 playwright

## Referência
Gabriel [Sobrenome]. 2025. Precificação hedônica de imóveis na Grande 
Florianópolis: uso e comparação de modelos de machine learning. 
TCC (MBA em Data Science e Analytics) — USP/ESALQ.

Todos os arquivos a subir
📁 scrapers/
   scraper_credito_real.py      ← scraper das duas etapas
   scraper_gralha.py            ← scraper via API JSON
   sanitizar_dados.py           ← script de limpeza e normalização

📁 notebooks/
   eda.ipynb
   regressao_hedonica.ipynb
   modelos_ml.ipynb
   comparacao_segmentos.ipynb

📁 outputs/
   eda_output/
      distribuicao_valor_comparativo.png
      qqplot_valor_transformacao.png
      boxplot_valor_por_regiao.png
      boxplot_valor_por_tipo.png
      boxplot_tipo_comparativo.png
      preco_m2_por_regiao.png
      preco_m2_bairros_norte.png
      heatmap_valor_tipo_regiao.png
      amenidades_impacto.png
      contagem_por_tipo.png
      contagem_por_regiao.png
      matriz_correlacao.png
      descritiva_geral.csv
      correlacao_spearman.csv
      correlacao_point_biserial.csv
      vif.csv

   regressao_output/
      coeficientes.csv
      vif_modelo.csv
      modelo_final_summary.txt
      stepwise_summary.txt
      qqplot_residuos.png
      residuos_vs_ajustados.png
      comparacao_especificacoes.csv

   ml_output/
      comparacao_modelos.csv
      comparacao_segmentos.csv
      shap_importancias.csv
      comparacao_modelos.png
      real_vs_predito.png
      shap_importancia.png
      shap_beeswarm.png
      comparacao_segmentos.png
      distribuicao_recortes.png

README.md
