Este projeto tem como objetivo **automatizar a transformação de uma planilha bruta** (com dados de pré-agendamentos ou atendimentos) em uma **planilha formatada e organizada** por profissionais de saúde, horários e tipos de atendimento.

O código foi desenvolvido em **Python**, utilizando a biblioteca **Pandas**, e gera uma planilha Excel com colunas padronizadas, horários ordenados e seções visualmente separadas entre profissionais.

---

## Funcionalidades Principais

- **Remove as 7 primeiras linhas** da planilha original (normalmente contendo cabeçalhos ou informações irrelevantes);
- **Ordena os atendimentos** por profissional e horário;
- **Agrupa profissionais de forma sequencial**, com **duas linhas em branco** entre cada profissional;
- **Mantém as informações do Biopsicossocial** (Psicologia, Serviço Social e Fisioterapia) vinculadas ao perito principal;
- Adiciona colunas padrão com espaços para preenchimento manual;
- **Simula caixas de seleção (checkboxes)** nas colunas `C` e `NC` utilizando o símbolo `☐`;
- Exporta automaticamente o resultado em um novo arquivo Excel (`planilha_transformada.xlsx`).

---

## Estrutura do Código

### 1. Importações
```python
import pandas as pd
