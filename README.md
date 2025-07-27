# 🧮 Sistema de Gestão de Estoque

Aplicativo desktop desenvolvido em Python para controle e apuração de inventário, com suporte à leitura de planilhas Excel e geração de relatórios financeiros detalhados. Ideal para empresas que fazem contagens manuais e desejam digitalizar o processo de apuração de estoque.

---

## 📦 Funcionalidades

- ✅ Carregamento de planilhas `.xlsx` ou `.xls`
- ✅ Salvamento e carregamento de dados em JSON
- ✅ Cálculo automático de estoque, diferença e valores
- ✅ Edição direta de contagem via interface
- ✅ Classificação por divergência de valores (VL. DIF.)
- ✅ Filtros por código, endereço, faltas e sobras
- ✅ Resumo automático com totais, percentuais e estatísticas
- ✅ Interface amigável com suporte a navegação por abas

## 🧾 Estrutura da Planilha

A planilha a ser carregada deve conter as seguintes colunas obrigatórias:

| Coluna       | Tipo Esperado | Descrição                       |
|--------------|---------------|----------------------------------|
| COD          | Texto/Número  | Código do produto               |
| PRODUTO      | Texto         | Nome do produto                 |
| VL. UNT.     | Número        | Valor unitário do item          |
| ENDEREÇO     | Texto         | Localização física no estoque   |
| QTD          | Número        | Quantidade registrada no sistema|

As colunas abaixo são geradas automaticamente:
- `CONTAGEM`
- `VL. ESTOQUE`
- `DIF. ETQ`
- `VL. DIF.`

---

## 🛠️ Tecnologias Utilizadas

- Python 3.x
- Tkinter (Interface Gráfica)
- Pandas (Manipulação de dados)
- JSON (Exportação e Importação)
- Excel via openpyxl

---

## 💡 Exemplos de Uso

- Contagem física de inventário com preenchimento automático
- Geração de relatórios financeiros de estoque
- Identificação de perdas e sobras de forma rápida
- Exportação de inventário ajustado para auditoria

---

## 👨‍💻 Autor

**Ricardo Leffers Gomes**  
📅 Início: 01/01/2025  
🔖 Versão: 001

