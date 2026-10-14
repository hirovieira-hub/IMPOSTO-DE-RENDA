# IMPOSTO-DE-RENDA
# 📊 Agregador de Dados para Imposto de Renda no Excel

Projeto do desafio da **DIO**: uma ferramenta em Excel para organizar e reunir as informações essenciais da declaração de imposto de renda, com entradas validadas, menus de navegação e links rápidos.

## 🎯 Objetivo

Criar uma planilha prática e robusta em que o usuário controla seus dados com segurança, sem precisar de programação, usando apenas recursos nativos do Excel.

## 🧰 Recursos utilizados

| Recurso | Para que serve |
|---|---|
| Tabelas (`Ctrl + T`) | Organizar dados com filtros e fórmulas que se expandem sozinhas |
| Validação de dados | Listas suspensas e bloqueio de entradas inválidas |
| Formatação condicional | Destacar erros e valores importantes |
| Hiperlinks e botões | Navegação entre abas e links rápidos |
| Funções de data | `DATADIF`, `FIMMÊS`, `ANO`, `MÊS`, `DIA`, `TEXTO` |
| Tratamento de erros | `SEERRO`, `ÉNÚM`, `SE` |

## 🗂️ Estrutura da planilha

- **Menu**: página inicial com botões de navegação.
- **Vendas**: registro de vendas com total calculado e data desmembrada (ano, mês, dia e dia da semana).
- **DATADIF**: tabela de ajuda com as unidades de diferença entre datas.
- **Envios**: controle de envios com cálculo do fim do mês, diferença de dias e identificação de registros com erro.

## 🧮 Principais fórmulas

**Total da venda**
```excel
=B2*C2
```

**Desmembrando a data**
```excel
=ANO(K2)
=MÊS(K2)
=DIA(K2)
=TEXTO(K2;"dddd")
```

**Diferença entre datas**
```excel
=DATADIF(B2;C2;"D")
```

| Unidade | Significado |
|---|---|
| `D` | Dias entre as duas datas |
| `M` | Meses completos |
| `Y` | Anos completos |
| `MD` | Dias, ignorando meses e anos |
| `YM` | Meses, ignorando anos |
| `YD` | Dias, ignorando anos |

> Entre 12/02/2023 e 22/04/2025: D = 800, M = 26, Y = 2, MD = 10, YM = 2, YD = 69.

**Fim do mês e detecção de erros**
```excel
=SEERRO(FIMMÊS(B2;0);"")
=SE(ÉNÚM(B2);SE(B2<DATA(2000;1;1);"Erro: data inválida";"OK");"Erro: não é data")
```

## 🖼️ Capturas de tela

### Menu principal
![Menu principal](images/menu.png)

### Tabela de vendas
![Tabela de vendas](images/vendas.png)

### Diferença entre datas
![DATADIF](images/datadif.png)

### Controle de envios com erros destacados
![Envios](images/envios.png)

## 🚀 Como usar

1. Baixe o arquivo `.xlsx` deste repositório.
2. Abra no Excel e habilite a edição.
3. Navegue pelo menu e preencha apenas as células de entrada.
4. Confira os campos marcados com erro antes de usar os dados na declaração.

## 📚 Aprendizados

- Dados em formato de **Tabela** facilitam a manutenção e evitam erros de intervalo.
- Funções de data exigem atenção: textos e datas inválidas geram erros e precisam ser tratados com `SEERRO` e `ÉNÚM`.
- A função `DATADIF` não aparece na lista de sugestões do Excel, mas funciona normalmente.
- Documentar o processo no GitHub ajuda a consolidar o conhecimento e a mostrar o trabalho.

## 🔗 Referências

- [Documentação do GitHub](https://docs.github.com/)
- [Guia de Markdown do GitHub](https://docs.github.com/pt/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax)
