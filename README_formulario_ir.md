# 🧾 Formulário de Coleta de Dados para Declaração de Imposto de Renda

Formulário em planilha, dividido em etapas, para **coletar e organizar as informações que um contribuinte precisa reunir para a declaração do Imposto de Renda (IRPF)**: dados do titular, saldos e informes dos bancos e rendimentos recebidos como freelancer/holerite.

> **Plataforma:** Microsoft Excel (Excel Online / Excel para a web)
> **Arquivo:** `atv_Forms_DIO_RL_1.xlsx`
> **Tipo de projeto:** formulário de entrada de dados com validação e navegação entre etapas, feito só com recursos nativos do Excel (sem macros)

> ℹ️ **Sobre o escopo:** a planilha **não emite nem calcula imposto**. Ela é um formulário de **coleta e conferência de dados** que antecede a declaração. Todos os dados de exemplo (nome, CPF, e-mail, valores) são **fictícios**.

---

## 🎯 Objetivo e problema que resolve

Preencher a declaração de IR exige reunir dados espalhados (documentos pessoais, informes de vários bancos, comprovantes de renda). O formulário organiza essa coleta em **3 etapas guiadas**, padroniza as respostas e já mostra o total consolidado dos informes bancários.

| Etapa | Aba | Pergunta que responde |
|---|---|---|
| 1 | `Titular` | Quem é o contribuinte e qual é a situação cadastral dele (cônjuge, dependente, residência no exterior, alterações desde a última entrega)? |
| 2 | `Informes` | Em quais bancos o contribuinte tem saldo e qual é o valor total? |
| 3 | `Notas` | Quais rendimentos (freelancer ou holerite) foram recebidos, quando e de quanto? |

---

## 🗂️ Estrutura do arquivo

| Aba | Visibilidade | Função |
|---|---|---|
| `Titular` | Visível | Etapa 1: dados pessoais e de contato. |
| `Informes` | Visível | Etapa 2: até 3 bancos, com valor atual e nome do comprovante. Mostra o total no topo. |
| `Notas` | Visível | Etapa 3: tabela de lançamentos (data, tipo e valor). |
| `Bancos_Apoio` | **Oculta** | Lista de 50 bancos (`código - nome`) que alimenta a lista suspensa da aba `Informes`. |

---

## 🧮 Detalhes técnicos: o que foi usado e por quê

### 1. Colunas "RESPOSTAS" e "RESPOSTAS FORMATADAS" + função `TEXTO` (`TEXT`)

```excel
CPF:               =TEXT(D7,  "000\.000\.000-00")      → 123.456.789-01
Título de eleitor: =TEXT(D8,  "0000"&" "&"0000"&" "&"0000") → 1234 5678 9012
CEP:               =TEXT(D12, "00000-000")             → 01310-100
Celular:           =TEXT(D14, "(00) 00000-0000")       → (11) 98765-4321
```

- **Por que duas colunas:** a coluna `RESPOSTAS` recebe o que o usuário digita (só números, sem se preocupar com pontos e traços). A coluna `RESPOSTAS FORMATADAS` devolve o dado já no formato oficial.
- **Por que `TEXT` e não só formatação de célula:** CPF, CEP e telefone digitados como número **perdem os zeros à esquerda** (um CEP `01310100` vira `1310100`). A máscara `00000-000` do `TEXT` **repõe os zeros** e insere a pontuação. O resultado é um texto pronto para copiar para outro documento.
- **Campos de texto simples** (nome, rua, e-mail, cônjuge) usam `=D6`, só para manter a coluna formatada completa e alinhada com a de respostas.

### 2. Validação de dados: lista suspensa

| Onde | Opções | Tipo de lista |
|---|---|---|
| `Titular!D16:D18` | `Sim, Não` | Lista digitada na própria regra |
| `Notas!D6:D25` | `OLHERITE, FREELANCER` | Lista digitada na própria regra |
| `Informes!D10, D16, D22` | 50 bancos | Lista vinda de **outra aba** (`Bancos_Apoio!$A$2:$A$51`) |

- **Por que lista suspensa:** padroniza a resposta (sem "sim", "SIM", "s"), elimina erros de digitação e acelera o preenchimento.
- **Por que a lista de bancos fica em uma aba separada e oculta:** são 50 itens, grandes demais para digitar dentro da regra. Em uma aba de apoio, a lista pode ser mantida sem tocar na regra de validação. Ocultar a aba mantém a interface limpa.
- **Formato `código - nome`:** reproduz o padrão de identificação bancária, o que reduz ambiguidade entre bancos com nomes parecidos.

### 3. Função `SOMA` (`SUM`): total dos informes

```excel
=SUM(D11, D17, D23)      → R$ 9.005.405,00
```

- **O que faz:** consolida o valor atual dos três bancos em um único total, exibido em destaque no topo da aba `Informes`.
- **Por que no topo:** o total é o dado de resumo; colocá-lo antes dos detalhes responde logo à pergunta "quanto tenho no total?".
- **Por que referências individuais (`D11, D17, D23`) e não um intervalo:** os valores estão em blocos separados por linhas de rótulo (banco, anexo). Somar um intervalo contínuo incluiria textos e linhas vazias.

### 4. Hiperlinks internos: navegação entre etapas (`PRÓXIMO ->` / `<- ANTERIOR`)

| Botão | Aba | Destino |
|---|---|---|
| `PRÓXIMO ->` | `Titular` | nome definido `tbl_informes` |
| `<- ANTERIOR` / `PRÓXIMO ->` | `Informes` | `tbl_titular` / `tbl_notas` |
| `<- ANTERIOR` | `Notas` | `tbl_informes` |

- **Por que hiperlinks e não botões com macro:** o arquivo foi criado no Excel para a web, onde **macros VBA não podem ser criadas, executadas nem editadas**. O hiperlink para um nome definido é a forma de dar navegação "de assistente" (wizard) sem VBA.
- **Por que apontar para nomes (`tbl_titular`, `tbl_informes`, `tbl_notas`) e não para endereços de célula:** se o formulário crescer e o intervalo mudar, basta atualizar o nome em um só lugar, e todos os links continuam corretos.

### 5. Tabela do Excel (`Tabela1`) na aba `Notas`

- **O que é:** o intervalo `C5:E25` (DATA, TIPO, VALOR) é uma **Tabela do Excel** com estilo `TableStyleLight17` e linhas zebradas.
- **Por que tabela:** a leitura de uma lista de lançamentos é mais rápida com linhas alternadas, e a tabela traz filtros no cabeçalho e cresce automaticamente quando novas linhas são adicionadas abaixo.
- **A lista suspensa `OLHERITE / FREELANCER`** cobre o tipo de cada lançamento (`D6:D25`).

### 6. Nomes definidos

| Nome | Intervalo |
|---|---|
| `tbl_titular` | `Titular!C5:E18` |
| `tbl_informes` | `Informes!C5:D24` |
| `tbl_notas` | `Notas!C5:E25` |

Servem de **âncora de navegação** dos botões e documentam o que cada bloco contém.

### 7. Hiperlink externo (`LINKEDIN`)

Célula no canto de cada aba que leva ao perfil do autor. Identifica a autoria do projeto e serve de contato quando o arquivo é compartilhado.

---

## 🎨 Decisões de design e usabilidade

| Recurso | Motivo |
|---|---|
| **Uma etapa por aba, numeradas (1, 2, 3)** | Reduz a carga mental: o usuário só vê o que precisa preencher naquele passo. |
| **Título grande em roxo com borda dupla + subtítulo em fundo rosa-claro** | Diz em que etapa o usuário está e o que se espera dele ("Preencha os dados da sua Pessoa Titular"). |
| **Cabeçalho roxo com texto branco** (`DADOS DO TITULAR` / `RESPOSTAS` / `RESPOSTAS FORMATADAS`) | Separa claramente a pergunta, a resposta digitada e o resultado formatado. |
| **Rótulos em negrito e fonte menor; campos de resposta com fundo colorido e borda fina** | Deixa evidente **onde digitar** e diferencia o rótulo do campo. |
| **Blocos "1° Banco" repetidos na vertical** | Cada banco ocupa um bloco idêntico, o que facilita o preenchimento e a leitura. |
| **Formato de moeda `R$ #,##0.00`** nos valores bancários | Valores monetários legíveis e consistentes. |
| **Colunas fora da área do formulário ocultas** (largura padrão `0`) | A tela mostra só o formulário, sem a grade infinita da planilha. |
| **Faixa roxa "CRIADO POR" e botão `LINKEDIN`** | Identidade visual e autoria do projeto. |

---

## ⚠️ Limitações do Excel Online (Excel para a web) por recurso

> Esta seção separa o que **está documentado pela Microsoft** do que é **observação sobre o próprio arquivo**, para não atribuir ao Excel Online limitações que são só de desenho da planilha.

| Recurso usado | Limitação / ponto de atenção | Origem |
|---|---|---|
| **Navegação por hiperlink** (`PRÓXIMO`, `ANTERIOR`) | O Excel para a web **não cria, executa nem edita macros VBA**. Por isso não há botões inteligentes (ex.: "Limpar formulário", "Enviar", "Validar se tudo foi preenchido"). A navegação ficou restrita a links. | Documentação Microsoft |
| **Validação de dados (listas suspensas)** | A validação funciona como no desktop, **exceto o modo IME** (entrada em idiomas asiáticos). Há relatos de usuários na Comunidade Microsoft de que a edição de fórmulas de lista é mais rígida na versão web (por exemplo, exigir referências absolutas como `$A$3`). | Documentação Microsoft (IME); relato de usuários (edição) |
| **Lista de bancos em outra aba oculta** | A lista de outra aba funciona, mas é **mais prático criá-la/editá-la no Excel desktop** e abri-la no web, dada a rigidez de edição relatada acima. Além disso, a validação só barra a digitação; **valores colados** sobre a célula não são verificados. | Relato de usuários; comportamento geral de validação |
| **Proteção de planilha** | Segundo a documentação, planilhas protegidas só podem ser **visualizadas** no Excel para a web. Por isso o arquivo **não está protegido**: as fórmulas das colunas `RESPOSTAS FORMATADAS` e o `TOTAL` podem ser sobrescritos por engano. | Documentação Microsoft |
| **Tabela do Excel (`Tabela1`)** | Tabelas são exibidas no navegador de forma semelhante ao desktop, sem limitação relevante para este uso. O ponto de atenção é do arquivo: o nome `tbl_notas` cobre um intervalo **fixo** (`C5:E25`) e **não acompanha** o crescimento da tabela. | Documentação Microsoft; observação do arquivo |
| **`TEXT` com máscaras** | A documentação consultada **não lista limitação específica da função `TEXT` na versão web**. O que vale é uma regra geral: `TEXT` devolve **texto** (não serve para cálculos) e **só formata números**. Se o usuário digitar `123.456.789-01` com pontuação, a célula vira texto e a máscara não é aplicada. | Comportamento geral da função |
| **Campo `ANEXO 🔗`** | O Excel não faz upload de arquivos para dentro da célula. O campo guarda **apenas o nome** do comprovante (texto), sem link nem arquivo anexado. Para vincular de verdade seria preciso inserir um hiperlink para o arquivo no OneDrive. | Observação do arquivo |
| **Aba oculta e colunas ocultas** | Ocultar é só visual: **não é segurança**. Qualquer pessoa com acesso ao arquivo consegue reexibir a aba `Bancos_Apoio` e os dados nela. | Comportamento geral do Excel |
| **Dados sensíveis (CPF, título de eleitor)** | A planilha **não criptografa** nem mascara dados. Compartilhar o arquivo (OneDrive, GitHub) expõe tudo o que estiver preenchido. | Comportamento geral do Excel |

---

## 🔧 Pontos de melhoria identificados

Registrados por transparência, para quem avaliar o projeto:

- **Máscara do telefone fixo incorreta:** o número `1132511000` (10 dígitos) usa a máscara de celular `(00) 00000-0000` e aparece como `(01) 13251-1000`. Para telefone fixo, a máscara deveria ser `(00) 0000-0000` (resultado: `(11) 3251-1000`).
- **Rótulo repetido:** os três blocos da aba `Informes` se chamam "1° Banco". O segundo e o terceiro deveriam ser "2° Banco" e "3° Banco".
- **Erros de digitação nos rótulos:** `TITTULAR`, `CÔJUGE`, `DEPEDENTE CÔNJUGE` e `OLHERITE` (o correto é `HOLERITE`, inclusive na lista suspensa).
- **Lançamento possivelmente duplicado:** a aba `Notas` tem duas linhas com a mesma data (`09/04/2024`, linhas 17 e 25), e todos os valores são a fórmula `=1000` em vez de valores digitados.
- **Fórmula residual na tabela:** a coluna `DATA` da `Tabela1` tem uma fórmula de coluna calculada (`COUNTBLANK`) que não é usada.
- **Nome definido quebrado:** `exemplo_tbl` aponta para `#REF!` e pode ser removido.
- **Validação sobrando:** `Informes!D19` tem uma lista `Sim/Não` em uma célula vazia, sem função no formulário.
- **Células auxiliares esquecidas:** o texto `Sim` aparece em colunas ocultas (`V6`, `V7`) nas abas.
- **Total e quantidade de bancos fixos:** o `TOTAL` soma exatamente três células, então adicionar um 4º banco exige alterar a fórmula.
- **Sem validação de formato dos dados:** nada impede um CPF com 10 dígitos, um CEP incompleto ou um e-mail sem `@`.

---

## ▶️ Como usar

1. Abra `atv_Forms_DIO_RL_1.xlsx` no Excel (web ou desktop).
2. Na aba **Titular**, preencha a coluna `RESPOSTAS` (só números em CPF, CEP e telefones) e escolha `Sim`/`Não` nas listas. Confira a coluna `RESPOSTAS FORMATADAS`. Clique em **PRÓXIMO ->**.
3. Na aba **Informes**, escolha o banco na lista, informe o valor atual e o nome do comprovante. O **TOTAL** no topo se atualiza sozinho. Clique em **PRÓXIMO ->**.
4. Na aba **Notas**, registre cada lançamento (data, tipo, valor).

---

## 🧰 Habilidades demonstradas

`Desenho de formulários em planilha` · `Validação de dados (listas e listas entre abas)` · `Função TEXT e máscaras de formatação` · `Hiperlinks internos e nomes definidos` · `Tabelas do Excel` · `Padronização de entrada de dados` · `Design de interface em planilha` · `Revisão crítica e documentação de limitações`

---

## 📚 Fontes consultadas

- [Differences between using a workbook in the browser and in Excel (Microsoft)](https://support.microsoft.com/excel/differences-between-using-a-workbook-in-the-browser-and-in-excel)
- [Work with VBA macros in Excel for the web (Microsoft)](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)
- [Limitations of web version Excel data validation (Microsoft Community)](https://techcommunity.microsoft.com/discussions/excelgeneral/limitations-of-web-version-office-365-excel-data-validation/4285683)

[LINKEDIN:](https://www.linkedin.com/in/fernando-henrique-935695323/)
