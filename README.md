# 🤖 Robô de Lançamento de Notas Acadêmicas

Projeto da **Aula 7 – Módulo 2: Automação Desktop**.

O professor tem uma planilha com as notas finais dos alunos e precisa lançá-las num
**Sistema Acadêmico Legado** que não aceita importação de arquivos. Este robô,
feito com **Python + pyautogui**, lê o arquivo de notas e digita aluno por aluno no
sistema, como um humano faria, e no final confere se tudo foi gravado certo.

## Estrutura

```
robo-lancamento-notas/
├── sistema_academico.py     # "Portal do Professor" (sistema legado simulado, Tkinter)
├── robo_lancamento.py       # o robô (pyautogui) + conferência + relatório
├── regras.py                # regras de cálculo (pesos e média de aprovação)
├── requirements.txt
└── dados/
    ├── cadastro_alunos.csv  # base de alunos do sistema (Nome;RA)
    ├── notas_alunos.csv     # arquivo do cliente (Nome;RA;Nota01;TF01;Nota02;TF02)
    └── notas_lancadas.xlsx  # planilha gravada pelo sistema (gerada na execução)
```

## Mapeamento do processo (Fase 1)

| Passo humano | O que o robô faz |
|---|---|
| Olhar a linha no Excel | `for` sobre cada linha do `notas_alunos.csv` |
| Digitar a matrícula e buscar | `write(RA)` + `Enter` no campo de busca |
| Esperar o sistema carregar | `time.sleep(ESPERA_BUSCA)` |
| Digitar as notas | `write(nota)` + `Tab` em Nota 1, TF 1, Nota 2, TF 2 |
| Clicar em salvar | `Enter` no último campo |
| Conferir | Lê a planilha gravada e compara com o arquivo de origem |

**Regras de cálculo** (`regras.py`):

- Média do bimestre = (Nota × 8 + TF × 2) / 10
- Média final = (Média 1 + Média 2) / 2
- **APROVADO** se média final ≥ 6,0; senão **REPROVADO**

O sistema busca o aluno por **RA ou Nome** (nome parcial mostra a lista para escolher).

## Como executar

```bash
pip install -r requirements.txt

# Terminal 1 – abre o sistema
python sistema_academico.py

# Terminal 2 – roda o robô (arquivo opcional; padrão: dados/notas_alunos.csv)
python robo_lancamento.py
```

Não mexa no mouse/teclado enquanto o robô roda.
⚠️ **FAIL-SAFE:** arraste o mouse para um **canto da tela** para parar o robô na hora.

Ao terminar, o robô gera o `relatorio_execucao.txt` com total lançado, aprovados,
reprovados, tempo gasto e o resultado da conferência.

## Resultado do teste

30 alunos lançados em ~60 s (≈2 s por aluno), **zero divergências** entre o arquivo
de origem e a planilha gravada pelo sistema.

## Formato do arquivo de entrada

```
Nome;RA;Nota01;TF01;Nota02;TF02
Ana Beatriz Souza;20251001;4,5;7,5;6,5;2,5
```

Aceita vírgula ou ponto como decimal e o cabeçalho `Notao2` do enunciado.
Linhas com nota inválida (vazia ou fora de 0–10) são ignoradas e registradas no relatório.
