# Guia de Testes por Spec

Roteiro para testar o `Nano_Cores_CoreEngine.xlsm` e o `Demonstrativo de Precificacao.xlsm` sempre que uma spec nova for implementada.

**Princípio:** a fonte da verdade é `src/`. Os `.xlsm` são só hospedeiros e ficam fora do git (`.gitignore`). Por isso todo teste começa instalando o código de `src/` numa **cópia de teste** do arquivo certo, nunca no arquivo de produção.

---

## 1. Em qual arquivo a spec roda

Cada spec declara o hospedeiro no campo `host_file` do cabeçalho (`.agent/specs/spec-NNN.md`).

| Arquivo | Specs |
|---|---|
| `Nano_Cores_CoreEngine.xlsm` | 001, 002, 003, 006, 007, 012, 013 (e 008 até a correção D7 da spec-012) |
| `Demonstrativo de Precificacao.xlsm` | 004, 005, 009, 010, 011 |

Os módulos de `src/vba/shared/` (por exemplo `mod_AppExecution`, `mod_Config`, `mod_AdoExcel`, `tst_RunAll`) vão para **os dois** arquivos.

Specs que ligam os dois arquivos (por exemplo a 012, em que o Engine cria e gerencia demonstrativos, com contratos via `Application.Run`) exigem também o teste conjunto do passo 6.

## 2. Montar a pasta de teste

1. Use o **Excel do Windows**. ADO/OLEDB e as DLLs da spec-013 não funcionam no Excel do Mac.
2. Copie o arquivo base para uma pasta de teste:
   - Engine: `dist/Nano_Cores_CoreEngine.xlsm`
   - Demonstrativo: `assets/templates/@Demonstrativo de precificação.xlsm`
3. Crie uma pasta `data\` ao lado da cópia. `Config_ResolveDatabaseFile` procura as bases em `ThisWorkbook.Path\data\` (entre outros caminhos).
4. Gere as bases sintéticas da spec. Os `.xlsx` não vão para o git, então recrie sempre:

   ```bash
   python3 tests/fixtures/gen_mock_custos.py
   ```

   Cada spec que precisa de dados traz o próprio `tests/fixtures/gen_*.py`. Testes de integração usam **apenas** essas bases sintéticas, nunca as de produção.

## 3. Ver o que a spec mudou

```bash
git show --stat HEAD
```

ou, para o branch inteiro:

```bash
git diff main --stat
```

Só os arquivos listados precisam ser reinstalados.

## 4. Instalar o código (Editor VBA, Alt+F11)

| Tipo | Como instalar |
|---|---|
| `.bas`, `.cls` | **Remova** o módulo antigo de mesmo nome (senão o Excel cria `mod_X1` e você testa o código velho) e use Arquivo > Importar. |
| `sheet_*.vb` | Cole no módulo da aba indicada no cabeçalho (`Instalação:`), substituindo todo o código. |
| `workbook_*.vb` | Cole no `ThisWorkbook`. Se o evento já existir (por exemplo `Workbook_BeforePrint`), acrescente só a chamada. |
| `frm_*.vb` | Inserir > UserForm, mude o (Name) conforme o cabeçalho e cole o código. Os controles são criados em tempo de execução, então não há `.frx`. |

Depois:

- **Depurar > Compilar VBAProject**: precisa terminar sem nenhum erro.
- **Ferramentas > Opções > Geral > "Interromper em erros não tratados"**. Com "Interromper em todos os erros", os testes que verificam erros esperados param no depurador.

## 5. Rodar os testes automáticos (Janela Imediata, Ctrl+G)

### Tudo de uma vez

```
RunAllTests
```

O `tst_RunAll` descobre o host pelo nome do arquivo (`coreengine` ou `demonstrativo`) e roda só os testes desse host. Para forçar o host:

```
RunAllTests "ENGINE"
RunAllTests "DEMONSTRATIVO"
```

No fim ele imprime um resumo, uma linha por teste:

- **OK**: rodou sem erro e restaurou `ScreenUpdating`, `EnableEvents` e `Calculation`.
- **FALHA**: o teste levantou erro (asserção reprovada ou erro de execução) ou deixou o estado do Application alterado. O estado é restaurado antes do próximo teste.
- **AUSENTE**: a macro não existe no arquivo, ou seja, o módulo de teste não foi importado.

Se houver falha ou ausência, o `RunAllTests` levanta um erro no fim, para não passar despercebido. Linhas `[SKIP]` ou `[WARN]` na saída indicam testes que rodaram sem encontrar a base. Confira essas linhas antes de dar a spec por testada.

### Um teste de cada vez

| Host | Ponto de entrada | Spec |
|---|---|---|
| Ambos | `Test_AppExecution_SafeZone` | 003 |
| Engine | `RunAllCostEngineContractTests` | 003 |
| Engine | `Test_CostEngine_Persistence_SafeZone` | 003 |
| Demonstrativo | `RunCostPackageContractTests` | 009 |
| Demonstrativo | `RunCostPackageIntegrationTests` (aceita o caminho da base sintética como argumento) | 009 |

Ordem recomendada: contratos (em memória), depois integração (base sintética), depois regressão das specs anteriores do mesmo arquivo. O `RunAllTests` já segue essa ordem.

## 6. Testes manuais: critérios de aceite

Abra a spec e percorra cada checkbox de **Critérios de aceite** usando a interface de verdade (botões, formulários, impressão). Exemplo da spec-009:

- escolher o pacote no `frmPacoteCustos`;
- digitar ou colar um código na `CfPcfic` e conferir o custo do pacote ativo;
- usar o botão Atualizar;
- abrir a Visualização de Impressão e conferir o cabeçalho `Custos: [pacote]`.

**Specs que ligam os dois arquivos:** abra o Engine de teste e deixe **ele** criar ou abrir o Demonstrativo de teste. Confira se as chamadas `Application.Run` chegam e se as respostas voltam.

## 7. Fechar o ciclo

1. Se você corrigiu algo direto no Editor VBA durante o teste, **leve a correção de volta para `src/`**. Se não fizer isso, a próxima instalação apaga a correção.
2. Marque os critérios atendidos na spec.
3. Faça o commit.

---

## Checklist para quem implementa uma spec nova

- [ ] Módulo `tst_<Assunto>_ContractTests.bas` com um `Public Sub Run...` sem argumentos obrigatórios, que levanta erro quando alguma asserção falha (padrão da skill `code-spec-validator`).
- [ ] Testes de integração usando uma base sintética gerada por `tests/fixtures/gen_*.py`.
- [ ] **Uma linha nova em `RegistroDeTestes`** (`src/vba/shared/tests/tst_RunAll.bas`), no formato `"HOST|Macro|spec-NNN|Descrição"`, e uma linha na tabela do passo 5 deste guia.
- [ ] Linha `Instalação:` no cabeçalho de todo `.vb`.
