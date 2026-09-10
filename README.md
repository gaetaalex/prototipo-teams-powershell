# Automação Microsoft Teams — Modo Automático via PowerShell

Aplicação desktop com interface gráfica nativa em PowerShell (WPF) para envio automatizado de mensagens individuais no Microsoft Teams desktop, com leitura dinâmica, mapeamento flexível de colunas, **configurador de limites de mensagens e regras da 2ª mensagem**, **tela de resumo pós-envio com motivos de erro** e atualização automática da planilha Excel.

---

## Principais Funcionalidades

1. **Tela de Resumo e Relatório Final:**
   - Ao término do disparo (ou se o processo for interrompido), uma **janela modal de resumo** abre automaticamente com:
     - **Cards de Indicadores:** Total Processados, Enviados com Sucesso (verde) e Erros/Falhas (vermelho).
     - **Tabela Detalhada com os Motivos:** Lista cada contato (Matrícula, Nome, E-mail, Tipo de Mensagem, Resultado) e a coluna **Motivo / Detalhe** explicando exatamente por que foi concluído com sucesso ou qual foi o erro (ex.: *E-mail inválido*, *Texto de mensagem vazio*, *Janela do Teams não respondeu*, *Interrompido pelo operador*, etc.).
     - **Botão `Exportar Relatório (.csv)`:** Permite salvar todo o relatório em formato CSV compatível com Excel para auditoria ou conferência.
     - **Botão `Ver Último Relatório`:** Na barra inferior da janela principal, permite reabrir o resumo a qualquer momento.

2. **Configurador de Limite de Mensagens por Usuário:**
   - Campo **`Máx. contatos p/ pessoa`**: Você define quantas mensagens no máximo cada pessoa pode receber no total (exemplo: `1` para mensagem única, `2` para permitir até a 2ª mensagem de follow-up, `3` etc.).
   - Se a pessoa já atingiu o limite configurado, o sistema bloqueia automaticamente e sinaliza na prévia: `[Bloqueado] Max. atingido (2/2)`.

3. **Configurador de Envio da 2ª Mensagem (Follow-up):**
   - Campo **`Dias de espera p/ 2ª msg`**: Você define o intervalo mínimo em dias após o 1º contato (exemplo: `3` dias).
   - O sistema confere a coluna `ÚLTIMA DATA DE CONTATO` do Excel:
     - Se o prazo **já passou**: libera como `[Apto] 2ª Msg liberada` e envia o texto da 2ª mensagem (`TEXTO SEGUNDO CONTATO`).
     - Se o prazo **ainda não passou**: bloqueia como `[Aguardar] até DD/MM/AAAA`.

4. **Mapeamento Dinâmico de Colunas:**
   - Se os nomes das colunas da planilha mudarem (ex.: `MATRICULA`, `CHAPA`, `TEXTO CONTATO`, etc.), você pode simplesmente selecionar a coluna correspondente nos menus suspensos (comboboxes) diretamente na tela.

5. **Sem Bloqueio de TI e Sem Erro 403:**
   - Utiliza o aplicativo desktop oficial do Microsoft Teams instalado na máquina, sem necessidade de permissões no Microsoft Entra ID (Azure AD).

6. **Segurança com Backup Automático:**
   - Cria uma cópia de segurança do Excel com carimbo de data e hora antes de qualquer disparo.

---

## Como Usar

1. Dê um duplo clique em **`iniciar.bat`**.
2. Na seção **1. Seleção da Planilha Excel**:
   - Clique em **`Procurar...`** e selecione a planilha (`.xlsx` ou `.xlsm`).
   - Selecione a **Aba** e clique em **`Carregar e Mapear`**.
3. Na seção **2. Mapeamento das Colunas**:
   - Verifique ou altere as colunas desejadas (Matrícula, Nome, 1ª Mensagem, 2ª Mensagem, Status, Número de Contatos, etc.).
4. Na seção **3. Configurador de Limites e 2ª Mensagem**:
   - Ajuste o **`Máx. contatos p/ pessoa`** (ex: `2`).
   - Ajuste os **`Dias de espera p/ 2ª msg`** (ex: `3`).
   - Clique em **`Aplicar Regras`** para atualizar a prévia.
5. Verifique a tabela de **Prévia dos Contatos**:
   - A coluna **`Msg a Enviar`** indica claramente se será enviada a `1ª Msg` ou a `2ª Msg (Follow-up)`.
   - A coluna **`Situação`** informa se está Apto, Bloqueado ou Aguardando prazo.
6. Disparo e Relatório:
   - **`Testar Apenas o 1º Contato`**: Envia somente para a primeira pessoa da lista para validar.
   - **`Iniciar Envio Automático`**: Envia para todos os aptos da fila com barra de progresso.
   - **`Parar Envio`**: Interrompe a qualquer momento.
   - Ao finalizar, a **Tela de Resumo** abre automaticamente com o relatório de sucessos, falhas e motivos detalhados, com opção de exportar em CSV.
