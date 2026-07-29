# Decisões técnicas

Registro curto (estilo ADR) das decisões não óbvias tomadas neste projeto —
o "porquê" por trás de escolhas que, olhando só o código, poderiam parecer
arbitrárias.

## 1. Status manual sempre começa como "Automático" ao abrir a planilha

**Contexto:** a planilha de origem tem uma coluna legada de status
(`STATUS.`) preenchida manualmente ao longo do tempo. Em uso real, colunas
de status mantidas manualmente tendem a ficar desatualizadas frente à data
de fim real do contrato.

**Decisão:** ao carregar a planilha, o override manual de status
(`statusManual`) **nunca** é inicializado a partir da coluna legada — ele
sempre começa `null` (= "Automático"), e o status exibido é sempre
recalculado a partir da data de fim. O texto legado fica disponível só
como referência somente-leitura (`statusLegadoRaw`), usado para gerar um
alerta de "status manual divergente" na tela de Gestão de Contratos — não
para decidir o que é exibido.

**Por quê:** usar a coluna legada para pilotar o status exibido
reintroduziria o mesmo problema (status desatualizado) que a régua
automática foi criada para resolver. Tratar a divergência como um item
acionável de qualidade de dados (em vez de propagar o valor antigo
silenciosamente) é mais seguro.

## 2. Backup é 1 cópia só, sobrescrita a cada salvamento

**Decisão:** "Salvar" grava o estado *anterior* ao save em
`<nome>.backup.xlsx`, sobrescrevendo o backup anterior. Não é um
histórico de versões.

**Por quê:** implementar histórico de versões completo exigiria um backend
ou integração com um sistema de controle de versão externo — fora do
escopo de um app client-side de arquivo único. Uma cópia de segurança
simples cobre o caso de uso mais comum (desfazer o último salvamento
acidental) sem essa complexidade. Times que precisem de histórico completo
devem versionar o `.xlsx` externamente (ex.: OneDrive/SharePoint com
histórico de versões, ou um repositório Git).

## 3. Mitigação de injeção de script na geração do snapshot

**Contexto:** o snapshot HTML embute dados de contratos digitados por
usuários (incluindo campos de texto livre como "Comentários") diretamente
no HTML gerado. Um comentário contendo o texto literal `</script>`
fecharia a tag prematuramente e quebraria (ou, no pior caso, permitiria
injetar) o script gerado.

**Decisão:** todo o conteúdo embutido no snapshot (CSS, bibliotecas JS,
lógica de negócio, e o payload de dados dos contratos) é serializado em
**base64** antes de ser embutido, e decodificado por um pequeno bootstrap
em runtime. Base64 nunca contém os caracteres `<` ou `>`, então nenhum
conteúdo de usuário pode fechar uma tag `<script>` prematuramente. O
payload de dados também passa por escape de caracteres antes da
serialização, como camada adicional.

**Por quê:** essa abordagem evita a necessidade de sanitizar/whitelistar
cada campo de texto livre individualmente, e funciona de forma uniforme
para qualquer dado que um usuário venha a digitar no formulário de
contrato.

## 4. Exclusão de contrato não remove a linha física

**Decisão:** "Excluir" um contrato limpa apenas as colunas gerenciadas da
linha correspondente (inclusive a razão social); a linha em si continua
existindo na planilha, e o parser a ignora em leituras futuras porque a
razão social vazia é o sinal de "linha vazia/pular".

**Por quê:** a planilha tem colunas não gerenciadas pelo app (totais
anuais, séries mensais, outras abas) que dependem do alinhamento de linha
com outras ferramentas externas (ex.: um modelo de BI). Remover a linha
física desalinharia essas colunas nas linhas vizinhas em algumas versões
de planilha/ferramenta consumidora. Limpar os campos gerenciados sem tocar
na posição da linha é mais seguro para interoperabilidade.
