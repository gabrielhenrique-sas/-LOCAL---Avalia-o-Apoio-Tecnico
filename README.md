# Relatórios de Avaliação do Apoio Técnico (calendário → semana → relatório → e-mail)

Projeto único que junta o antigo `atm_sync` e o `envio_avaliacoes`.

## Fluxo
1. Você abre a página e escolhe a semana (segunda a sexta). Por padrão é a semana atual.
2. A cada abertura, recarga ou clique em **↻ Atualizar**, o programa consulta **o calendário online** e **as respostas do formulário**.
   O calendário é baixado só para a memória (nada de cache em disco; a planilha online nunca é alterada).
3. Os ATMs da semana são organizados (município oficial, gerências válidas, modalidade) e gravados no `CONTROLE.xlsx`.
4. A tela mostra cada atividade com **modalidade**, gerências, **quantidade de avaliações** (e as do dia seguinte, separadas) e o **status de envio**
   (verde, com data e hora, se já foi enviado).
5. **Preparar envio** (ou **Reenviar**) gera o PDF, mostra o e-mail para você editar e envia. Também dá para baixar o PDF e um rascunho `.eml`.

Mesmo município em vários dias da semana: marque as linhas e use "Preparar as selecionadas" para um único relatório, com uma seção por data.

## Instalação
Python 3.9+. No Windows, dê dois cliques em `iniciar.bat` (instala o que falta e abre a página). No Linux/macOS: `./iniciar.sh`.
Ou: `pip install -r requirements.txt` e `python app.py` (abre http://127.0.0.1:5000).

## Configuração (`config.json`)
| Campo | O que é |
|---|---|
| `calendario.fonte` | link do Google Planilhas do calendário (compartilhado como "Qualquer pessoa com o link", leitor) ou caminho de um .xlsx |
| `respostas.*` | planilha de respostas do Forms (Google privado, link público ou arquivo). Para a conexão privada rode `python conectar_google.py` uma vez (passo a passo no README antigo, mantido abaixo) |
| `controle.arquivo` | `CONTROLE.xlsx` (criado sozinho na 1ª abertura) |
| `controle.modelo` | planilha de controle ANTIGA: só na criação, copia dela as abas Gerências (e-mails) e as mensagens de Configurações. Pode apagar depois |
| `incluir_dia_seguinte`, `campo_data_busca` | mesmas regras de antes |
| `email.*` | remetente e SMTP. **A senha não fica no arquivo**: defina a variável de ambiente `SMTP_SENHA` |

Windows: `setx SMTP_SENHA "sua-senha-de-app"` e abra um novo terminal (ou reinicie o `iniciar.bat`).

## CONTROLE.xlsx (criado e mantido pelo programa)
- **Atividades**: os ATMs lidos do calendário + status do envio, nº de respostas, data/hora, link do PDF. Linhas amarelas = aguardando revisão.
- **Gerências**: nome, gerente, e-mail, ativa (Sim/Não). Você edita.
- **Configurações**: e-mail "Para", nome do remetente e textos do e-mail (vazio = texto padrão).
- **Log**: um registro por envio. **Mudanças do calendário**: novo / alterado / removido / reativado / confirmado.
Feche o arquivo no Excel antes de usar o programa (ele grava nele). Antes de cada gravação é feito um backup em `saida/backups/` (os 20 mais recentes).

## Revisão
Aba **Revisão** na página:
- **Atividades para revisar**: ATMs do calendário sem município reconhecido ou sem gerência válida. Corrija data, município, modalidade e gerências e clique em
  **Confirmar**; só então entram na lista da semana. A correção vale enquanto o texto daquela célula no calendário não mudar (se mudar, volta para a revisão).
- **Avisos** (só leitura): ATM em fim de semana (ignorado), avaliações sem atividade correspondente, avaliações que chegaram depois do envio, gerência sem e-mail.

## Regras principais
- Só segunda a sexta. ATM em sábado/domingo é ignorado e vira aviso.
- Avaliações de uma atividade: mesmo município e mesma data do apoio. Com `incluir_dia_seguinte`, entram também as do dia seguinte, exceto se o município tiver
  outra atividade nesse dia.
- Gerências aceitas: GEBEN, GEPSB, GEPMC, GEFEAS, VSA, CAD, GETEP, GEPAC (apelidos GFEAS→GEFEAS, GEPBS→GEPSB; GVINF → CAD e VSA).
- Município no mesmo dia = uma só atividade (gerências e modalidades unidas).
- Proteção: se a leitura do calendário vier vazia ou com menos da metade dos ATMs já cadastrados na semana, nada é alterado e aparece o aviso
  (com a opção "Aceitar a leitura mesmo assim").
- Se o calendário ou o formulário estiverem fora do ar, a tela avisa e mostra os últimos dados salvos.

## Testes
`python testes.py` (usa os dados de `dados/` e uma pasta temporária; não mexe nos seus arquivos).

## Estrutura
```
app.py            página web
config.json       configurações
nucleo/calendario.py   download (em memória) e leitura do calendário
nucleo/sync.py         mescla o calendário com as atividades (novo/alterado/removido/manual)
nucleo/controle.py     CONTROLE.xlsx
nucleo/respostas.py    respostas do Forms        nucleo/google_planilhas.py  acesso privado ao Google
nucleo/fluxo.py        semana, revisão, preparo e registro do envio
nucleo/relatorio.py    gráficos e PDF            nucleo/email_util.py        e-mail
templates/ static/ assets/ dados/ municipios_pe.txt
```

## Conexão privada com o Google (respostas)
1. Em console.cloud.google.com crie um projeto, ative a **Google Sheets API** e, na tela de consentimento OAuth, escolha usuário **Interno**.
2. Credenciais > ID do cliente OAuth > **Aplicativo para computador**. Baixe o JSON como `credentials.json` nesta pasta.
3. Preencha `respostas.google_sheet_id` (trecho entre `/d/` e `/edit`) e `google_gid` no `config.json`. Na planilha: Arquivo > Configurações > Região = Brasil.
4. Rode `python conectar_google.py`, autorize no navegador. Não compartilhe `credentials.json` nem `token.json`.
