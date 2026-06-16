# Guia de Configuração

[English](SETUP.md)

Este guia é detalhado de propósito, para um colega não-técnico conseguir seguir.

O calendário do Teams é o mesmo da sua caixa do Microsoft 365. Então, para os
eventos do seu Google Calendar aparecerem no Teams, o app precisa gravá-los nesse
calendário da caixa. Só existem duas formas de fazer isso, e o guia cobre as duas.

## Antes de começar

- Windows 10 ou 11.
- Sua conta institucional que existe nos dois lados, por exemplo
  `voce@seu-dominio` no Google (calendário/e-mail) e o mesmo endereço no Microsoft
  365 (Teams).
- Alguns minutos. A primeira sincronização pode levar 1 a 2 minutos.

## Qual forma eu uso?

| | Forma A: autorização do admin | Forma B: Outlook Classic |
|---|---|---|
| Outlook instalado no PC | Não precisa | Obrigatório (Classic) |
| Precisa do admin de TI | Sim, uma vez para todos | Não |
| Boa para | Muitos colegas, sem Outlook | Um PC que você controla |

- Você consegue instalar o Outlook Classic no seu PC: use a **Forma B**. Não
  precisa da TI.
- Você quer vários colegas usando e eles não têm Outlook: peça a **Forma A** para
  a TI. Uma aprovação cobre todo mundo, então ninguém fica pedindo de novo.

---

# Forma A: com autorização do admin (Microsoft Graph)

O app grava no seu calendário do Microsoft 365 pela API do Microsoft Graph, então
não precisa de Outlook no PC. O tenant da universidade bloqueia cada usuário de
aprovar isso sozinho, então precisa de UMA aprovação de um admin de TI para a
organização inteira. Depois disso, ninguém mais é solicitado.

## Parte 1 - Para o admin de TI (uma vez só, para toda a organização)

1. Acesse o centro de administração do Microsoft Entra: https://entra.microsoft.com
2. Abra **Identidade → Aplicativos → Registros de aplicativo → Novo registro**.
3. Dê um nome (ex.: `GCalSync`). Em **Tipos de conta com suporte**, escolha
   **Contas apenas neste diretório organizacional (single tenant)**. Clique em
   **Registrar**.
4. Abra o app criado → **Autenticação** → **Adicionar uma plataforma** →
   **Aplicativos móveis e da área de trabalho** → marque
   `https://login.microsoftonline.com/common/oauth2/nativeclient` e adicione a URI
   de redirecionamento `http://localhost`. Coloque **Permitir fluxos de cliente
   público** em **Sim**. Salve.
5. Abra **Permissões de API → Adicionar uma permissão → Microsoft Graph →
   Permissões delegadas →** procure e adicione **`Calendars.ReadWrite`**.
6. Clique em **Conceder consentimento de administrador para [organização]** e
   confirme. Essa é a única aprovação, e ela cobre todos os usuários.
7. Na **Visão geral** do app, copie o **ID do aplicativo (cliente)** e o **ID do
   diretório (tenant)**. Envie os dois para quem vai usar o app.

## Parte 2 - Para cada usuário

1. Instale o GCalSync (baixe o `GCalSync-Setup.exe` na
   [página de Releases](https://github.com/vinicq/gcal-to-outlook/releases) e
   execute).
2. O assistente abre. Se ele não encontrar o Outlook, pede os dados do Graph.
   Coloque o **client ID** e o **tenant ID** que o admin enviou. (Dá para editar o
   `config.json` em `%LOCALAPPDATA%\GCalSync` direto também:)
   ```json
   "microsoft": {
     "mode": "graph",
     "client_id": "CLIENT-ID-DO-ADMIN",
     "tenant_id": "TENANT-ID-DO-ADMIN"
   }
   ```
3. O navegador abre duas vezes: uma para entrar no **Google** (aprovar acesso ao
   calendário) e uma para entrar na **Microsoft**. O app não guarda senhas.
4. A janela do monitor mostra as duas contas como **Connected**. Clique em
   **Sync now**.

## Conferir

Abra o Teams → Calendário. Seus eventos do Google aparecem em um ou dois minutos.
Marque **Start automatically when Windows starts** para deixar rodando em segundo
plano.

---

# Forma B: com o Outlook Classic (sem precisar de admin)

O app grava os eventos no calendário do Outlook pelo programa local. O Teams mostra
esses eventos porque o Outlook tem a sua conta do Microsoft 365. Funciona em um PC
só, sem envolver a TI.

> **Você precisa usar o Outlook Classic, não o "novo Outlook".**
> O "novo Outlook" é outro app, baseado na web. Ele não dá ao GCalSync o acesso
> local que ele precisa, então o app não funciona com ele. Todos os passos abaixo
> assumem o Outlook Classic.

## Passo 0 - Veja qual Outlook você tem

Abra o Outlook. Olhe o canto superior direito da janela:

- Se houver um botão escrito **"Novo Outlook"** e ele estiver **ligado**,
  **desligue**. O Outlook reinicia no modo Classic.
- O Outlook Classic tem a faixa de opções antiga e um menu **Arquivo** no canto
  superior esquerdo. O novo Outlook é mais simples e tem o botão ligado.

Se você não tem o Outlook Classic, instale (próximo passo).

## Passo 1 - Instale o Outlook Classic

- Instruções (Microsoft): https://support.microsoft.com/en-US/Outlook/install-or-reinstall-classic-outlook-on-a-windows-pc
- Instalador direto (Inglês): https://go.microsoft.com/fwlink/?linkid=2276500&clcid=0x409
- Instalador direto (Português - Brasil): https://go.microsoft.com/fwlink/?linkid=2276500&clcid=0x416

Rode o instalador e abra o **Outlook (clássico)** pelo menu Iniciar. Na primeira
vez ele pode pedir para criar um perfil; aceite o padrão.

## Passo 2 - Adicione a conta Microsoft 365 (obrigatória)

É a conta cujo calendário o Teams lê. Ela precisa estar adicionada para o app
funcionar.

1. No Outlook Classic, vá em **Arquivo → Adicionar Conta** (canto superior
   esquerdo).
2. Digite o seu endereço institucional (ex.: `voce@seu-dominio`) e clique em
   **Conectar**.
3. Coloque a senha da **Microsoft** e conclua o login / MFA, se pedir.
4. O Outlook detecta a caixa Exchange / Microsoft 365 sozinho e finaliza.
5. Espere a caixa terminar de carregar e o calendário dela aparecer na visão
   **Calendário**. Em **Arquivo → Informações sobre Contas** ela aparece como
   **Microsoft Exchange**.

## Passo 3 - (Opcional) Adicione a conta Google

Você NÃO precisa disso para o app. O GCalSync lê o seu Google Calendar pela API do
Google no login do próprio app, não pelo Outlook. Adicione a conta Google só se
você também quiser o e-mail do Google dentro do Outlook.

1. Vá em **Arquivo → Adicionar Conta** de novo.
2. Digite o endereço do Google e clique em **Conectar**. O Outlook adiciona como
   **IMAP/SMTP**; a página de login do Google abre, entre e aprove.
3. Conclua. (O IMAP traz só o e-mail, não o calendário do Google. Isso é o
   esperado.)

## Passo 4 - Confirme que está no Outlook Classic

Olhe o canto superior direito de novo. Se o botão **"Novo Outlook"** estiver
ligado, **desligue**. O app só funciona com o Classic.

## Passo 5 - Instale o GCalSync

1. Baixe o `GCalSync-Setup.exe` na
   [página de Releases](https://github.com/vinicq/gcal-to-outlook/releases).
2. Execute. O SmartScreen do Windows pode avisar sobre editor não reconhecido:
   clique em **Mais informações → Executar assim mesmo**. Instala por usuário, sem
   admin.

## Passo 6 - Primeira execução: o que o assistente pede

Na primeira vez que você abre o app, um assistente aparece numa janela de
terminal. É rápido. Siga o que ele pede:

1. **"Press Enter to begin"** - aperte Enter.
2. O assistente verifica o Outlook e seleciona o modo **Outlook COM** sozinho.
   Ele **não** pergunta a conta aqui.
3. **"Press Enter to open the browser"** - aperte Enter e, no navegador, **entre
   na sua conta Google e clique em aprovar** (acesso de leitura ao calendário).
   Nenhuma senha é guardada.
4. A primeira sincronização roda. Pode levar 1 a 2 minutos.
5. **"Schedule auto-start? [y/n]"** - digite **y** para o sync iniciar sozinho
   junto com o Windows. Dá para mudar isso depois na janela do monitor.

> Se você tem mais de uma conta Microsoft no Outlook (por exemplo uma pessoal e a
> institucional), o app pode escolher o calendário errado. Abra a janela do
> monitor, clique em **Reconfigure** ao lado de **Outlook / Teams** e digite o seu
> endereço institucional (`voce@seu-dominio`) para sincronizar no calendário
> certo.

## Passo 7 - Confira no Teams

Abra o Teams → Calendário. Seus eventos do Google aparecem em um ou dois minutos
(Teams e Outlook usam o mesmo calendário do Microsoft 365). Eles também aparecem
no Outlook.

## Passo 8 - Deixe rodando

Na janela do monitor, marque **Start automatically when Windows starts**. A
sincronização passa a rodar sozinha em segundo plano a cada login.

---

# Resolução de problemas

### Pop-up do Outlook: "Um programa está tentando acessar informações de endereço de e-mail"

Aparece quando o Windows reporta que o seu antivírus está desligado, vencido ou
sem reportar status. É uma configuração de segurança do Outlook, não do app. Para
parar:

1. Feche o Outlook. Reabra **como administrador** (clique direito no atalho do
   Outlook → **Executar como administrador**).
2. Vá em **Arquivo → Opções → Central de Confiabilidade → Configurações da Central
   de Confiabilidade → Acesso Programático** e escolha **Nunca me avisar sobre
   atividades suspeitas**.
3. Reinicie o Outlook normalmente. (A opção fica cinza se o Outlook não for aberto
   como administrador.)

Ou garanta que há um antivírus ativo e atualizado no Windows Security; aí o aviso
para sozinho. Assinar o app não remove esse aviso.

### "invalid_client: The provided client secret is invalid"

O client secret do OAuth do Google foi redefinido ou revogado. Gere um secret novo
para o cliente OAuth do tipo Desktop no Google Cloud Console, baixe o
`google_credentials.json` atualizado, apague o `google_token.json` e faça login de
novo.

### Os eventos não aparecem no Teams

- Confirme que você adicionou a conta **Microsoft 365** no Outlook (Passo 2), não
  só a Google. O Teams mostra só o calendário daquela caixa.
- Confirme que você está no Outlook **Classic**, não no novo Outlook.
- Espere alguns minutos; o Teams pode atrasar em relação ao Outlook.

### O assistente diz que não encontrou o Outlook

Provavelmente você está no "novo Outlook" ou o Outlook Classic não está instalado.
Faça o Passo 0 e o Passo 1.

---

# Perguntas frequentes

**Isso sincroniza eventos do Teams de volta para o Google?** Não. A sincronização
é unidirecional, só do Google para a Microsoft.

**Meus dados vão para algum lugar?** Não. O app roda localmente. Seu token do
Google fica na sua máquina. O app lê o Google e grava no seu calendário do
Microsoft 365.

**Preciso do Microsoft 365 desktop / licença paga para a Forma B?** Você precisa do
Outlook Classic com a sua conta Microsoft 365 conectada. Siga o link de instalação
da Microsoft no Passo 1.
