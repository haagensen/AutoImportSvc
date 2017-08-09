Certo dia precisei criar um "serviço" do Windows para um software da empresa.

Nada muito difícil: ele ficaria rodando, mesmo sem um usuário "logado" (afinal, pra isso servem os Serviços, certo?) e, na hora "tal", dispararia um outro executável.

Sopinha, certo? Escolhi o Visual C# 2013 para a tarefa, achando que, bem, com o Visual Studio 2013 eu conseguiria construir esse tipo de projeto de forma _muito_ simples. Possivelmente ele já teria um "modelo" com todo o código base pronto, bonitinho e funcionando, clicaria no botão "play" e já estava lá um "serviço exemplo", rodando lindo, levre e solto.

Bem... não.

O Visual Studio tem, sim, um modelo (ou "template", se preferir) para criar serviços do Windows. Ou, ao menos, "algumas" versões o têm -- por algum motivo, a edição "Express" não tem. Acontece que o modelo é algo bem, bem básico mesmo. Para começar a funcionar, você precisa dar uma olhada numa página como [Walkthrough: Creating a Windows Service Application in the Component Designer](https://docs.microsoft.com/en-us/dotnet/framework/windows-services/walkthrough-creating-a-windows-service-application-in-the-component-designer).

Estava eu seguindo o "walkthrough", mas quando cheguei no ponto 3 da seção "Criando um Serviço" -- que diz explicitamente, "no menu Edit, escolha Encontrar e Substituir, Encontrar nos Arquivos, mude todas as ocorrências de "Service1" para "MeuNovoServiço" (...) -- bem, sinceramente, quando vi isso eu fiquei meio assustado.

Quer dizer... isso de "encontrar e substituir" é o tipo de coisa que qualquer IDE moderna faz de forma automática (ou semi-automática) pra você. Ou "deveria" fazer. Encontre a variável "tal", mande renomear para "xpto", e ela sai fazendo isso sozinha, sem que você se preocupe com mais nada. Ver que o "walkthrough" manda você fazer um localizar/substituir explícito me fez sentir usando o Visual Studio 6, de 1997 (ou 1998?).

Não há uma forma simples de "debugar" um serviço -- é preciso gerar um executável, instalar no SCM (Service Control Management) do Windows usando uma conta de administrador local, emporcalhar o log de eventos com um monte de mensagens de debug (aqueles "print"s do tipo "passei aqui..."), etc. O Visual Studio teoricamente faria isso pra você, através de um programa que deve ser executado em linha de comando (!), mas descobri que ele simplesmente não funciona.

Mas tem mais. Aparentemente, não existe uma forma "limpa" de terminar o programa. Nem o "serviço modelo" nem o "walkthrough" falam nada sobre isso. Qualquer coisa como um "Stop()", "Environment.Exit", etc. vai deixar uma marca horrorosa no log de eventos do Windows, algo do tipo "seu serviço foi finalizado com um erro", ou "finalizado inesperadamente", coisas assim. Uma turma no Stack Overflow sugeriu usar uma API Win32 para fazer isso, mas sem muito sucesso (perdi o link sobre o assunto, depois eu coloco aqui).

Enfim. No fim das contas, se eu tivesse usado, sei lá, Visual Basic 6, acho que teria tido o mesmo efeito de usar C# 2013.

O código aqui possui diversos trechos que pesquei aqui e ali Internet afora, para se criar um serviço minimamente decente. Em "Program.cs" há diversos métodos que o "modelo" do Visual Studio não cria, ou não se interessa em criar, como "IsInstalled()", "IsRunning()", etc, que podem ser bastante úteis em um programa real.

De brinde, um leitor de arquivos INI, também! :-)
