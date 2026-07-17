# Twenty · Study OS

Aplicação PWA local-first para organizar o semestre e estudar com contexto. Foi desenhada para funcionar sem backend e sem conta: os dados e os PDFs enviados ficam no dispositivo.

## Abrir no computador

Uma PWA deve ser servida por HTTP; não abras apenas `index.html` com duplo clique.

- macOS/Linux: abre `INICIAR.command` ou executa `python3 start-server.py`.
- Windows: abre `INICIAR.bat` (Python 3 necessário).
- Alternativa: nesta pasta executa `npx serve .`.

A app abre em `http://127.0.0.1:8080`.

## Instalar no Android

Para instalar noutro dispositivo, publica esta pasta num alojamento estático com HTTPS (GitHub Pages, Netlify, Cloudflare Pages, etc.), abre o endereço no Chrome Android e escolhe **Adicionar ao ecrã principal / Instalar app**.

O servidor local `127.0.0.1` só é acessível no próprio computador. Para usar os mesmos dados em vários dispositivos seria necessário acrescentar sincronização com um backend; esta versão não finge ter cloud.

## Primeiro arranque

Quando não há informação, o onboarding configura:

1. perfil, curso e instituição;
2. semestre e ano letivo;
3. cadeiras, ECTS, tipos de aula e método de avaliação;
4. horário semanal;
5. testes, projetos ou exames já conhecidos.

O tutorial inicial pode ser saltado e reaberto em **Admin & dados**.

## Funcionalidades

- aula em direto detetada pelo horário, com nome real, materiais, quiz e perguntas anteriores;
- cada aula é ligada a uma ocorrência válida do horário: mesma cadeira, dia e tipo (T, TP, P, LAB ou OT);
- depois de preparares a próxima aula, o botão da Home muda automaticamente para **Abrir aula** e **Editar aula**;
- opção **Calendário** diretamente na navegação, com alternância entre **Horário** recorrente e calendário em vista diária, 3 dias, semanal ou mensal;
- calendário com aulas pelo nome real, testes, projetos, exames, eventos e tarefas;
- dashboard contextual que dá prioridade à aula em direto, tarefas atrasadas, aulas por rever, avaliações próximas, próxima aula preparada e eventos do dia;
- aula com data, tipo, sala, tópicos e apontamentos;
- PDFs/slides do ano atual e de anos anteriores;
- etiqueta de ano apenas nos materiais antigos;
- perguntas de testes anteriores ligadas a uma ou várias aulas;
- testes e exames anteriores organizados por ano letivo, com importação integral por ficheiro ou texto JSON;
- exemplo JSON e prompt rigoroso para converter um teste com IA sem inventar partes ilegíveis ou soluções em falta;
- importação de cadeiras e respetivos métodos de avaliação por JSON, também com exemplo e prompt;
- imagens no enunciado, solução e explicação de cada pergunta, por upload local ou caminho do projeto;
- imagens em eventos da faculdade;
- **BEFIRST™**: cada aula terminada fica por rever até concluíres um quiz associado;
- tarefa automática **Quiz da aula** depois do fim da aula, incluindo aulas em atraso;
- geração de quiz a partir das perguntas de testes anteriores da aula;
- quizzes normais/manuais que também podem misturar perguntas anteriores depois de escolheres a aula;
- perguntas antigas abertas em modo de autoavaliação: respondes, revelas a solução guardada e marcas se sabias, sem opções falsas inventadas;
- tarefas, TPC, projetos, revisões e eventos da faculdade;
- botão **Rever aula** para criar uma tarefa no dia seguinte, com acesso direto à aula;
- avaliações com aulas teóricas/práticas incluídas;
- cada nota ligada a uma avaliação concreta (teste, projeto, exame, apresentação) ou a uma aula concreta;
- avaliações ligadas explicitamente à componente correspondente do método de avaliação;
- método de avaliação com quantidade e peso total de testes, projetos, exame e componentes personalizadas, como quizzes ou mini-projetos;
- mínimos de aprovação por componente, defesas orais/práticas, limitação da nota sem defesa e registo da nota final após defesa;
- testes de consulta, lembrete para comprar folha de teste e substituição de avaliações concretas por exame;
- média por cadeira e média global ponderada por ECTS;
- simulador “se tiver X na próxima avaliação” sem guardar notas reais;
- planeamento diário com blocos de tempo, drag-and-drop no computador e alternativa tátil para Android;
- preenchimento automático do dia com sessões, pausas e almoço, além de cópia de rotinas;
- revisão semanal com tarefas atrasadas, aulas por rever, quizzes, dúvidas e prioridades;
- estimativa de horas por cadeira a partir dos ECTS, avaliações próximas, pesos e trabalho pendente;
- arquivo de semestres, preservando consulta;
- pesquisa global;
- menu **Cantina** com estado aberta/fechada no topo, almoço em destaque e jantar recolhido, além de kcal e alergénios;
- preço da refeição social e períodos de encerramento lidos da página oficial de Alimentação da SAS NOVA;
- atualização direta através da API REST da SAS NOVA, mantendo a última ementa e informação de serviço em cache para consulta offline;
- modo campus com contador explicitamente identificado como simulação;
- funcionamento offline depois da primeira abertura.

## BEFIRST™ depois da aula

Ao terminar uma aula, a Twenty cria uma ação de quiz. Abre os slides, responde ao quiz, identifica o que não sabias e leva essas dúvidas ao professor. Ao concluir qualquer quiz ligado à aula, a aula fica marcada como revista e a tarefa automática é concluída.

Se a aula já tiver perguntas de testes anteriores, podes gerar o quiz com um botão. Se não tiver, continuas a poder criar o quiz normal manualmente. Uma pergunta anterior sem opções usa a resposta que guardaste para autoavaliação; a app não fabrica respostas erradas.

## Editar dados sem abrir a app

Edita `data/academic-data.json`. A app relê o ficheiro ao abrir, ao voltares à janela ou através de **Admin & dados → Reler**. A união é feita por `id`, por isso usa identificadores estáveis.

O botão **Exportar** cria um JSON já preenchido. Podes substituir o ficheiro da pasta por esse export e continuar a editá-lo. A documentação do esquema está em `data/README.md`; `data/academic-data.example.json` contém dados exclusivamente demonstrativos.

Importante: um navegador não pode escrever silenciosamente num ficheiro do projeto. Por isso, alterações feitas dentro da app são guardadas no IndexedDB e exportadas por botão; alterações feitas fora são lidas do JSON.

## Importar testes anteriores e cadeiras

Em **Cadeira → Perguntas anteriores → Importar teste** podes criar o teste de origem e importar todas as perguntas de uma vez. O formulário aceita um ficheiro `.json` ou JSON colado e disponibiliza:

- um exemplo com todos os campos suportados;
- um prompt pronto a copiar para uma IA;
- validação integral antes de gravar qualquer pergunta;
- caminhos de imagem para enunciado, solução e explicação.

Depois da importação, cada pergunta pode ser editada para carregar imagens diretamente no dispositivo e para associar as aulas corretas. Em **Cadeiras → Importar JSON** existe o fluxo equivalente para criar cadeiras e métodos de avaliação.

O prompt incluído proíbe completar informação em falta. Texto ilegível deve permanecer marcado como `[ILEGÍVEL]`; soluções e explicações só devem ser incluídas quando existirem no documento original.

## Planeamento e revisão semanal

Em **Calendário → Dia de estudo** podes arrastar tarefas, aulas, quizzes e avaliações para uma hora. Em ecrãs táteis, usa o botão de adicionar de cada item. **Preencher dia** usa a configuração de duração, pausas e almoço sem apagar blocos existentes; **Copiar rotina** replica os blocos de outro dia e preserva o destino.

A página **Estudar → Revisão semanal** reúne o que ficou pendente e permite guardar prioridades, dúvidas e notas. A distribuição de horas é uma estimativa: parte do total semanal configurado e pondera ECTS, proximidade e peso das avaliações, aulas por rever e tarefas pendentes.

## PDFs

- Upload dentro da app: bytes guardados no IndexedDB do dispositivo.
- Ficheiro no projeto: coloca-o em `assets/slides/` e referencia o caminho no JSON.
- PowerPoint: pode ser guardado/aberto, mas a pré-visualização depende do navegador. Para leitura integrada, exporta para PDF.

## Cantina da FCT

A página **Cantina** consulta dois endpoints públicos do WordPress da SAS NOVA:

`https://sas.unl.pt/wp-json/wp/v2/pages/326?_fields=acf,link`

`https://sas.unl.pt/wp-json/wp/v2/pages/309?_fields=acf,link,modified`

A ementa, o preço e os períodos de encerramento estão em `acf.seccao`, não num ficheiro JSON separado com um nome óbvio. A app interpreta esse conteúdo, calcula o estado aberta/fechada no fuso horário de Lisboa e guarda a última resposta no dispositivo. Se a rede ou o site oficial falharem, usa primeiro a cópia local e, num primeiro arranque offline, `data/canteen-menu.json`.

No fundo da página existem ligações para a ementa, o preçário e as condições oficiais. A ementa, os valores nutricionais e os avisos de encerramento continuam a ser responsabilidade da SAS NOVA e podem sofrer alterações.

## Privacidade e limites

Não existem contas, servidor, utilizadores reais ou competição online. O contador de estudantes é uma simulação visual claramente assinalada. Limpar os dados do navegador pode apagar uploads locais; mantém backups JSON e cópias dos PDFs originais.
