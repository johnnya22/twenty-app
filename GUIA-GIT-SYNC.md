# Twenty — Git como base de dados

## O que esta versão faz

- Continua a guardar primeiro no dispositivo, por isso funciona sem Internet.
- Cada alteração fica numa fila local.
- Quando há Internet, cada alteração cria o seu próprio commit no GitHub.
- PC e telemóvel são coordenados pelo Cloudflare Durable Object antes do commit.
- Alterações em itens ou campos diferentes são fundidas.
- Se os dois dispositivos alterarem exatamente o mesmo campo, a última alteração recebida ganha nesse campo, mas o conflito e o valor anterior ficam registados no JSON e no histórico Git.
- Se um dispositivo apagar um item que o outro alterou, o item alterado é preservado e o conflito fica registado.

## O que não sincroniza ainda

Os PDFs e imagens enviados pela app continuam apenas no IndexedDB de cada dispositivo. O Git sincroniza os dados académicos e os metadados dos materiais, mas não os ficheiros binários.

## Preparação

1. Cria um repositório **privado** no GitHub, por exemplo `twenty-data`.
2. Inicializa-o com um README para a branch `main` existir.
3. Cria um **fine-grained personal access token** limitado apenas a esse repositório.
4. Dá ao token a permissão **Repository permissions → Contents → Read and write**.
5. Abre `CONFIGURAR-GIT-SYNC.command` com `Control + clique → Abrir`.
6. Introduz o username, o repositório e o token.
7. No fim, abre a Twenty e vai a **Admin & dados → Git como base de dados → Configurar**.
8. Cola o URL do Worker e a chave mostrada pelo instalador.

## No telemóvel

O endereço `twenty.co` configurado no ficheiro `hosts` só funciona no Mac. Para usar a app fora do Mac, publica esta pasta num alojamento estático HTTPS, por exemplo Cloudflare Pages, e adiciona esse endereço à variável `ALLOWED_ORIGINS` do Worker.

Exemplo:

```text
https://twenty.co,https://twenty-study-os.pages.dev
```

Depois executa novamente:

```bash
cd cloudflare-worker
npx wrangler deploy
```

## Onde ficam os dados

O Worker cria ou atualiza este ficheiro no repositório:

```text
data/twenty-state.json
```

O token do GitHub fica guardado como segredo encriptado no Cloudflare Worker e nunca é incluído na PWA.

### Publicação automática para o telemóvel

Também podes executar `PUBLICAR-TWENTY-TELEMOVEL.command`. O script cria uma cópia pública limpa da PWA e faz o Direct Upload para Cloudflare Pages. Não envia o token do GitHub nem a chave de sincronização.
