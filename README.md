# Controle de Produtividade Banho — Pintura Eletrostatica

Monitoramento em tempo real do tratamento de pecas (pre-banho e banho):
grade de 19 cestos, leitura da OP por coletora, pausa de tempo, multiplas OPs
por cesto, 1 ou 2 operadores, dashboards, painel publico e Excel. Roda 24h no
Railway com PostgreSQL. Otimizado para tablet (Samsung Galaxy Tab A9).

==================================================================
IMPORTANTE: como subir sem erro
==================================================================
1. Os arquivos tem que ficar na RAIZ do repositorio. O app.py precisa
   aparecer DIRETO na pagina do repo, nao dentro de uma pasta.
2. NAO adicione nixpacks.toml. O Railway/Nixpacks instala as dependencias
   sozinho pelo requirements.txt. (Um nixpacks.toml com 'pip install' quebra
   o build com 'pip: command not found'.)
3. O comando de start usa 'python -m gunicorn' (no Procfile e railway.json) —
   isso evita o erro 'gunicorn: command not found'.

Se voce ja tinha subido uma versao antiga: APAGUE todos os arquivos do repo
(inclusive nixpacks.toml, start.sh, runtime.txt se existirem) e suba estes.

==================================================================
Fluxo de uso
==================================================================
1. Preparacao: grade de cestos 1-19 (livres coloridos, em uso com cadeado).
   Toca num livre -> escolhe 1 ou 2 operadores -> Iniciar (cronometro comeca).
2. Pode Pausar (cafe, ginastica) — tempo parado nao conta — e Retomar.
3. Ao terminar, toca em Parar tempo (cronometro congela). DEPOIS preenche:
   botao Adicionar OP (pode por VARIAS OPs no mesmo cesto; cada uma puxa
   codigo/descricao/qtd da lista mestra, editaveis), processo, tipo, e conclui.
   Vai para a fila do banho.
4. Banho: operador inicia e finaliza. O cesto volta a ficar livre.
5. Qualquer card pode ser editado depois (tocar no cesto ocupado).

==================================================================
Usuarios padrao (criados na 1a execucao — troque as senhas)
==================================================================
  admin / admin123   -> admin (dashboard, usuarios, lista mestra, Excel)
  banho / banho123   -> operador de banho
  op1..op6 / op1234  -> operadores de preparacao

==================================================================
Deploy no Railway
==================================================================
1. Suba os arquivos na RAIZ do repo no GitHub (app.py direto, nao em pasta).
2. railway.app -> New Project -> Deploy from GitHub repo -> escolha o repo.
3. New -> Database -> Add PostgreSQL (cria DATABASE_URL automaticamente).
4. No servico web, em Variables, adicione SECRET_KEY = uma frase longa.
5. Settings -> Networking -> Generate Domain para a URL publica.

Painel da gerencia (so leitura, sem senha): SUA_URL/painel

==================================================================
Lista mestra (Excel do SAP)
==================================================================
Tela Lista Mestra (admin) -> importar .xlsx do SAP. Colunas usadas:
Ordem (OP), Material (codigo), Texto breve material, Quantidade total.
Ao bipar a OP na preparacao, esses dados preenchem o item automaticamente.
Use o campo "Testar uma OP" para conferir se foi importada.

==================================================================
Relatorios / backup
==================================================================
Dashboard -> Excel pre-banho e Excel banho (uma linha por OP, com data/hora
no nome do arquivo). Dados no PostgreSQL nao se perdem em reinicio/deploy.

==================================================================
Rodar local
==================================================================
  pip install -r requirements.txt
  python app.py    (http://localhost:5000)
Sem DATABASE_URL usa SQLite local (dados_local.db) so para teste.
