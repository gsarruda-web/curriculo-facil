# Currículo Fácil - MVP v1.1

Aplicação web para geração guiada de currículos pelo celular.

## Fluxo
QR Code -> página com uma pergunta por vez -> revisão -> geração do currículo em Word.

## O que esta versão já faz
- interface mobile-first;
- uma pergunta por vez;
- formação, cursos e múltiplas experiências;
- perfil e objetivo automáticos quando o usuário não escreve;
- modelo Word fixo para preservar o padrão visual;
- seções vazias removidas automaticamente;
- aviso de privacidade;
- arquivo gerado temporariamente e apagado do servidor após o download;
- endpoint de saúde `/health`;
- configuração pronta para Render.

## Arquivos principais
- `modelo_curriculo_base.docx`: modelo visual fixo.
- `gerador_curriculo.py`: preenche o modelo.
- `app.py`: servidor FastAPI.
- `static/index.html`: questionário.
- `static/privacidade.html`: aviso de privacidade.
- `render.yaml`: configuração de deploy no Render.
- `.python-version`: versão do Python.
- `gerar_qr.py`: gera o QR Code depois que existir URL pública.

## Teste local
```bash
pip install -r requirements.txt
uvicorn app:app --reload
```
Abra `http://127.0.0.1:8000`.

## Publicação no Render
1. Crie um repositório no GitHub e envie esta pasta para ele.
2. No Render, crie um **Web Service** conectado ao repositório.
3. O arquivo `render.yaml` já contém os comandos essenciais.
4. Se configurar manualmente:
   - Build: `pip install -r requirements.txt`
   - Start: `uvicorn app:app --host 0.0.0.0 --port $PORT`
5. Após o deploy, o Render fornecerá uma URL pública `*.onrender.com`.

## Gerar o QR Code definitivo
Com a URL pública em mãos:
```bash
python gerar_qr.py https://SEU-ENDERECO.onrender.com
```
O arquivo `qr_curriculo.png` poderá ser usado em cartazes e cards.

## Privacidade
Esta versão não grava os dados enviados em banco. O currículo é criado em arquivo temporário e removido após o envio ao usuário. Para operação pública, revise o aviso de privacidade e identifique o responsável pela iniciativa antes do lançamento oficial.
