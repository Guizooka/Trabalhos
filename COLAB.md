# Preço site por EAN — rodando no Google Colab

Cole cada bloco numa célula.

**Célula 1 — instalação (toda sessão nova)**
```python
from google.colab import drive
drive.mount('/content/drive')
!pip install -q playwright
!python -m playwright install --with-deps chromium
```

**Célula 2 — pasta de trabalho no Drive** (o script, a régua e a pasta `preco_site_dados/` ficam aqui)
```python
%cd /content/drive/MyDrive/precos_site
!pwd   # tem que mostrar /content/drive/MyDrive/... — senão cache e histórico somem ao fim da sessão
```

**Célula 3 — rodar**
```python
!python preco_site_por_ean.py "Busca preços - Atualizado.xlsx"
```

- Sempre com `!python` (não `%run`): o Playwright síncrono não roda dentro do notebook.
- `--paralelo N` (padrão 6) = sites consultados ao mesmo tempo. Cada site continua no seu ritmo e com
  uma consulta por vez; o que muda é que os sites não esperam uns pelos outros. `--paralelo 1` = como antes.
- Durante a rodada, `preco_site_PARCIAL.xlsx` é regravado a cada 15 min; no fim, vira o Excel datado.
- Caiu a sessão? Rode a célula 3 de novo: continua de onde parou (cache no Drive).
- Teste rápido: `--limite 3`.
