"""
server_producao.py — Servidor do Gerador de Tribos
===================================================
Com proteção por senha simples.

Para rodar:
    python3 server_producao.py

Configurações:
    SENHA         — senha de acesso ao sistema
    PORT          — porta do servidor (padrão 8000)
"""

import sys, os, json, base64, tempfile, traceback, hashlib, secrets
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# ══════════════════════════════════════════
# CONFIGURAÇÕES — ALTERE AQUI
# ══════════════════════════════════════════
SENHA = "acampamento2026"   # ← TROQUE ESTA SENHA
PORT  = 8000
# ══════════════════════════════════════════

sys.path.insert(0, os.path.dirname(__file__))
try:
    from gerar_tribos import ler_inscricoes, identificar_casais, atribuir_tribos, rodar_testes, gerar, TRIBOS
    GERADOR_OK = True
except ImportError as e:
    GERADOR_OK = False
    IMPORT_ERR = str(e)

# Sessões ativas (token → True)
SESSOES = {}

def nova_sessao():
    token = secrets.token_hex(32)
    SESSOES[token] = True
    return token

def sessao_valida(token):
    return token and token in SESSOES

def senha_correta(senha):
    return senha == SENHA


class Handler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        print(f"  {self.address_string()} {fmt % args}")

    def do_GET(self):
        path = urlparse(self.path).path

        if path in ('/', '/index.html', '/gerador_tribos.html'):
            # Verifica cookie de sessão
            cookies = self._get_cookies()
            if sessao_valida(cookies.get('session')):
                self._serve_file('gerador_tribos.html', 'text/html; charset=utf-8')
            else:
                self._serve_login()
        else:
            self._404()

    def do_POST(self):
        if self.path == '/login':
            self._handle_login()
        elif self.path == '/api/gerar_tribos':
            cookies = self._get_cookies()
            if not sessao_valida(cookies.get('session')):
                self._json({'error': 'Não autorizado. Faça login.'}, 401)
                return
            self._handle_generate()
        elif self.path == '/logout':
            cookies = self._get_cookies()
            token = cookies.get('session')
            if token in SESSOES:
                del SESSOES[token]
            self.send_response(302)
            self.send_header('Location', '/')
            self.send_header('Set-Cookie', 'session=; Max-Age=0; Path=/')
            self.end_headers()
        else:
            self._404()

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def _handle_login(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length).decode('utf-8')
        try:
            data = json.loads(body)
            senha = data.get('senha', '')
        except Exception:
            senha = ''

        if senha_correta(senha):
            token = nova_sessao()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            # Cookie válido por 8 horas
            self.send_header('Set-Cookie', f'session={token}; Max-Age=28800; Path=/; HttpOnly')
            self._cors()
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True}).encode())
        else:
            self._json({'ok': False, 'error': 'Senha incorreta'}, 401)

    def _handle_generate(self):
        if not GERADOR_OK:
            self._json({'error': f'Erro ao importar gerar_tribos.py: {IMPORT_ERR}'}, 500)
            return

        try:
            ct = self.headers.get('Content-Type', '')
            length = int(self.headers.get('Content-Length', 0))
            raw = self.rfile.read(length)
            files, fields = parse_multipart(ct, raw)
        except Exception as e:
            self._json({'error': f'Erro ao ler dados: {e}'}, 400)
            return

        base_data = files.get('base')
        anjos_str = fields.get('anjos', '{}')
        if not base_data:
            self._json({'error': 'Arquivo não enviado.'}, 400)
            return

        try:
            anjos = json.loads(anjos_str)
        except Exception:
            anjos = {}

        base_fname = fields.get('base_filename', 'inscricoes.xlsx')
        _, base_ext = os.path.splitext(base_fname)
        base_ext = base_ext.lower() if base_ext.lower() in ['.xlsx','.xls','.xlsm','.csv'] else '.xlsx'

        with tempfile.NamedTemporaryFile(suffix=base_ext, delete=False) as f:
            f.write(base_data); base_path = f.name

        out_path = tempfile.mktemp(suffix='.xlsx')

        try:
            import pandas as pd
            print(f"\n📂 Lendo arquivo...")
            df = ler_inscricoes(base_path)

            print("\n🔍 Identificando casais...")
            partner = identificar_casais(df)
            print(f"   {len(partner)//2} casais")

            print("\n⚖️  Balanceando tribos...")
            df = atribuir_tribos(df, partner)

            print("\n✅ Rodando testes...")
            erros, avisos = rodar_testes(df, partner)

            print(f"\n📝 Gerando planilha...")
            has_anjos = any(v.get('a1','(preencher)') != '(preencher)' for v in anjos.values())
            gerar(df, partner, anjos=anjos if has_anjos else None, output_path=out_path)

            with open(out_path, 'rb') as f:
                xlsx_b64 = base64.b64encode(f.read()).decode()

            tribos_result = []
            for t in TRIBOS:
                g = df[df['TRIBO'] == t]
                m = int((g['Sexo'] == 'M').sum())
                fc = int((g['Sexo'] == 'F').sum())
                tribos_result.append({
                    'nome': t, 'mulheres': fc, 'homens': m,
                    'peso_medio': round(float(g['Peso_N'].mean()), 1),
                    'idade_media': round(float(g['Idade_N'].mean()), 1),
                    'ok': (m == fc),
                })

            casais_result = []
            seen = set()
            for fc_a, fc_b in partner.items():
                if fc_a > fc_b or (fc_a, fc_b) in seen: continue
                seen.add((fc_a, fc_b))
                ra = df[df['FC'] == fc_a]
                rb = df[df['FC'] == fc_b]
                if ra.empty or rb.empty: continue
                casais_result.append({
                    'fc1': int(fc_a), 'n1': ra.iloc[0]['Nome'], 't1': ra.iloc[0]['TRIBO'],
                    'fc2': int(fc_b), 'n2': rb.iloc[0]['Nome'], 't2': rb.iloc[0]['TRIBO'],
                })

            self._json({
                'n_campistas': len(df),
                'n_casais': len(partner) // 2,
                'all_ok': len(erros) == 0,
                'erros': erros, 'avisos': avisos,
                'tribos': tribos_result,
                'casais': casais_result,
                'xlsx_b64': xlsx_b64,
            })
            print("✅ Concluído!\n")

        except Exception as e:
            traceback.print_exc()
            self._json({'error': str(e)}, 500)
        finally:
            for p in [base_path, out_path]:
                if p and os.path.exists(p):
                    try: os.unlink(p)
                    except: pass

    def _serve_login(self):
        html = '''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gerador de Tribos — Login</title>
<link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;900&family=Crimson+Pro:wght@300;400&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Crimson Pro',serif;background:#FAF6EE;min-height:100vh;display:flex;align-items:center;justify-content:center}
body::before{content:'';position:fixed;inset:0;background:radial-gradient(ellipse 80% 60% at 20% 20%,rgba(200,150,42,.1),transparent 60%),radial-gradient(ellipse 60% 80% at 80% 80%,rgba(255,69,0,.07),transparent 60%);pointer-events:none}
.card{position:relative;z-index:1;background:white;border-radius:20px;padding:48px 40px;width:100%;max-width:400px;box-shadow:0 8px 48px rgba(0,0,0,.08);text-align:center}
.flames{display:flex;align-items:flex-end;justify-content:center;gap:5px;margin-bottom:20px}
.flame{border-radius:50% 50% 40% 40%;animation:flicker 2s ease-in-out infinite}
.flame-lg{width:20px;height:28px;background:linear-gradient(175deg,#F5C842,#FF4500)}
.flame-sm{width:12px;height:18px;background:linear-gradient(175deg,#FFD700,#FF4500);opacity:.7;animation-delay:.4s}
@keyframes flicker{0%,100%{transform:scaleY(1) rotate(-1deg)}50%{transform:scaleY(1.1) rotate(1deg)}}
h1{font-family:'Cinzel',serif;font-size:22px;font-weight:900;color:#1A1208;margin-bottom:6px}
h1 span{color:#C8962A}
.sub{font-size:13px;color:#6B5E4A;font-style:italic;margin-bottom:32px}
.field{margin-bottom:20px;text-align:left}
label{display:block;font-family:'Cinzel',serif;font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#C8962A;margin-bottom:8px}
input[type=password]{width:100%;padding:12px 16px;border:1.5px solid rgba(200,150,42,.3);border-radius:10px;font-family:'Crimson Pro',serif;font-size:15px;color:#1A1208;background:#FAF6EE;outline:none;transition:border-color .2s}
input[type=password]:focus{border-color:#C8962A;background:white}
.btn{width:100%;padding:14px;font-family:'Cinzel',serif;font-size:14px;font-weight:600;letter-spacing:.08em;color:white;background:linear-gradient(135deg,#C8962A,#FF4500);border:none;border-radius:50px;cursor:pointer;box-shadow:0 4px 16px rgba(200,150,42,.3);transition:all .2s}
.btn:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(200,150,42,.4)}
.btn:active{transform:none}
.err{display:none;margin-top:16px;padding:10px 14px;background:rgba(192,57,43,.08);border:1px solid rgba(192,57,43,.2);border-radius:8px;font-size:13px;color:#C0392B}
.err.vis{display:block}
</style>
</head>
<body>
<div class="card">
  <div class="flames"><div class="flame flame-sm"></div><div class="flame flame-lg"></div><div class="flame flame-sm" style="animation-delay:.8s"></div></div>
  <h1>Gerador de <span>Tribos</span></h1>
  <p class="sub">Acampamento Espírito Empreendedor</p>
  <div class="field">
    <label>Senha de Acesso</label>
    <input type="password" id="senha" placeholder="Digite a senha..." autofocus>
  </div>
  <button class="btn" id="btn" onclick="entrar()">Entrar</button>
  <div class="err" id="err">Senha incorreta. Tente novamente.</div>
</div>
<script>
document.getElementById('senha').addEventListener('keydown', e => { if(e.key==='Enter') entrar(); });
async function entrar(){
  const senha = document.getElementById('senha').value;
  const btn   = document.getElementById('btn');
  const err   = document.getElementById('err');
  err.classList.remove('vis');
  btn.textContent = 'Verificando...';
  btn.disabled = true;
  try{
    const r = await fetch('/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({senha})});
    const d = await r.json();
    if(d.ok){ window.location.href = '/'; }
    else { err.classList.add('vis'); btn.textContent='Entrar'; btn.disabled=false; document.getElementById('senha').focus(); }
  }catch(e){ err.textContent='Erro de conexão.'; err.classList.add('vis'); btn.textContent='Entrar'; btn.disabled=false; }
}
</script>
</body>
</html>'''
        data = html.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', len(data))
        self.end_headers()
        self.wfile.write(data)

    def _serve_file(self, filename, ctype):
        path = os.path.join(os.path.dirname(__file__), filename)
        try:
            with open(path, 'rb') as f: data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', len(data))
            self._cors()
            self.end_headers()
            self.wfile.write(data)
        except FileNotFoundError:
            self._404()

    def _json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _404(self):
        self.send_response(404)
        self.send_header('Content-Type', 'text/plain')
        self.end_headers()
        self.wfile.write(b'Not found')

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _get_cookies(self):
        cookies = {}
        raw = self.headers.get('Cookie', '')
        for part in raw.split(';'):
            if '=' in part:
                k, v = part.strip().split('=', 1)
                cookies[k.strip()] = v.strip()
        return cookies


def parse_multipart(content_type, body):
    import re
    boundary = None
    for part in content_type.split(';'):
        part = part.strip()
        if part.startswith('boundary='):
            boundary = part[9:].strip('"')
            break
    if not boundary: return {}, {}
    files, fields = {}, {}
    delimiter = ('--' + boundary).encode()
    parts = body.split(delimiter)
    for part in parts[1:]:
        if part.strip() in (b'', b'--', b'--\r\n'): continue
        if part.startswith(b'--'): continue
        if b'\r\n\r\n' in part: header_raw, content = part.split(b'\r\n\r\n', 1)
        elif b'\n\n' in part: header_raw, content = part.split(b'\n\n', 1)
        else: continue
        content = content.rstrip(b'\r\n')
        headers_str = header_raw.decode('utf-8', errors='replace')
        cd = re.search(r'Content-Disposition:[^\r\n]*name="([^"]+)"', headers_str, re.IGNORECASE)
        if not cd: continue
        name = cd.group(1)
        fn = re.search(r'filename="([^"]*)"', headers_str, re.IGNORECASE)
        if fn: files[name] = content
        else: fields[name] = content.decode('utf-8', errors='replace')
    return files, fields


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    print("=" * 50)
    print("  Gerador de Tribos — Servidor")
    print("=" * 50)
    print(f"\n🔒  Senha configurada: {SENHA}")
    print(f"🌐  Rodando em: http://localhost:{PORT}")
    print("   (Ctrl+C para parar)\n")
    if not GERADOR_OK:
        print(f"❌  Erro: {IMPORT_ERR}\n")
    server = HTTPServer(('', PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServidor encerrado.')
