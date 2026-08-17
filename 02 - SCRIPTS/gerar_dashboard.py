"""Gera um dashboard gerencial HTML moderno com histórico de custos e variações percentuais."""

from __future__ import annotations

import argparse
import html
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


PASTA_SCRIPT = Path(__file__).resolve().parent
PASTA_PROJETO = PASTA_SCRIPT.parent
ARQUIVO_PADRAO = PASTA_PROJETO / "01 - DADOS" / "TELEFONIA.xlsx"
SAIDA_PADRAO = PASTA_PROJETO / "DASH.html"
HISTORICO_PATH = PASTA_PROJETO / "04 - SAIDAS" / "historico_dashboard.json"
STATUS_EXIBIDOS = {"ATIVA", "ESTOQUE", "DESLIGADO", "VERIFICAR"}


def normalizar(valor: object) -> str:
    texto = "" if valor is None else str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(letra for letra in texto if not unicodedata.combining(letra))
    return re.sub(r"\s+", " ", texto)


def texto(valor: object) -> str:
    return "" if valor is None else str(valor).strip()


def localizar_colunas(ws) -> dict[str, int]:
    indices = {normalizar(celula.value): celula.column for celula in ws[1] if celula.value}
    obrigatorias = {"EMPRESA", "LINHA", "CHAPA/CPF", "NOME", "CPF", "COD CDC", "CDC", "STATUS"}
    ausentes = obrigatorias - indices.keys()
    if ausentes:
        raise ValueError("Colunas ausentes na aba Planos: " + ", ".join(sorted(ausentes)))
    return indices


def carregar_dados(arquivo: Path) -> list[dict[str, object]]:
    wb = load_workbook(arquivo, read_only=True, data_only=True)
    try:
        if "Planos" not in wb.sheetnames:
            raise ValueError("A aba 'Planos' nao foi encontrada.")
        ws = wb["Planos"]
        colunas = localizar_colunas(ws)
        col_valor = colunas.get("VALOR") or {normalizar(c.value): c.column for c in ws[1] if c.value}.get("VALOR")
        
        registros = []
        for valores in ws.iter_rows(min_row=2, values_only=True):
            if not any(valor not in (None, "") for valor in valores):
                continue
            status = normalizar(valores[colunas["STATUS"] - 1])
            if status not in STATUS_EXIBIDOS:
                continue
            empresa = normalizar(valores[colunas["EMPRESA"] - 1])
            
            val_num = 0.0
            if col_valor:
                val_raw = valores[col_valor - 1]
                if isinstance(val_raw, (int, float)):
                    val_num = float(val_raw)
                elif val_raw:
                    try:
                        limpo = str(val_raw).replace("R$", "").replace(".", "").replace(",", ".").strip()
                        val_num = float(limpo)
                    except ValueError:
                        val_num = 0.0

            registros.append({
                "operadora": empresa if empresa in {"VIVO", "TIM"} else "OUTRAS",
                "linha": texto(valores[colunas["LINHA"] - 1]),
                "chapaCpf": texto(valores[colunas["CHAPA/CPF"] - 1]),
                "nome": texto(valores[colunas["NOME"] - 1]),
                "cpf": texto(valores[colunas["CPF"] - 1]),
                "codCdc": texto(valores[colunas["COD CDC"] - 1]) or "SEM CODIGO",
                "cdc": texto(valores[colunas["CDC"] - 1]) or "SEM CENTRO DE CUSTO",
                "status": status,
                "valor": val_num,
            })
        return registros
    finally:
        wb.close()


def processar_historico(dados: list[dict[str, object]], gerado_em: datetime) -> dict[str, object]:
    """Salva a medição atual no histórico e retorna a comparação com a anterior."""
    HISTORICO_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    totais_atuais = {
        "data": gerado_em.strftime("%d/%m/%Y às %H:%M"),
        "total_linhas": len(dados),
        "custo_total": sum(float(d["valor"]) for d in dados),
        "ATIVA": {"qtd": 0, "val": 0.0},
        "ESTOQUE": {"qtd": 0, "val": 0.0},
        "DESLIGADO": {"qtd": 0, "val": 0.0},
        "VERIFICAR": {"qtd": 0, "val": 0.0},
    }
    for d in dados:
        st = str(d["status"])
        if st in totais_atuais:
            totais_atuais[st]["qtd"] += 1
            totais_atuais[st]["val"] += float(d["valor"])

    historico = []
    if HISTORICO_PATH.exists():
        try:
            historico = json.loads(HISTORICO_PATH.read_text(encoding="utf-8"))
            if not isinstance(historico, list):
                historico = []
        except Exception:
            historico = []

    anterior = historico[-1] if historico else None
    
    def calc_delta(atual_val: float, ant_val: float) -> dict[str, object]:
        diff = atual_val - ant_val
        pct = (diff / ant_val * 100) if ant_val > 0 else 0.0
        return {"diff": diff, "pct": round(pct, 1), "tem_anterior": ant_val > 0}

    comparativo = {
        "tem_anterior": anterior is not None,
        "data_anterior": anterior.get("data") if anterior else None,
        "custo_total": calc_delta(totais_atuais["custo_total"], anterior.get("custo_total", 0.0) if anterior else 0.0),
        "ATIVA": calc_delta(totais_atuais["ATIVA"]["val"], anterior.get("ATIVA", {}).get("val", 0.0) if anterior else 0.0),
        "ESTOQUE": calc_delta(totais_atuais["ESTOQUE"]["val"], anterior.get("ESTOQUE", {}).get("val", 0.0) if anterior else 0.0),
        "DESLIGADO": calc_delta(totais_atuais["DESLIGADO"]["val"], anterior.get("DESLIGADO", {}).get("val", 0.0) if anterior else 0.0),
        "VERIFICAR": calc_delta(totais_atuais["VERIFICAR"]["val"], anterior.get("VERIFICAR", {}).get("val", 0.0) if anterior else 0.0),
    }

    # Evita gravar snapshots duplicados se rodar no mesmo minuto
    if not historico or historico[-1].get("data") != totais_atuais["data"]:
        historico.append(totais_atuais)
        if len(historico) > 30:  # Guarda os últimos 30 snapshots
            historico = historico[-30:]
        HISTORICO_PATH.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")

    return comparativo


HTML_TEMPLATE = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dashboard Gerencial de Telefonia</title>
  <style>
    :root {
      --primary: #0b5269;
      --primary-light: #e8f3f6;
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --text: #1e293b;
      --text-muted: #64748b;
      --border: #e2e8f0;
      --ativo: #059669;
      --ativo-bg: #ecfdf5;
      --estoque: #3b82f6;
      --estoque-bg: #eff6ff;
      --desligado: #dc2626;
      --desligado-bg: #fef2f2;
      --verificar: #d97706;
      --verificar-bg: #fffbeb;
      --badge-up: #ef4444;
      --badge-up-bg: #fef2f2;
      --badge-down: #10b981;
      --badge-down-bg: #ecfdf5;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: var(--bg); color: var(--text); padding: 24px 16px; line-height: 1.5; }
    .container { max-width: 1320px; margin: 0 auto; }
    
    header { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; margin-bottom: 24px; padding-bottom: 16px; border-bottom: 1px solid var(--border); }
    h1 { font-size: 24px; font-weight: 700; color: var(--primary); }
    .timestamp { font-size: 13px; color: var(--text-muted); }
    
    .filtros { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 24px; background: var(--card-bg); padding: 16px; border-radius: 8px; border: 1px solid var(--border); }
    .campo { display: flex; flex-direction: column; gap: 6px; flex: 1; min-width: 220px; }
    .campo label { font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); }
    select, input[type="text"] { height: 40px; padding: 0 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; background: #fff; color: var(--text); outline: none; }
    select:focus, input[type="text"]:focus { border-color: var(--primary); }
    
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px; margin-bottom: 24px; }
    .card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; padding: 18px; cursor: pointer; transition: transform 0.15s, box-shadow 0.15s; text-align: left; }
    .card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    .card.ativo { border-top: 4px solid var(--ativo); }
    .card.estoque { border-top: 4px solid var(--estoque); }
    .card.desligado { border-top: 4px solid var(--desligado); }
    .card.verificar { border-top: 4px solid var(--verificar); }
    
    .card-top { display: flex; justify-content: space-between; align-items: center; }
    .card-label { font-size: 13px; font-weight: 600; color: var(--text-muted); }
    
    /* Apple-style Trend Badges */
    .trend-badge { display: inline-flex; align-items: center; gap: 3px; padding: 2px 7px; border-radius: 12px; font-size: 11px; font-weight: 700; }
    .trend-badge.down { background: var(--badge-down-bg); color: var(--badge-down); }
    .trend-badge.up { background: var(--badge-up-bg); color: var(--badge-up); }
    .trend-badge.neutral { background: #f1f5f9; color: var(--text-muted); }
    
    .card-val-qtd { display: flex; align-items: baseline; justify-content: space-between; margin-top: 10px; }
    .card-qtd { font-size: 26px; font-weight: 700; color: var(--text); }
    .card-valor { font-size: 15px; font-weight: 600; color: var(--primary); }
    .card-subtext { font-size: 11px; color: var(--text-muted); margin-top: 4px; }
    
    .nav-tabs { display: flex; gap: 8px; border-bottom: 2px solid var(--border); margin-bottom: 20px; }
    .tab-btn { padding: 10px 20px; border: none; background: transparent; font-size: 14px; font-weight: 600; color: var(--text-muted); cursor: pointer; border-bottom: 3px solid transparent; margin-bottom: -2px; }
    .tab-btn.active { color: var(--primary); border-bottom-color: var(--primary); }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    
    .grid-graficos { display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 20px; }
    .painel-grafico { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; padding: 20px; }
    .painel-grafico h3 { font-size: 16px; color: var(--primary); margin-bottom: 16px; }
    .barra-item { margin-bottom: 14px; }
    .barra-info { display: flex; justify-content: space-between; font-size: 13px; margin-bottom: 4px; font-weight: 500; }
    .barra-trilha { height: 12px; background: #f1f5f9; border-radius: 6px; overflow: hidden; }
    .barra-preenchimento { height: 100%; border-radius: 6px; transition: width 0.4s ease; }
    
    .tabela-container { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; overflow: hidden; }
    .tabela-scroll { max-height: 480px; overflow-y: auto; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 12px 14px; text-align: left; border-bottom: 1px solid var(--border); }
    th { position: sticky; top: 0; background: var(--primary); color: #fff; font-weight: 600; font-size: 12px; text-transform: uppercase; }
    td.num, th.num { text-align: right; }
    tr:hover td { background: var(--primary-light); cursor: pointer; }
    .tag { padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; display: inline-block; }
    .tag.ativa { background: var(--ativo-bg); color: var(--ativo); }
    .tag.estoque { background: var(--estoque-bg); color: var(--estoque); }
    .tag.desligado { background: var(--desligado-bg); color: var(--desligado); }
    .tag.verificar { background: var(--verificar-bg); color: var(--verificar); }
    
    .busca-detalhe { display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; background: #f8fafc; border-bottom: 1px solid var(--border); }
    .btn-limpar { padding: 6px 12px; background: #fff; border: 1px solid var(--border); border-radius: 4px; font-size: 12px; cursor: pointer; }
  </style>
</head>
<body>
<div class="container">
  <header>
    <div>
      <h1>Dashboard Gerencial de Telefonia</h1>
      <p class="timestamp">Base de Dados: aba Planos · Atualizado em __DATA_GERADO__</p>
    </div>
  </header>

  <section class="filtros">
    <div class="campo">
      <label>Operadora</label>
      <select id="filtroOperadora">
        <option value="TODAS">Todas as Operadoras</option>
        <option value="VIVO">VIVO</option>
        <option value="TIM">TIM</option>
      </select>
    </div>
    <div class="campo">
      <label>Centro de Custo</label>
      <select id="filtroCdc">
        <option value="TODOS">Todos os Centros de Custo</option>
      </select>
    </div>
  </section>

  <!-- Cards Indicadores com Flag de Variação -->
  <section class="cards">
    <div class="card ativo" onclick="selecionarFiltroCard('ATIVA')">
      <div class="card-top">
        <span class="card-label">Linhas Ativas</span>
        <span id="badgeAtiva" class="trend-badge neutral">0.0%</span>
      </div>
      <div class="card-val-qtd">
        <span class="card-qtd" id="qtdAtiva">0</span>
        <span class="card-valor" id="valAtiva">R$ 0,00</span>
      </div>
      <div class="card-subtext" id="subAtiva">Base atualizada</div>
    </div>
    
    <div class="card estoque" onclick="selecionarFiltroCard('ESTOQUE')">
      <div class="card-top">
        <span class="card-label">Em Estoque</span>
        <span id="badgeEstoque" class="trend-badge neutral">0.0%</span>
      </div>
      <div class="card-val-qtd">
        <span class="card-qtd" id="qtdEstoque">0</span>
        <span class="card-valor" id="valEstoque">R$ 0,00</span>
      </div>
      <div class="card-subtext" id="subEstoque">Reserva técnica</div>
    </div>
    
    <div class="card desligado" onclick="selecionarFiltroCard('DESLIGADO')">
      <div class="card-top">
        <span class="card-label">Colaboradores Desligados</span>
        <span id="badgeDesligado" class="trend-badge neutral">0.0%</span>
      </div>
      <div class="card-val-qtd">
        <span class="card-qtd" id="qtdDesligado">0</span>
        <span class="card-valor" id="valDesligado" style="color:var(--desligado);">R$ 0,00</span>
      </div>
      <div class="card-subtext" id="subDesligado">Economia potencial</div>
    </div>
    
    <div class="card verificar" onclick="selecionarFiltroCard('VERIFICAR')">
      <div class="card-top">
        <span class="card-label">Linhas a Verificar</span>
        <span id="badgeVerificar" class="trend-badge neutral">0.0%</span>
      </div>
      <div class="card-val-qtd">
        <span class="card-qtd" id="qtdVerificar">0</span>
        <span class="card-valor" id="valVerificar" style="color:var(--verificar);">R$ 0,00</span>
      </div>
      <div class="card-subtext" id="subVerificar">Aguardando validação</div>
    </div>
  </section>

  <div class="nav-tabs">
    <button class="tab-btn active" onclick="trocarAba(event, 'abaGraficos')">📊 Visão Gráfica & Top CDCs</button>
    <button class="tab-btn" onclick="trocarAba(event, 'abaResumoCdc')">🏢 Resumo por Centro de Custo</button>
    <button class="tab-btn" onclick="trocarAba(event, 'abaDetalhes')">📋 Detalhamento das Linhas</button>
  </div>

  <div id="abaGraficos" class="tab-content active">
    <div class="grid-graficos">
      <div class="painel-grafico">
        <h3>Top 8 CDCs por Custo Mensal (R$)</h3>
        <div id="graficoCusto"></div>
      </div>
      <div class="painel-grafico">
        <h3>Top 8 CDCs com Pendências (Verificar + Desligados)</h3>
        <div id="graficoPendencias"></div>
      </div>
    </div>
  </div>

  <div id="abaResumoCdc" class="tab-content">
    <div class="tabela-container">
      <div class="tabela-scroll">
        <table>
          <thead>
            <tr>
              <th>Cod CDC</th>
              <th>Centro de Custo</th>
              <th class="num">Ativas</th>
              <th class="num">Estoque</th>
              <th class="num">Desligado</th>
              <th class="num">Verificar</th>
              <th class="num">Custo Total</th>
            </tr>
          </thead>
          <tbody id="tabelaResumoCdc"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div id="abaDetalhes" class="tab-content">
    <div class="tabela-container">
      <div class="busca-detalhe">
        <input type="text" id="buscaTexto" placeholder="🔍 Filtrar por Nome, Linha, Chapa ou CDC..." oninput="atualizarDetalhe()" style="width: 320px; height: 34px;">
        <button class="btn-limpar" onclick="limparFiltrosDetalhe()">Limpar Filtros</button>
      </div>
      <div class="tabela-scroll">
        <table>
          <thead>
            <tr>
              <th>Operadora</th>
              <th>Linha</th>
              <th>Chapa/CPF</th>
              <th>Nome</th>
              <th>Cod CDC</th>
              <th>Centro de Custo</th>
              <th class="num">Valor</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody id="tabelaDetalhes"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<script>
const dados = __DADOS_JSON__;
const comparativo = __COMPARATIVO_JSON__;
let statusFiltro = null;

const fmtMoeda = v => Number(v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const esc = v => String(v || '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));

function renderizarBadge(idBadge, idSub, deltaInfo, isCostReductionGood = true) {
  const badgeEl = document.getElementById(idBadge);
  const subEl = document.getElementById(idSub);
  if (!badgeEl) return;

  if (!comparativo.tem_anterior || !deltaInfo || !deltaInfo.tem_anterior) {
    badgeEl.style.display = 'none';
    return;
  }

  const pct = deltaInfo.pct;
  const isZero = pct === 0;
  const isDrop = pct < 0;

  if (isZero) {
    badgeEl.className = 'trend-badge neutral';
    badgeEl.textContent = '0.0%';
  } else if (isDrop) {
    // Queda de custo é bom (Verde)
    badgeEl.className = isCostReductionGood ? 'trend-badge down' : 'trend-badge up';
    badgeEl.textContent = '▼ ' + pct + '%';
  } else {
    // Aumento de custo é alerta (Vermelho)
    badgeEl.className = isCostReductionGood ? 'trend-badge up' : 'trend-badge down';
    badgeEl.textContent = '▲ +' + pct + '%';
  }

  if (subEl && comparativo.data_anterior) {
    subEl.textContent = 'vs. ' + comparativo.data_anterior.split(' às ')[0];
  }
}

function trocarAba(evt, idAba) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  if (evt && evt.target) evt.target.classList.add('active');
  document.getElementById(idAba).classList.add('active');
}

function preencherSelectCdc() {
  const select = document.getElementById('filtroCdc');
  const cdcs = [...new Map(dados.map(d => [d.codCdc + '|' + d.cdc, d])).values()]
    .sort((a,b) => (a.codCdc + ' ' + a.cdc).localeCompare(b.codCdc + ' ' + b.cdc, 'pt-BR'));
  
  select.innerHTML = '<option value="TODOS">Todos os Centros de Custo</option>' +
    cdcs.map(c => `<option value="${esc(c.codCdc + '|' + c.cdc)}">${esc(c.codCdc + ' — ' + c.cdc)}</option>`).join('');
}

function obterDadosFiltrados() {
  const op = document.getElementById('filtroOperadora').value;
  const cdc = document.getElementById('filtroCdc').value;
  return dados.filter(d => {
    const opOk = op === 'TODAS' || d.operadora === op;
    const cdcOk = cdc === 'TODOS' || (d.codCdc + '|' + d.cdc) === cdc;
    const stOk = !statusFiltro || d.status === statusFiltro;
    return opOk && cdcOk && stOk;
  });
}

function renderizarCards(itens) {
  const totais = { ATIVA: { q:0, v:0 }, ESTOQUE: { q:0, v:0 }, DESLIGADO: { q:0, v:0 }, VERIFICAR: { q:0, v:0 } };
  itens.forEach(d => {
    if (totais[d.status]) {
      totais[d.status].q++;
      totais[d.status].v += d.valor;
    }
  });
  ['Ativa', 'Estoque', 'Desligado', 'Verificar'].forEach(st => {
    const chave = st.toUpperCase();
    document.getElementById('qtd' + st).textContent = totais[chave].q.toLocaleString('pt-BR');
    document.getElementById('val' + st).textContent = fmtMoeda(totais[chave].v);
  });

  // Atualiza as flags de variação
  renderizarBadge('badgeAtiva', 'subAtiva', comparativo.ATIVA);
  renderizarBadge('badgeEstoque', 'subEstoque', comparativo.ESTOQUE);
  renderizarBadge('badgeDesligado', 'subDesligado', comparativo.DESLIGADO);
  renderizarBadge('badgeVerificar', 'subVerificar', comparativo.VERIFICAR);
}

function renderizarGraficos(itens) {
  const porCdc = new Map();
  itens.forEach(d => {
    const k = d.codCdc + ' - ' + d.cdc;
    if (!porCdc.has(k)) porCdc.set(k, { nome: k, valor: 0, pendencias: 0 });
    const c = porCdc.get(k);
    c.valor += d.valor;
    if (d.status === 'VERIFICAR' || d.status === 'DESLIGADO') c.pendencias++;
  });

  const lista = [...porCdc.values()];
  
  // Top 8 Custo
  const topCusto = lista.sort((a,b) => b.valor - a.valor).slice(0, 8);
  const maxCusto = topCusto[0]?.valor || 1;
  document.getElementById('graficoCusto').innerHTML = topCusto.map(c => `
    <div class="barra-item">
      <div class="barra-info"><span>${esc(c.nome)}</span><span>${fmtMoeda(c.valor)}</span></div>
      <div class="barra-trilha"><div class="barra-preenchimento" style="width: ${(c.valor/maxCusto)*100}%; background: var(--primary);"></div></div>
    </div>
  `).join('') || '<p style="color:var(--text-muted)">Sem dados</p>';

  // Top 8 Pendencias
  const topPend = lista.filter(c => c.pendencias > 0).sort((a,b) => b.pendencias - a.pendencias).slice(0, 8);
  const maxPend = topPend[0]?.pendencias || 1;
  document.getElementById('graficoPendencias').innerHTML = topPend.map(c => `
    <div class="barra-item">
      <div class="barra-info"><span>${esc(c.nome)}</span><span>${c.pendencias} pendência(s)</span></div>
      <div class="barra-trilha"><div class="barra-preenchimento" style="width: ${(c.pendencias/maxPend)*100}%; background: var(--verificar);"></div></div>
    </div>
  `).join('') || '<p style="color:var(--text-muted)">Nenhuma pendência encontrada</p>';
}

function renderizarResumoCdc(itens) {
  const grupos = new Map();
  itens.forEach(d => {
    const k = d.codCdc + '|' + d.cdc;
    if (!grupos.has(k)) grupos.set(k, { cod: d.codCdc, cdc: d.cdc, ATIVA:0, ESTOQUE:0, DESLIGADO:0, VERIFICAR:0, valor:0 });
    const g = grupos.get(k);
    g[d.status] = (g[d.status] || 0) + 1;
    g.valor += d.valor;
  });

  const linhas = [...grupos.values()].sort((a,b) => b.VERIFICAR - a.VERIFICAR || b.valor - a.valor);
  document.getElementById('tabelaResumoCdc').innerHTML = linhas.map(g => `
    <tr onclick="filtrarCdcTabela('${esc(g.cod + '|' + g.cdc)}')">
      <td><strong>${esc(g.cod)}</strong></td>
      <td>${esc(g.cdc)}</td>
      <td class="num">${g.ATIVA}</td>
      <td class="num">${g.ESTOQUE}</td>
      <td class="num" style="color:var(--desligado); font-weight:600;">${g.DESLIGADO}</td>
      <td class="num" style="color:var(--verificar); font-weight:600;">${g.VERIFICAR}</td>
      <td class="num" style="font-weight:600;">${fmtMoeda(g.valor)}</td>
    </tr>
  `).join('');
}

function atualizarDetalhe() {
  const busca = document.getElementById('buscaTexto').value.trim().toLowerCase();
  const itens = obterDadosFiltrados().filter(d => {
    if (!busca) return true;
    return d.nome.toLowerCase().includes(busca) ||
           d.linha.toLowerCase().includes(busca) ||
           d.chapaCpf.toLowerCase().includes(busca) ||
           d.cdc.toLowerCase().includes(busca);
  });

  document.getElementById('tabelaDetalhes').innerHTML = itens.map(d => `
    <tr>
      <td>${esc(d.operadora)}</td>
      <td><strong>${esc(d.linha)}</strong></td>
      <td>${esc(d.chapaCpf)}</td>
      <td>${esc(d.nome)}</td>
      <td>${esc(d.codCdc)}</td>
      <td>${esc(d.cdc)}</td>
      <td class="num">${fmtMoeda(d.valor)}</td>
      <td><span class="tag ${d.status.toLowerCase()}">${esc(d.status)}</span></td>
    </tr>
  `).join('') || '<tr><td colspan="8" style="text-align:center; padding:20px; color:var(--text-muted);">Nenhum registro encontrado</td></tr>';
}

function selecionarFiltroCard(st) {
  statusFiltro = (statusFiltro === st) ? null : st;
  atualizarTudo();
}

function filtrarCdcTabela(chaveCdc) {
  document.getElementById('filtroCdc').value = chaveCdc;
  trocarAba(null, 'abaResumoCdc');
  atualizarTudo();
}

function limparFiltrosDetalhe() {
  statusFiltro = null;
  document.getElementById('filtroOperadora').value = 'TODAS';
  document.getElementById('filtroCdc').value = 'TODOS';
  document.getElementById('buscaTexto').value = '';
  atualizarTudo();
}

function atualizarTudo() {
  const dadosAtuais = obterDadosFiltrados();
  renderizarCards(dadosAtuais);
  renderizarGraficos(dadosAtuais);
  renderizarResumoCdc(dadosAtuais);
  atualizarDetalhe();
}

document.getElementById('filtroOperadora').addEventListener('change', atualizarTudo);
document.getElementById('filtroCdc').addEventListener('change', atualizarTudo);

preencherSelectCdc();
atualizarTudo();
</script>
</body>
</html>"""


def gerar_html(dados: list[dict[str, object]], comparativo: dict[str, object], gerado_em: datetime) -> str:
    dados_json = json.dumps(dados, ensure_ascii=False).replace("</", "<\\/")
    comparativo_json = json.dumps(comparativo, ensure_ascii=False).replace("</", "<\\/")
    data_formatada = html.escape(gerado_em.strftime("%d/%m/%Y às %H:%M"))
    return (
        HTML_TEMPLATE
        .replace("__DADOS_JSON__", dados_json)
        .replace("__COMPARATIVO_JSON__", comparativo_json)
        .replace("__DATA_GERADO__", data_formatada)
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--telefonia", type=Path, default=ARQUIVO_PADRAO)
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = parser.parse_args()
    if not args.telefonia.exists():
        raise FileNotFoundError(f"Planilha de telefonia nao encontrada: {args.telefonia}")
    
    momento_atual = datetime.now()
    dados = carregar_dados(args.telefonia)
    comparativo = processar_historico(dados, momento_atual)
    
    args.saida.parent.mkdir(parents=True, exist_ok=True)
    args.saida.write_text(gerar_html(dados, comparativo, momento_atual), encoding="utf-8")
    
    print(f"Dashboard atualizado: {args.saida}")
    print(f"Registros considerados: {len(dados)}")
    if comparativo["tem_anterior"]:
        print(f"Comparativo com a medicao anterior ({comparativo['data_anterior']}):")
        print(f" • Variacao Custo Total: {comparativo['custo_total']['pct']}%")


if __name__ == "__main__":
    main()