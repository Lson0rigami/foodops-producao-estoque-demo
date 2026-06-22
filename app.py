from flask import Flask, render_template, request, redirect, url_for, send_file, session, has_request_context
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from PIL import Image, ImageDraw, ImageFont, ImageWin
try:
    import win32ui
except ImportError:
    win32ui = None
import sqlite3
import os
import json
import math


# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

app = Flask(__name__)
app.secret_key = "troque_essa_chave_depois_123"

DB_NAME = "database.db"
EMPADAS_POR_TABULEIRO = 35

DIAS_EXCLUSIVOS = ["Segunda-feira", "Quarta-feira"]

DIAS_SEMANA = [
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado",
    "Domingo"
]


# ============================================================
# CONFIGURAÇÕES DO MÓDULO DE COMPRAS
# ============================================================

CATEGORIAS_COMPRA_INICIAIS = [
    "Bebidas",
    "Refeitório",
    "Material de Limpeza",
    "Descartáveis",
    "Material de Escritório",
    "Insumos Secos",
    "Câmara Fria"
]

STATUS_COMPRA = ["Pendente", "Comprado", "Cancelado"]

UNIDADES_COMPRA = [
    "un",
    "kg",
    "g",
    "L",
    "ml",
    "pct",
    "cx",
    "fardo",
    "saco",
    "balde"
]

PERFIS_USUARIO = ["admin", "producao", "estoque", "compras", "loja"]

# ============================================================
# PEDIDOS DAS LOJAS / EXPEDIÇÃO - V2.8
# ============================================================

STATUS_PEDIDO_LOJA = [
    "Recebido",
    "Em separação",
    "Em expedição",
    "Em rota",
    "Finalizado",
    "Cancelado"
]

STATUS_ROMANEIO = [
    "Em preparação",
    "Conferido",
    "Em rota",
    "Finalizado",
    "Cancelado"
]

STATUS_ITEM_EXPEDICAO = [
    "Pendente",
    "Parcial",
    "Separado",
    "Indisponível"
]

ORIGENS_EXPEDICAO_INICIAIS = [
    "Produção",
    "Bebidas",
    "Descartáveis e Apoio",
    "Limpeza",
    "Escritório",
    "Almoxarifado Geral"
]

# ============================================================
# FECHAMENTO DE PEDIDOS / DEMANDA CONSOLIDADA - V2.9
# ============================================================

STATUS_RODADA_PEDIDOS = ["Aberta", "Fechada", "Planejada", "Cancelada"]
FORMAS_ABASTECIMENTO = ["Produzido internamente", "Separado diretamente do estoque"]

# ============================================================
# CUSTOS GERENCIAIS - V3.1
# ============================================================

METODOS_CUSTO = ["Automático", "Custo padrão", "Ficha técnica"]

# ============================================================
# NOTAS DE ENTRADA / EVENTOS DE COMPRA - V3.2
# ============================================================

STATUS_NOTA_ENTRADA = ["Rascunho", "Lançada"]
NATUREZAS_ITEM_NOTA = [
    "Compra normal",
    "Bonificação",
    "Sem movimentação de estoque",
]



# ============================================================
# CONFIGURAÇÕES DO MÓDULO DE ESTOQUE
# ============================================================

CATEGORIAS_ESTOQUE_INICIAIS = [
    "Bebidas",
    "Refeitório",
    "Material de Limpeza",
    "Descartáveis",
    "Material de Escritório",
    "Insumos Secos",
    "Câmara Fria",
    "Recheios",
    "Produção Própria"
]

TIPOS_MOVIMENTACAO_ESTOQUE = ["Entrada", "Saída"]
ORIGENS_ESTOQUE = [
    "Compra",
    "Entrada Manual",
    "Saída para Produção",
    "Transferência",
    "Ajuste Positivo",
    "Ajuste Negativo",
    "Perda",
    "Inventário",
    "Devolução",
    "Produção Interna",
    "Transferência para Célula",
    "Retorno de Célula",
    "Solicitação Interna",
    "Validade - Baixa",
    "Validade - Descarte",
    "Validade - Ajuste"
]

EMBALAGENS_PADRAO_EMPADAS = [
    ("Unidade", 1),
    ("Tabuleiro c/35", 35),
    ("Misto 18 unidades", 18),
    ("Misto 17 unidades", 17),
]



# ============================================================
# RASTREABILIDADE, MOTIVOS E PERDAS - V3.4
# ============================================================

def popular_motivos_perda_iniciais():
    motivos = [
        ("Vencimento", "Validade", 0, "Produto descartado por prazo de validade."),
        ("Avaria de embalagem", "Ambos", 0, "Embalagem violada, danificada ou imprópria."),
        ("Quebra ou derramamento", "Ambos", 0, "Perda física durante manuseio ou transporte."),
        ("Contaminação", "Ambos", 1, "Exige detalhamento da ocorrência."),
        ("Erro de produção", "Produção", 1, "Falha de processo, receita, cocção ou montagem."),
        ("Sobra não reaproveitável", "Produção", 0, "Sobra que não pode retornar ao processo."),
        ("Divergência de inventário", "Ambos", 1, "Diferença identificada na conferência física."),
        ("Outro", "Ambos", 1, "Motivo não previsto; observação obrigatória."),
        ("Legado / sem classificação", "Ambos", 0, "Usado somente em registros anteriores à v3.4."),
    ]
    conn = conectar(); cursor = conn.cursor()
    for nome, aplicacao, exige, obs in motivos:
        cursor.execute("""
            INSERT OR IGNORE INTO motivos_perda
            (nome, aplicacao, exige_observacao, ativo, observacao)
            VALUES (?, ?, ?, 1, ?)
        """, (nome, aplicacao, exige, obs))
    conn.commit(); conn.close()


def carregar_motivos_perda(aplicacao="", apenas_ativos=True):
    conn = conectar(); cursor = conn.cursor()
    filtros = ["1=1"]; params = []
    if apenas_ativos:
        filtros.append("ativo = 1")
    aplicacao = str(aplicacao or "").strip()
    if aplicacao:
        filtros.append("(aplicacao = ? OR aplicacao = 'Ambos')")
        params.append(aplicacao)
    cursor.execute(f"SELECT * FROM motivos_perda WHERE {' AND '.join(filtros)} ORDER BY ativo DESC, nome", params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close(); return rows


def buscar_motivo_perda_cursor(cursor, motivo_id, aplicacao=""):
    if not str(motivo_id or "").isdigit():
        return None
    cursor.execute("SELECT * FROM motivos_perda WHERE id = ?", (int(motivo_id),))
    row = cursor.fetchone()
    if not row or int(row["ativo"] or 0) != 1:
        return None
    motivo = dict(row)
    if aplicacao and motivo["aplicacao"] not in ["Ambos", aplicacao]:
        return None
    return motivo


def _registrar_perda_cursor(cursor, origem_tipo, origem_id, produto_id, quantidade, unidade,
                            motivo_id=None, observacao="", lote_id=None, consumo_id=None,
                            item_consumo_id=None, celula_id=None, data_registro=None, usuario=None):
    quantidade = float(quantidade or 0)
    if quantidade <= 0:
        return None
    custo = _calcular_custo_produto_cursor(cursor, produto_id, cache={}, pilha=set())
    custo_unitario = float(custo.get("custo_unitario", 0) or 0)
    custo_total = quantidade * custo_unitario
    cursor.execute("""
        INSERT INTO registros_perda
        (origem_tipo, origem_id, lote_id, consumo_id, item_consumo_id, produto_id,
         celula_id, motivo_id, quantidade, unidade, data_registro, usuario,
         observacao, custo_unitario_snapshot, custo_total_snapshot)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(origem_tipo, origem_id) DO UPDATE SET
            motivo_id = excluded.motivo_id,
            quantidade = excluded.quantidade,
            observacao = excluded.observacao,
            custo_unitario_snapshot = excluded.custo_unitario_snapshot,
            custo_total_snapshot = excluded.custo_total_snapshot
    """, (
        origem_tipo, int(origem_id), lote_id, consumo_id, item_consumo_id,
        int(produto_id), celula_id, motivo_id, quantidade, unidade,
        data_registro or _agora_texto(), usuario or _usuario_atual(),
        str(observacao or "").strip(), custo_unitario, custo_total
    ))
    return cursor.lastrowid


def garantir_perdas_legadas():
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("SELECT id FROM motivos_perda WHERE nome = 'Legado / sem classificação'")
    motivo = cursor.fetchone()
    motivo_id = motivo["id"] if motivo else None
    try:
        cursor.execute("BEGIN")
        cursor.execute("""
            SELECT mv.id, mv.lote_id, mv.quantidade, mv.data_movimentacao, mv.usuario,
                   mv.observacao, lv.produto_id, lv.unidade
            FROM movimentacoes_validade mv
            INNER JOIN lotes_validade lv ON lv.id = mv.lote_id
            WHERE mv.tipo = 'Descarte'
        """)
        for row in cursor.fetchall():
            _registrar_perda_cursor(
                cursor, "Validade", row["id"], row["produto_id"], row["quantidade"], row["unidade"],
                motivo_id=motivo_id, observacao=row["observacao"], lote_id=row["lote_id"],
                data_registro=row["data_movimentacao"], usuario=row["usuario"]
            )

        cursor.execute("""
            SELECT i.id, i.consumo_id, i.produto_id, i.quantidade_perda, i.unidade,
                   i.observacao, c.celula_id, c.data_confirmacao, c.confirmado_por
            FROM itens_consumo_real_producao i
            INNER JOIN consumos_reais_producao c ON c.id = i.consumo_id
            WHERE c.status = 'Confirmado' AND i.quantidade_perda > 0
        """)
        for row in cursor.fetchall():
            _registrar_perda_cursor(
                cursor, "Produção", row["id"], row["produto_id"], row["quantidade_perda"], row["unidade"],
                motivo_id=row["motivo_perda_id"] if "motivo_perda_id" in row.keys() else motivo_id,
                observacao=(row["observacao_perda"] if "observacao_perda" in row.keys() else None) or row["observacao"],
                consumo_id=row["consumo_id"], item_consumo_id=row["id"], celula_id=row["celula_id"],
                data_registro=row["data_confirmacao"], usuario=row["confirmado_por"]
            )
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()


def _parse_data_rastreabilidade(valor):
    texto = str(valor or "").strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return datetime.min


def carregar_destinos_lote(lote_id):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT rl.quantidade, r.id AS romaneio_id, r.codigo AS romaneio_codigo,
               r.status AS romaneio_status, r.data_saida, r.data_retorno,
               r.motorista, r.veiculo, r.placa, l.id AS loja_id, l.nome AS loja_nome,
               ri.descricao AS item_descricao, ri.origem_nome
        FROM romaneio_lotes rl
        INNER JOIN romaneio_itens ri ON ri.id = rl.romaneio_item_id
        INNER JOIN romaneios_loja r ON r.id = ri.romaneio_id
        INNER JOIN lojas l ON l.id = r.loja_id
        WHERE rl.lote_id = ?
        ORDER BY COALESCE(r.data_saida, r.data_criacao), r.id
    """, (lote_id,))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close(); return rows


def carregar_contexto_producao_lote(lote):
    if not lote or not lote.get("producao_realizada_id"):
        return {"producao": None, "celulas": [], "consumos": []}
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT pr.*, pp.codigo AS planejamento_codigo, pp.status AS planejamento_status,
               rp.data_entrega, rp.status AS rodada_status
        FROM producoes_realizadas pr
        LEFT JOIN planejamentos_producao pp ON pp.id = pr.planejamento_id
        LEFT JOIN rodadas_pedidos_loja rp ON rp.planejamento_id = pp.id
        WHERE pr.id = ?
    """, (lote["producao_realizada_id"],))
    producao = cursor.fetchone()
    cursor.execute("""
        SELECT DISTINCT cp.id, cp.nome, cp.centro_custo, cr.id AS consumo_id,
               cr.status, cr.data_confirmacao, cr.confirmado_por
        FROM consumos_reais_producao cr
        INNER JOIN celulas_producao cp ON cp.id = cr.celula_id
        WHERE cr.producao_realizada_id = ?
        ORDER BY cp.nome
    """, (lote["producao_realizada_id"],))
    celulas = [dict(r) for r in cursor.fetchall()]
    cursor.execute("""
        SELECT cr.id AS consumo_id, cp.nome AS celula_nome, p.codigo, p.nome AS produto_nome,
               i.quantidade_prevista, i.quantidade_utilizada, i.quantidade_perda,
               i.quantidade_devolvida, i.unidade
        FROM consumos_reais_producao cr
        INNER JOIN celulas_producao cp ON cp.id = cr.celula_id
        INNER JOIN itens_consumo_real_producao i ON i.consumo_id = cr.id
        INNER JOIN produtos_estoque p ON p.id = i.produto_id
        WHERE cr.producao_realizada_id = ? AND cr.status = 'Confirmado'
        ORDER BY cp.nome, p.nome
    """, (lote["producao_realizada_id"],))
    consumos = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return {"producao": dict(producao) if producao else None, "celulas": celulas, "consumos": consumos}


def carregar_linha_tempo_lote(lote_id):
    lote = buscar_lote_validade(lote_id)
    if not lote:
        return []
    eventos = [{
        "data": lote.get("data_criacao") or lote.get("data_producao"),
        "tipo": "Criação",
        "titulo": f"Lote {lote['codigo_lote']} criado",
        "detalhe": f"Origem: {lote.get('origem') or '-'} • Quantidade inicial: {float(lote.get('quantidade_inicial') or 0):.2f} {lote.get('unidade') or ''}",
        "usuario": lote.get("criado_por") or "-",
        "classe": "evento-criacao",
    }]
    for mov in carregar_movimentacoes_validade(lote_id):
        eventos.append({
            "data": mov["data_movimentacao"], "tipo": mov["tipo"],
            "titulo": f"{mov['tipo']}: {float(mov['quantidade'] or 0):.2f} {lote.get('unidade') or ''}",
            "detalhe": mov["observacao"] or "Movimentação sem observação.",
            "usuario": mov["usuario"] or "-", "classe": "evento-movimento",
        })
    for destino in carregar_destinos_lote(lote_id):
        eventos.append({
            "data": destino.get("data_saida") or "",
            "tipo": "Expedição",
            "titulo": f"{float(destino.get('quantidade') or 0):.2f} {lote.get('unidade') or ''} destinado(s) à loja {destino.get('loja_nome')}",
            "detalhe": f"Romaneio {destino.get('romaneio_codigo')} • {destino.get('romaneio_status')} • {destino.get('motorista') or 'Motorista não informado'}",
            "usuario": destino.get("motorista") or "-", "classe": "evento-expedicao",
            "url": f"/expedicao/romaneios/{destino.get('romaneio_id')}",
        })
    eventos.sort(key=lambda e: _parse_data_rastreabilidade(e.get("data")))
    return eventos


def carregar_lotes_rastreabilidade(busca="", status="", produto_id="", loja_id=""):
    lotes = carregar_lotes_validade(busca="", status=status, produto_id=produto_id)
    busca_lower = str(busca or "").strip().lower()
    loja_id = str(loja_id or "").strip()
    resultado = []
    for lote in lotes:
        destinos = carregar_destinos_lote(lote["id"])
        nomes_lojas = sorted({d.get("loja_nome") for d in destinos if d.get("loja_nome")})
        lote["destinos"] = destinos
        lote["lojas_destino"] = nomes_lojas
        lote["total_expedido"] = sum(float(d.get("quantidade") or 0) for d in destinos if d.get("romaneio_status") in ["Em rota", "Finalizado"])
        if loja_id and not any(str(d.get("loja_id", "")) == loja_id for d in destinos):
            # consulta antiga não retorna loja_id; usa consulta específica abaixo
            conn = conectar(); cur = conn.cursor()
            cur.execute("""SELECT 1 FROM romaneio_lotes rl JOIN romaneio_itens ri ON ri.id=rl.romaneio_item_id JOIN romaneios_loja r ON r.id=ri.romaneio_id WHERE rl.lote_id=? AND r.loja_id=? LIMIT 1""", (lote["id"], int(loja_id)))
            ok = cur.fetchone() is not None; conn.close()
            if not ok:
                continue
        if busca_lower:
            texto = " ".join([str(lote.get("codigo_lote", "")), str(lote.get("produto_nome", "")), str(lote.get("local_descricao", "")), " ".join(nomes_lojas)]).lower()
            if busca_lower not in texto:
                continue
        resultado.append(lote)
    return resultado


def carregar_relatorio_perdas(data_inicio="", data_fim="", origem_tipo="", produto_id="", motivo_id="", celula_id=""):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT rp.*, p.codigo AS produto_codigo, p.nome AS produto_nome,
               mp.nome AS motivo_nome, mp.aplicacao AS motivo_aplicacao,
               cp.nome AS celula_nome, lv.codigo_lote
        FROM registros_perda rp
        INNER JOIN produtos_estoque p ON p.id = rp.produto_id
        LEFT JOIN motivos_perda mp ON mp.id = rp.motivo_id
        LEFT JOIN celulas_producao cp ON cp.id = rp.celula_id
        LEFT JOIN lotes_validade lv ON lv.id = rp.lote_id
        ORDER BY rp.id DESC
    """)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    inicio = _data_para_iso(data_inicio) if data_inicio else ""
    fim = _data_para_iso(data_fim) if data_fim else ""
    filtrados = []
    for row in rows:
        dt = _parse_data_rastreabilidade(row.get("data_registro"))
        data_iso = dt.strftime("%Y-%m-%d") if dt != datetime.min else ""
        if inicio and data_iso < inicio: continue
        if fim and data_iso > fim: continue
        if origem_tipo and row.get("origem_tipo") != origem_tipo: continue
        if str(produto_id or "") and str(row.get("produto_id")) != str(produto_id): continue
        if str(motivo_id or "") and str(row.get("motivo_id")) != str(motivo_id): continue
        if str(celula_id or "") and str(row.get("celula_id")) != str(celula_id): continue
        filtrados.append(row)
    por_unidade = {}
    por_motivo = {}
    por_produto = {}
    total_custo = 0.0
    for row in filtrados:
        unidade = row.get("unidade") or "-"
        por_unidade[unidade] = por_unidade.get(unidade, 0) + float(row.get("quantidade") or 0)
        motivo = row.get("motivo_nome") or "Sem classificação"
        por_motivo[motivo] = por_motivo.get(motivo, 0) + float(row.get("custo_total_snapshot") or 0)
        produto = row.get("produto_nome") or "Produto"
        por_produto[produto] = por_produto.get(produto, 0) + float(row.get("custo_total_snapshot") or 0)
        total_custo += float(row.get("custo_total_snapshot") or 0)
    ranking_motivos = sorted([{"nome": k, "valor": v} for k,v in por_motivo.items()], key=lambda x:x["valor"], reverse=True)
    ranking_produtos = sorted([{"nome": k, "valor": v} for k,v in por_produto.items()], key=lambda x:x["valor"], reverse=True)
    return {
        "registros": filtrados, "total_registros": len(filtrados), "total_custo": total_custo,
        "por_unidade": por_unidade, "ranking_motivos": ranking_motivos[:8],
        "ranking_produtos": ranking_produtos[:8],
    }

# ============================================================
# LOTES E VALIDADES - V2.6
# ============================================================

STATUS_VALIDADE_FILTROS = [
    "Ativo",
    "Próximo",
    "Crítico",
    "Vence hoje",
    "Vencido",
    "Encerrado",
    "Cancelado"
]

TIPOS_MOVIMENTACAO_VALIDADE = [
    "Baixa operacional",
    "Descarte",
    "Ajuste positivo",
    "Ajuste negativo"
]

LOCAIS_VALIDADE = [
    "Estoque Central",
    "Câmara Fria",
    "Célula de Produção",
    "Expedição",
    "Loja",
    "Outro"
]


# ============================================================
# RASTREABILIDADE E PERDAS - V3.4
# ============================================================

APLICACOES_MOTIVO_PERDA = ["Ambos", "Validade", "Produção"]
ORIGENS_PERDA_FILTROS = ["Validade", "Produção"]



# ============================================================
# CÉLULAS DE PRODUÇÃO / TRANSFERÊNCIAS INTERNAS - V2.3
# ============================================================

TIPOS_TRANSFERENCIA_INTERNA = [
    "Estoque para Célula",
    "Célula para Estoque"
]

STATUS_TRANSFERENCIA_INTERNA = ["Rascunho", "Confirmada", "Cancelada"]


# ============================================================
# SOLICITAÇÕES INTERNAS / CENTROS DE CUSTO - V2.7
# ============================================================

CENTROS_CUSTO_INICIAIS = [
    "Refeitório",
    "Administrativo",
    "DP/RH",
    "Estoque",
    "Marketing",
    "Produção",
    "Cocção",
    "Produção - Suspiro",
    "Produção - Suco",
    "Limpeza Geral",
    "Manutenção",
    "Caminhão",
    "Portaria",
]

TIPOS_DESTINO_SOLICITACAO = ["Centro de Custo", "Célula"]
PRIORIDADES_SOLICITACAO = ["Normal", "Urgente"]
STATUS_SOLICITACAO_INTERNA = [
    "Pendente",
    "Em separação",
    "Pronta para confirmar",
    "Atendida",
    "Atendida parcialmente",
    "Cancelada",
]


# ============================================================
# CONFIGURAÇÕES DA IMPRESSORA ARGOX
# ============================================================

PRINTER_NAME = "Argox OS-2140 PPLA"
PRINTER_DPI = 203

ETIQUETA_LARGURA_MM = 60
ETIQUETA_ALTURA_MM = 40

# Ajustes finos da impressão.
# Se sair grande/pequena, altere apenas ESCALA_IMPRESSAO.
# Se sair virada, altere ROTACAO_IMPRESSAO para 90, -90, 180 ou 0.
ESCALA_IMPRESSAO = 1.00
ROTACAO_IMPRESSAO = 90


# ============================================================
# VALIDADE POR SABOR
# ============================================================

VALIDADE_POR_SABOR = {
    # 7 dias
    "Banana c/ Canela": 7,
    "Chocobanana": 7,
    "Pizza": 7,
    "Frango": 7,
    "Tomate Seco": 7,
    "Camarão": 7,
    "Palmito": 7,

    # 10 dias
    "Charque": 10,
    "Calabresa": 10,
    "Queijo": 10,
    "Bacalhau": 10,
    "Chocolate": 10,
    "Goiabada": 10,
    "Doce de Leite": 10,
    "Queijo do Reino": 10,
    "Dois Amores": 10
}


def obter_dias_validade(sabor):
    return VALIDADE_POR_SABOR.get(sabor, 7)


def calcular_data_validade(sabor):
    dias = obter_dias_validade(sabor)
    data_validade = datetime.now() + timedelta(days=dias)
    return data_validade.strftime("%d/%m/%Y"), dias


# ============================================================
# BANCO DE DADOS
# ============================================================

def conectar():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def criar_banco():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sabores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            classe TEXT NOT NULL,
            chuva INTEGER NOT NULL,
            normal INTEGER NOT NULL,
            verao INTEGER NOT NULL,
            baixa INTEGER NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS metas_dia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dia_semana TEXT NOT NULL,
            sabor_id INTEGER NOT NULL,
            meta INTEGER NOT NULL,
            UNIQUE(dia_semana, sabor_id),
            FOREIGN KEY (sabor_id) REFERENCES sabores(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS producoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TEXT NOT NULL,
            cenario TEXT NOT NULL,
            dia_semana TEXT NOT NULL,
            total_tabuleiros INTEGER NOT NULL,
            total_empadas INTEGER NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_producao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            producao_id INTEGER NOT NULL,
            sabor TEXT NOT NULL,
            classe TEXT NOT NULL,
            inicio INTEGER NOT NULL,
            pedido INTEGER NOT NULL,
            meta INTEGER NOT NULL,
            origem_meta TEXT NOT NULL,
            producao INTEGER NOT NULL,
            estoque_final INTEGER NOT NULL,
            empadas INTEGER NOT NULL,
            FOREIGN KEY (producao_id) REFERENCES producoes(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            senha_hash TEXT NOT NULL,
            perfil TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS categorias_compra (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fornecedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            telefone TEXT,
            email TEXT,
            observacao TEXT,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notas_entrada (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fornecedor_id INTEGER NOT NULL,
            pedido_compra_id INTEGER,
            numero TEXT NOT NULL,
            serie TEXT,
            chave_acesso TEXT,
            data_emissao TEXT,
            data_entrada TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Rascunho',
            valor_produtos REAL NOT NULL DEFAULT 0,
            desconto_geral REAL NOT NULL DEFAULT 0,
            frete REAL NOT NULL DEFAULT 0,
            outras_despesas REAL NOT NULL DEFAULT 0,
            valor_total REAL NOT NULL DEFAULT 0,
            observacao TEXT,
            criado_por TEXT NOT NULL,
            data_criacao TEXT NOT NULL,
            lancado_por TEXT,
            data_lancamento TEXT,
            FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id),
            FOREIGN KEY (pedido_compra_id) REFERENCES pedidos_compra(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_nota_entrada (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nota_id INTEGER NOT NULL,
            item_compra_id INTEGER,
            produto_id INTEGER,
            embalagem_id INTEGER,
            embalagem_nome TEXT,
            fator_conversao REAL NOT NULL DEFAULT 1,
            centro_custo_id INTEGER,
            descricao TEXT NOT NULL,
            natureza TEXT NOT NULL,
            quantidade_faturada REAL NOT NULL DEFAULT 0,
            quantidade_bonificada REAL NOT NULL DEFAULT 0,
            quantidade_total REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            valor_unitario REAL NOT NULL DEFAULT 0,
            desconto_item REAL NOT NULL DEFAULT 0,
            valor_total_item REAL NOT NULL DEFAULT 0,
            movimenta_estoque INTEGER NOT NULL DEFAULT 1,
            custo_total_entrada REAL NOT NULL DEFAULT 0,
            custo_efetivo_unitario REAL NOT NULL DEFAULT 0,
            custo_medio_anterior REAL NOT NULL DEFAULT 0,
            custo_medio_novo REAL NOT NULL DEFAULT 0,
            ultimo_custo_anterior REAL NOT NULL DEFAULT 0,
            ultimo_custo_novo REAL NOT NULL DEFAULT 0,
            movimentacao_estoque_id INTEGER,
            reconciliado_movimento_existente INTEGER NOT NULL DEFAULT 0,
            observacao TEXT,
            FOREIGN KEY (nota_id) REFERENCES notas_entrada(id),
            FOREIGN KEY (item_compra_id) REFERENCES itens_compra(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (embalagem_id) REFERENCES produto_embalagens(id),
            FOREIGN KEY (centro_custo_id) REFERENCES centros_custo(id),
            FOREIGN KEY (movimentacao_estoque_id) REFERENCES movimentacoes_estoque(id)
        )
    """)

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_notas_entrada_fornecedor ON notas_entrada(fornecedor_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_notas_entrada_status ON notas_entrada(status)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_itens_nota_nota ON itens_nota_entrada(nota_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_itens_nota_produto ON itens_nota_entrada(produto_id)")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pedidos_compra (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            solicitante TEXT NOT NULL,
            data_solicitacao TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Pendente',
            observacao TEXT,
            fornecedor_id INTEGER,
            data_compra TEXT,
            observacao_compra TEXT,
            comprado_por TEXT,
            data_cancelamento TEXT,
            cancelado_por TEXT,
            observacao_cancelamento TEXT,
            FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_compra (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido_id INTEGER NOT NULL,
            descricao TEXT NOT NULL,
            quantidade REAL NOT NULL,
            unidade TEXT NOT NULL,
            categoria_id INTEGER NOT NULL,
            observacao TEXT,
            status TEXT NOT NULL DEFAULT 'Pendente',
            fornecedor_id INTEGER,
            data_compra TEXT,
            observacao_compra TEXT,
            comprado_por TEXT,
            produto_estoque_id INTEGER,
            estoque_movimentacao_id INTEGER,
            nota_entrada_item_id INTEGER,
            FOREIGN KEY (pedido_id) REFERENCES pedidos_compra(id),
            FOREIGN KEY (categoria_id) REFERENCES categorias_compra(id),
            FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id),
            FOREIGN KEY (produto_estoque_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (estoque_movimentacao_id) REFERENCES movimentacoes_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS categorias_estoque (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS produtos_estoque (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            nome TEXT NOT NULL UNIQUE,
            categoria_id INTEGER,
            unidade_padrao TEXT NOT NULL,
            estoque_minimo REAL NOT NULL DEFAULT 0,
            custo_padrao REAL NOT NULL DEFAULT 0,
            custo_medio REAL NOT NULL DEFAULT 0,
            ultimo_custo REAL NOT NULL DEFAULT 0,
            data_ultimo_custo TEXT,
            metodo_custo TEXT NOT NULL DEFAULT 'Automático',
            ativo INTEGER NOT NULL DEFAULT 1,
            ativo_venda INTEGER NOT NULL DEFAULT 0,
            controla_validade INTEGER NOT NULL DEFAULT 0,
            dias_validade INTEGER,
            data_cadastro TEXT NOT NULL,
            observacao TEXT,
            FOREIGN KEY (categoria_id) REFERENCES categorias_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS produto_embalagens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            produto_id INTEGER NOT NULL,
            nome TEXT NOT NULL,
            fator_conversao REAL NOT NULL DEFAULT 1,
            padrao INTEGER NOT NULL DEFAULT 0,
            ativo INTEGER NOT NULL DEFAULT 1,
            observacao TEXT,
            UNIQUE(produto_id, nome),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS movimentacoes_estoque (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            produto_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            quantidade REAL NOT NULL,
            quantidade_embalagem REAL,
            embalagem_id INTEGER,
            embalagem_nome TEXT,
            fator_conversao REAL,
            data_movimentacao TEXT NOT NULL,
            origem TEXT NOT NULL,
            origem_id INTEGER,
            item_compra_id INTEGER,
            nota_entrada_item_id INTEGER,
            fornecedor_id INTEGER,
            usuario TEXT NOT NULL,
            observacao TEXT,
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (embalagem_id) REFERENCES produto_embalagens(id),
            FOREIGN KEY (item_compra_id) REFERENCES itens_compra(id),
            FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS producoes_realizadas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_criacao TEXT NOT NULL,
            cenario TEXT NOT NULL,
            dia_semana TEXT NOT NULL,
            total_tabuleiros REAL NOT NULL DEFAULT 0,
            total_unidades REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Pendente',
            origem TEXT NOT NULL DEFAULT 'Etiquetas',
            criado_por TEXT NOT NULL,
            data_confirmacao TEXT,
            confirmado_por TEXT,
            data_cancelamento TEXT,
            cancelado_por TEXT,
            observacao TEXT,
            planejamento_id INTEGER,
            FOREIGN KEY (planejamento_id) REFERENCES planejamentos_producao(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_producao_realizada (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            producao_realizada_id INTEGER NOT NULL,
            sabor TEXT NOT NULL,
            produto_estoque_id INTEGER,
            quantidade_tabuleiros REAL NOT NULL DEFAULT 0,
            quantidade_unidades REAL NOT NULL DEFAULT 0,
            embalagem_id INTEGER,
            embalagem_nome TEXT,
            fator_conversao REAL,
            status TEXT NOT NULL DEFAULT 'Pendente',
            estoque_movimentacao_id INTEGER,
            observacao TEXT,
            FOREIGN KEY (producao_realizada_id) REFERENCES producoes_realizadas(id),
            FOREIGN KEY (produto_estoque_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (embalagem_id) REFERENCES produto_embalagens(id),
            FOREIGN KEY (estoque_movimentacao_id) REFERENCES movimentacoes_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fichas_tecnicas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            produto_final_id INTEGER NOT NULL UNIQUE,
            tipo TEXT NOT NULL DEFAULT 'Produto acabado',
            rendimento_quantidade REAL NOT NULL DEFAULT 1,
            rendimento_unidade TEXT NOT NULL DEFAULT 'un',
            ativo INTEGER NOT NULL DEFAULT 1,
            versao INTEGER NOT NULL DEFAULT 1,
            data_cadastro TEXT NOT NULL,
            atualizado_em TEXT,
            usuario TEXT,
            observacao TEXT,
            FOREIGN KEY (produto_final_id) REFERENCES produtos_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_ficha_tecnica (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ficha_id INTEGER NOT NULL,
            insumo_produto_id INTEGER NOT NULL,
            quantidade REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            ativo INTEGER NOT NULL DEFAULT 1,
            atualizado_em TEXT,
            usuario TEXT,
            observacao TEXT,
            FOREIGN KEY (ficha_id) REFERENCES fichas_tecnicas(id),
            FOREIGN KEY (insumo_produto_id) REFERENCES produtos_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS ficha_tecnica_revisoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ficha_id INTEGER NOT NULL,
            versao INTEGER NOT NULL,
            acao TEXT NOT NULL,
            dados_json TEXT NOT NULL,
            data_hora TEXT NOT NULL,
            usuario TEXT NOT NULL,
            UNIQUE(ficha_id, versao),
            FOREIGN KEY (ficha_id) REFERENCES fichas_tecnicas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS ficha_tecnica_utilizacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ficha_id INTEGER NOT NULL,
            versao_ficha INTEGER NOT NULL,
            origem TEXT NOT NULL,
            origem_id INTEGER,
            dados_json TEXT NOT NULL,
            data_hora TEXT NOT NULL,
            usuario TEXT NOT NULL,
            observacao TEXT,
            FOREIGN KEY (ficha_id) REFERENCES fichas_tecnicas(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS planejamentos_producao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_criacao TEXT NOT NULL,
            cenario TEXT NOT NULL,
            dia_semana TEXT NOT NULL,
            total_tabuleiros REAL NOT NULL DEFAULT 0,
            total_unidades REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Rascunho',
            origem TEXT NOT NULL DEFAULT 'Meta de Produção',
            criado_por TEXT NOT NULL,
            producao_realizada_id INTEGER,
            observacao TEXT,
            FOREIGN KEY (producao_realizada_id) REFERENCES producoes_realizadas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_planejamento_producao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            planejamento_id INTEGER NOT NULL,
            sabor TEXT NOT NULL,
            produto_final_id INTEGER,
            quantidade_tabuleiros REAL NOT NULL DEFAULT 0,
            quantidade_unidades REAL NOT NULL DEFAULT 0,
            ficha_id INTEGER,
            versao_ficha INTEGER,
            utilizacao_ficha_id INTEGER,
            status TEXT NOT NULL DEFAULT 'Planejado',
            alerta TEXT,
            snapshot_json TEXT,
            observacao TEXT,
            FOREIGN KEY (planejamento_id) REFERENCES planejamentos_producao(id),
            FOREIGN KEY (produto_final_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (ficha_id) REFERENCES fichas_tecnicas(id),
            FOREIGN KEY (utilizacao_ficha_id) REFERENCES ficha_tecnica_utilizacoes(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS consumos_previstos_planejamento (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            planejamento_id INTEGER NOT NULL,
            item_planejamento_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            quantidade REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            nivel INTEGER NOT NULL DEFAULT 1,
            caminho TEXT,
            ficha_origem_id INTEGER,
            versao_ficha_origem INTEGER,
            ficha_componente_id INTEGER,
            versao_ficha_componente INTEGER,
            alerta TEXT,
            observacao TEXT,
            FOREIGN KEY (planejamento_id) REFERENCES planejamentos_producao(id),
            FOREIGN KEY (item_planejamento_id) REFERENCES itens_planejamento_producao(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (ficha_origem_id) REFERENCES fichas_tecnicas(id),
            FOREIGN KEY (ficha_componente_id) REFERENCES fichas_tecnicas(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS separacoes_estoque (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            planejamento_id INTEGER NOT NULL UNIQUE,
            data_criacao TEXT NOT NULL,
            data_atualizacao TEXT,
            status TEXT NOT NULL DEFAULT 'Pendente',
            criado_por TEXT NOT NULL,
            atualizado_por TEXT,
            data_conclusao TEXT,
            concluido_por TEXT,
            observacao TEXT,
            FOREIGN KEY (planejamento_id) REFERENCES planejamentos_producao(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_separacao_estoque (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            separacao_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            quantidade_prevista REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            quantidade_separada REAL NOT NULL DEFAULT 0,
            saldo_no_planejamento REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Pendente',
            alerta TEXT,
            observacao TEXT,
            atualizado_em TEXT,
            atualizado_por TEXT,
            UNIQUE(separacao_id, produto_id, unidade),
            FOREIGN KEY (separacao_id) REFERENCES separacoes_estoque(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS celulas_producao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            descricao TEXT,
            centro_custo TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            data_cadastro TEXT NOT NULL,
            criado_por TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transferencias_internas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo TEXT NOT NULL,
            celula_origem_id INTEGER,
            celula_destino_id INTEGER,
            separacao_id INTEGER,
            status TEXT NOT NULL DEFAULT 'Rascunho',
            data_criacao TEXT NOT NULL,
            criado_por TEXT NOT NULL,
            data_confirmacao TEXT,
            confirmado_por TEXT,
            data_cancelamento TEXT,
            cancelado_por TEXT,
            observacao TEXT,
            FOREIGN KEY (celula_origem_id) REFERENCES celulas_producao(id),
            FOREIGN KEY (celula_destino_id) REFERENCES celulas_producao(id),
            FOREIGN KEY (separacao_id) REFERENCES separacoes_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_transferencia_interna (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transferencia_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            quantidade REAL NOT NULL,
            unidade TEXT NOT NULL,
            separacao_item_id INTEGER,
            estoque_movimentacao_id INTEGER,
            celula_movimentacao_id INTEGER,
            observacao TEXT,
            FOREIGN KEY (transferencia_id) REFERENCES transferencias_internas(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (separacao_item_id) REFERENCES itens_separacao_estoque(id),
            FOREIGN KEY (estoque_movimentacao_id) REFERENCES movimentacoes_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS movimentacoes_celula (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            celula_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            quantidade REAL NOT NULL,
            data_movimentacao TEXT NOT NULL,
            origem TEXT NOT NULL,
            origem_id INTEGER,
            usuario TEXT NOT NULL,
            observacao TEXT,
            FOREIGN KEY (celula_id) REFERENCES celulas_producao(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS centros_custo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            descricao TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            data_cadastro TEXT NOT NULL,
            criado_por TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS solicitacoes_internas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            solicitante TEXT NOT NULL,
            destino_tipo TEXT NOT NULL,
            centro_custo_id INTEGER,
            celula_id INTEGER,
            data_solicitacao TEXT NOT NULL,
            data_necessidade TEXT,
            prioridade TEXT NOT NULL DEFAULT 'Normal',
            status TEXT NOT NULL DEFAULT 'Pendente',
            observacao TEXT,
            analisado_por TEXT,
            data_analise TEXT,
            confirmado_por TEXT,
            data_confirmacao TEXT,
            cancelado_por TEXT,
            data_cancelamento TEXT,
            motivo_cancelamento TEXT,
            custo_total REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (centro_custo_id) REFERENCES centros_custo(id),
            FOREIGN KEY (celula_id) REFERENCES celulas_producao(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_solicitacao_interna (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            solicitacao_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            embalagem_id INTEGER,
            embalagem_nome TEXT,
            quantidade_solicitada_embalagem REAL NOT NULL DEFAULT 0,
            fator_conversao REAL NOT NULL DEFAULT 1,
            quantidade_solicitada_base REAL NOT NULL DEFAULT 0,
            quantidade_separada_base REAL NOT NULL DEFAULT 0,
            quantidade_atendida_base REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Pendente',
            recusado INTEGER NOT NULL DEFAULT 0,
            motivo_recusa TEXT,
            observacao TEXT,
            custo_unitario_snapshot REAL NOT NULL DEFAULT 0,
            custo_total_snapshot REAL NOT NULL DEFAULT 0,
            estoque_movimentacao_id INTEGER,
            celula_movimentacao_id INTEGER,
            atualizado_em TEXT,
            atualizado_por TEXT,
            FOREIGN KEY (solicitacao_id) REFERENCES solicitacoes_internas(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (embalagem_id) REFERENCES produto_embalagens(id),
            FOREIGN KEY (estoque_movimentacao_id) REFERENCES movimentacoes_estoque(id),
            FOREIGN KEY (celula_movimentacao_id) REFERENCES movimentacoes_celula(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS consumos_reais_producao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            producao_realizada_id INTEGER NOT NULL,
            planejamento_id INTEGER,
            celula_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'Rascunho',
            data_criacao TEXT NOT NULL,
            criado_por TEXT NOT NULL,
            data_confirmacao TEXT,
            confirmado_por TEXT,
            data_cancelamento TEXT,
            cancelado_por TEXT,
            observacao TEXT,
            UNIQUE(producao_realizada_id, celula_id),
            FOREIGN KEY (producao_realizada_id) REFERENCES producoes_realizadas(id),
            FOREIGN KEY (planejamento_id) REFERENCES planejamentos_producao(id),
            FOREIGN KEY (celula_id) REFERENCES celulas_producao(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_consumo_real_producao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            consumo_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            origem TEXT NOT NULL DEFAULT 'Planejamento',
            quantidade_prevista REAL NOT NULL DEFAULT 0,
            quantidade_recebida REAL NOT NULL DEFAULT 0,
            saldo_celula_inicial REAL NOT NULL DEFAULT 0,
            quantidade_utilizada REAL NOT NULL DEFAULT 0,
            quantidade_perda REAL NOT NULL DEFAULT 0,
            quantidade_devolvida REAL NOT NULL DEFAULT 0,
            saldo_celula_final REAL,
            unidade TEXT NOT NULL,
            movimento_uso_id INTEGER,
            movimento_perda_id INTEGER,
            transferencia_retorno_id INTEGER,
            observacao TEXT,
            UNIQUE(consumo_id, produto_id),
            FOREIGN KEY (consumo_id) REFERENCES consumos_reais_producao(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (movimento_uso_id) REFERENCES movimentacoes_celula(id),
            FOREIGN KEY (movimento_perda_id) REFERENCES movimentacoes_celula(id),
            FOREIGN KEY (transferencia_retorno_id) REFERENCES transferencias_internas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS lotes_validade (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_lote TEXT NOT NULL UNIQUE,
            produto_id INTEGER NOT NULL,
            producao_realizada_id INTEGER,
            item_producao_realizada_id INTEGER UNIQUE,
            data_producao TEXT NOT NULL,
            data_validade TEXT NOT NULL,
            quantidade_inicial REAL NOT NULL DEFAULT 0,
            quantidade_atual REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            local_tipo TEXT NOT NULL DEFAULT 'Estoque Central',
            celula_id INTEGER,
            local_descricao TEXT,
            status TEXT NOT NULL DEFAULT 'Ativo',
            origem TEXT NOT NULL DEFAULT 'Produção Confirmada',
            criado_por TEXT NOT NULL,
            data_criacao TEXT NOT NULL,
            atualizado_em TEXT,
            encerrado_em TEXT,
            observacao TEXT,
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (producao_realizada_id) REFERENCES producoes_realizadas(id),
            FOREIGN KEY (item_producao_realizada_id) REFERENCES itens_producao_realizada(id),
            FOREIGN KEY (celula_id) REFERENCES celulas_producao(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS movimentacoes_validade (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lote_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            quantidade REAL NOT NULL,
            data_movimentacao TEXT NOT NULL,
            usuario TEXT NOT NULL,
            estoque_movimentacao_id INTEGER,
            observacao TEXT,
            FOREIGN KEY (lote_id) REFERENCES lotes_validade(id),
            FOREIGN KEY (estoque_movimentacao_id) REFERENCES movimentacoes_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS lojas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            nome TEXT NOT NULL UNIQUE,
            ativo INTEGER NOT NULL DEFAULT 1,
            observacao TEXT,
            data_cadastro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS origens_expedicao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            ativo INTEGER NOT NULL DEFAULT 1,
            observacao TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS motoristas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            telefone TEXT,
            categoria_cnh TEXT,
            validade_cnh TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            observacao TEXT,
            data_cadastro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS veiculos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            descricao TEXT NOT NULL,
            placa TEXT NOT NULL UNIQUE,
            tipo TEXT,
            capacidade TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            observacao TEXT,
            data_cadastro TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pedidos_loja (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loja_id INTEGER NOT NULL,
            criado_por TEXT NOT NULL,
            data_criacao TEXT NOT NULL,
            data_entrega_desejada TEXT,
            status TEXT NOT NULL DEFAULT 'Recebido',
            observacao TEXT,
            data_cancelamento TEXT,
            cancelado_por TEXT,
            motivo_cancelamento TEXT,
            data_finalizacao TEXT,
            finalizado_por TEXT,
            FOREIGN KEY (loja_id) REFERENCES lojas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_pedido_loja (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido_id INTEGER NOT NULL,
            tipo_item TEXT NOT NULL DEFAULT 'Normal',
            produto_id INTEGER,
            embalagem_id INTEGER,
            embalagem_nome TEXT,
            fator_conversao REAL,
            quantidade_comercial REAL NOT NULL DEFAULT 0,
            quantidade_base REAL NOT NULL DEFAULT 0,
            sabor_1_id INTEGER,
            sabor_2_id INTEGER,
            quantidade_sabor_1_base REAL NOT NULL DEFAULT 0,
            quantidade_sabor_2_base REAL NOT NULL DEFAULT 0,
            origem_expedicao_id INTEGER,
            status_expedicao TEXT NOT NULL DEFAULT 'Pendente',
            quantidade_separada_comercial REAL NOT NULL DEFAULT 0,
            quantidade_separada_base REAL NOT NULL DEFAULT 0,
            observacao_loja TEXT,
            observacao_expedicao TEXT,
            data_separacao TEXT,
            separado_por TEXT,
            movimento_estoque_id INTEGER,
            movimento_estoque_2_id INTEGER,
            FOREIGN KEY (pedido_id) REFERENCES pedidos_loja(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (embalagem_id) REFERENCES produto_embalagens(id),
            FOREIGN KEY (sabor_1_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (sabor_2_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (origem_expedicao_id) REFERENCES origens_expedicao(id),
            FOREIGN KEY (movimento_estoque_id) REFERENCES movimentacoes_estoque(id),
            FOREIGN KEY (movimento_estoque_2_id) REFERENCES movimentacoes_estoque(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS rodadas_pedidos_loja (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_entrega TEXT NOT NULL UNIQUE,
            status TEXT NOT NULL DEFAULT 'Aberta',
            data_criacao TEXT NOT NULL,
            criado_por TEXT NOT NULL,
            data_fechamento TEXT,
            fechado_por TEXT,
            cenario TEXT,
            dia_semana TEXT,
            observacao TEXT,
            planejamento_id INTEGER,
            data_reabertura TEXT,
            reaberto_por TEXT,
            FOREIGN KEY (planejamento_id) REFERENCES planejamentos_producao(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS demanda_producao_rodada (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rodada_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            quantidade_normal_unidades REAL NOT NULL DEFAULT 0,
            quantidade_misto_unidades REAL NOT NULL DEFAULT 0,
            quantidade_total_unidades REAL NOT NULL DEFAULT 0,
            tabuleiros_normais REAL NOT NULL DEFAULT 0,
            observacao TEXT,
            UNIQUE(rodada_id, produto_id),
            FOREIGN KEY (rodada_id) REFERENCES rodadas_pedidos_loja(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mistos_rodada (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rodada_id INTEGER NOT NULL,
            pedido_id INTEGER NOT NULL,
            item_pedido_id INTEGER NOT NULL,
            loja_id INTEGER NOT NULL,
            sabor_1_id INTEGER NOT NULL,
            sabor_2_id INTEGER NOT NULL,
            quantidade_tabuleiros REAL NOT NULL DEFAULT 0,
            quantidade_sabor_1 REAL NOT NULL DEFAULT 0,
            quantidade_sabor_2 REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (rodada_id) REFERENCES rodadas_pedidos_loja(id),
            FOREIGN KEY (pedido_id) REFERENCES pedidos_loja(id),
            FOREIGN KEY (item_pedido_id) REFERENCES itens_pedido_loja(id),
            FOREIGN KEY (loja_id) REFERENCES lojas(id),
            FOREIGN KEY (sabor_1_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (sabor_2_id) REFERENCES produtos_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS romaneios_loja (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            pedido_id INTEGER NOT NULL UNIQUE,
            rodada_id INTEGER,
            loja_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'Em preparação',
            data_criacao TEXT NOT NULL,
            criado_por TEXT NOT NULL,
            data_conferencia TEXT,
            conferido_por TEXT,
            motorista_id INTEGER,
            veiculo_id INTEGER,
            motorista TEXT,
            veiculo TEXT,
            placa TEXT,
            data_saida TEXT,
            saida_por TEXT,
            data_retorno TEXT,
            retorno_por TEXT,
            recebido_loja_por TEXT,
            observacao TEXT,
            observacao_retorno TEXT,
            data_finalizacao TEXT,
            custo_total_snapshot REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (pedido_id) REFERENCES pedidos_loja(id),
            FOREIGN KEY (rodada_id) REFERENCES rodadas_pedidos_loja(id),
            FOREIGN KEY (loja_id) REFERENCES lojas(id),
            FOREIGN KEY (motorista_id) REFERENCES motoristas(id),
            FOREIGN KEY (veiculo_id) REFERENCES veiculos(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS romaneio_itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            romaneio_id INTEGER NOT NULL,
            item_pedido_id INTEGER NOT NULL,
            tipo_item TEXT NOT NULL,
            produto_id INTEGER,
            sabor_1_id INTEGER,
            sabor_2_id INTEGER,
            descricao TEXT NOT NULL,
            origem_nome TEXT,
            embalagem_nome TEXT,
            fator_conversao REAL NOT NULL DEFAULT 1,
            quantidade_pedida_comercial REAL NOT NULL DEFAULT 0,
            quantidade_pedida_base REAL NOT NULL DEFAULT 0,
            quantidade_separada_comercial REAL NOT NULL DEFAULT 0,
            quantidade_separada_base REAL NOT NULL DEFAULT 0,
            quantidade_conferida_comercial REAL NOT NULL DEFAULT 0,
            quantidade_conferida_base REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Aguardando conferência',
            observacao TEXT,
            UNIQUE(romaneio_id, item_pedido_id),
            FOREIGN KEY (romaneio_id) REFERENCES romaneios_loja(id),
            FOREIGN KEY (item_pedido_id) REFERENCES itens_pedido_loja(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (sabor_1_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (sabor_2_id) REFERENCES produtos_estoque(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS romaneio_lotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            romaneio_item_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            lote_id INTEGER NOT NULL,
            quantidade REAL NOT NULL,
            FOREIGN KEY (romaneio_item_id) REFERENCES romaneio_itens(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (lote_id) REFERENCES lotes_validade(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS romaneio_custos_itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            romaneio_id INTEGER NOT NULL,
            romaneio_item_id INTEGER NOT NULL,
            produto_id INTEGER NOT NULL,
            quantidade_base REAL NOT NULL DEFAULT 0,
            unidade TEXT NOT NULL,
            custo_unitario_snapshot REAL NOT NULL DEFAULT 0,
            custo_total_snapshot REAL NOT NULL DEFAULT 0,
            origem_custo TEXT NOT NULL DEFAULT 'Sem custo',
            ficha_id INTEGER,
            ficha_versao INTEGER,
            observacao TEXT,
            data_snapshot TEXT NOT NULL,
            UNIQUE(romaneio_item_id, produto_id),
            FOREIGN KEY (romaneio_id) REFERENCES romaneios_loja(id),
            FOREIGN KEY (romaneio_item_id) REFERENCES romaneio_itens(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (ficha_id) REFERENCES fichas_tecnicas(id)
        )
    """)


    cursor.execute("""
        CREATE TABLE IF NOT EXISTS motivos_perda (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            aplicacao TEXT NOT NULL DEFAULT 'Ambos',
            exige_observacao INTEGER NOT NULL DEFAULT 0,
            ativo INTEGER NOT NULL DEFAULT 1,
            observacao TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS registros_perda (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            origem_tipo TEXT NOT NULL,
            origem_id INTEGER NOT NULL,
            lote_id INTEGER,
            consumo_id INTEGER,
            item_consumo_id INTEGER,
            produto_id INTEGER NOT NULL,
            celula_id INTEGER,
            motivo_id INTEGER,
            quantidade REAL NOT NULL,
            unidade TEXT NOT NULL,
            data_registro TEXT NOT NULL,
            usuario TEXT NOT NULL,
            observacao TEXT,
            custo_unitario_snapshot REAL NOT NULL DEFAULT 0,
            custo_total_snapshot REAL NOT NULL DEFAULT 0,
            UNIQUE(origem_tipo, origem_id),
            FOREIGN KEY (lote_id) REFERENCES lotes_validade(id),
            FOREIGN KEY (consumo_id) REFERENCES consumos_reais_producao(id),
            FOREIGN KEY (item_consumo_id) REFERENCES itens_consumo_real_producao(id),
            FOREIGN KEY (produto_id) REFERENCES produtos_estoque(id),
            FOREIGN KEY (celula_id) REFERENCES celulas_producao(id),
            FOREIGN KEY (motivo_id) REFERENCES motivos_perda(id)
        )
    """)

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_registros_perda_data ON registros_perda(data_registro)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_registros_perda_produto ON registros_perda(produto_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_registros_perda_lote ON registros_perda(lote_id)")

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_romaneios_status ON romaneios_loja(status)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_romaneio_lotes_item ON romaneio_lotes(romaneio_item_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_romaneio_custos_romaneio ON romaneio_custos_itens(romaneio_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_romaneio_custos_produto ON romaneio_custos_itens(produto_id)")

    conn.commit()
    conn.close()


def garantir_colunas_novas():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("PRAGMA table_info(itens_compra)")
    colunas_itens_compra = [row["name"] for row in cursor.fetchall()]

    if "produto_estoque_id" not in colunas_itens_compra:
        cursor.execute("ALTER TABLE itens_compra ADD COLUMN produto_estoque_id INTEGER")

    if "estoque_movimentacao_id" not in colunas_itens_compra:
        cursor.execute("ALTER TABLE itens_compra ADD COLUMN estoque_movimentacao_id INTEGER")
    if "nota_entrada_item_id" not in colunas_itens_compra:
        cursor.execute("ALTER TABLE itens_compra ADD COLUMN nota_entrada_item_id INTEGER")

    cursor.execute("PRAGMA table_info(produtos_estoque)")
    colunas_produtos_estoque = [row["name"] for row in cursor.fetchall()]

    if "codigo" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN codigo TEXT")

    if "custo_padrao" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN custo_padrao REAL NOT NULL DEFAULT 0")

    if "custo_medio" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN custo_medio REAL NOT NULL DEFAULT 0")
    if "ultimo_custo" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN ultimo_custo REAL NOT NULL DEFAULT 0")
    if "data_ultimo_custo" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN data_ultimo_custo TEXT")
    if "metodo_custo" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN metodo_custo TEXT NOT NULL DEFAULT 'Automático'")

    if "ativo_venda" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN ativo_venda INTEGER NOT NULL DEFAULT 0")

    if "controla_validade" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN controla_validade INTEGER NOT NULL DEFAULT 0")

    if "dias_validade" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN dias_validade INTEGER")

    if "origem_expedicao_id" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN origem_expedicao_id INTEGER")

    if "ordem_loja" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN ordem_loja INTEGER NOT NULL DEFAULT 0")


    forma_abastecimento_criada = False
    if "forma_abastecimento" not in colunas_produtos_estoque:
        cursor.execute("ALTER TABLE produtos_estoque ADD COLUMN forma_abastecimento TEXT NOT NULL DEFAULT 'Separado diretamente do estoque'")
        forma_abastecimento_criada = True

    cursor.execute("PRAGMA table_info(produto_embalagens)")
    colunas_produto_embalagens = [row["name"] for row in cursor.fetchall()]

    if "disponivel_loja" not in colunas_produto_embalagens:
        cursor.execute("ALTER TABLE produto_embalagens ADD COLUMN disponivel_loja INTEGER NOT NULL DEFAULT 0")

    cursor.execute("PRAGMA table_info(pedidos_loja)")
    colunas_pedidos_loja = [row["name"] for row in cursor.fetchall()]

    if "rodada_id" not in colunas_pedidos_loja:
        cursor.execute("ALTER TABLE pedidos_loja ADD COLUMN rodada_id INTEGER")

    cursor.execute("PRAGMA table_info(planejamentos_producao)")
    colunas_planejamentos = [row["name"] for row in cursor.fetchall()]

    if "rodada_id" not in colunas_planejamentos:
        cursor.execute("ALTER TABLE planejamentos_producao ADD COLUMN rodada_id INTEGER")

    cursor.execute("PRAGMA table_info(usuarios)")
    colunas_usuarios = [row["name"] for row in cursor.fetchall()]

    if "loja_id" not in colunas_usuarios:
        cursor.execute("ALTER TABLE usuarios ADD COLUMN loja_id INTEGER")

    cursor.execute("PRAGMA table_info(movimentacoes_estoque)")
    colunas_movimentacoes_estoque = [row["name"] for row in cursor.fetchall()]

    if "quantidade_embalagem" not in colunas_movimentacoes_estoque:
        cursor.execute("ALTER TABLE movimentacoes_estoque ADD COLUMN quantidade_embalagem REAL")

    if "embalagem_id" not in colunas_movimentacoes_estoque:
        cursor.execute("ALTER TABLE movimentacoes_estoque ADD COLUMN embalagem_id INTEGER")

    if "embalagem_nome" not in colunas_movimentacoes_estoque:
        cursor.execute("ALTER TABLE movimentacoes_estoque ADD COLUMN embalagem_nome TEXT")

    if "fator_conversao" not in colunas_movimentacoes_estoque:
        cursor.execute("ALTER TABLE movimentacoes_estoque ADD COLUMN fator_conversao REAL")
    if "nota_entrada_item_id" not in colunas_movimentacoes_estoque:
        cursor.execute("ALTER TABLE movimentacoes_estoque ADD COLUMN nota_entrada_item_id INTEGER")

    cursor.execute("PRAGMA table_info(producoes_realizadas)")
    colunas_producoes_realizadas = [row["name"] for row in cursor.fetchall()]

    if "planejamento_id" not in colunas_producoes_realizadas:
        cursor.execute("ALTER TABLE producoes_realizadas ADD COLUMN planejamento_id INTEGER")

    cursor.execute("PRAGMA table_info(fichas_tecnicas)")
    colunas_fichas_tecnicas = [row["name"] for row in cursor.fetchall()]

    if "versao" not in colunas_fichas_tecnicas:
        cursor.execute("ALTER TABLE fichas_tecnicas ADD COLUMN versao INTEGER NOT NULL DEFAULT 1")

    cursor.execute("PRAGMA table_info(itens_ficha_tecnica)")
    colunas_itens_ficha = [row["name"] for row in cursor.fetchall()]

    if "atualizado_em" not in colunas_itens_ficha:
        cursor.execute("ALTER TABLE itens_ficha_tecnica ADD COLUMN atualizado_em TEXT")

    if "usuario" not in colunas_itens_ficha:
        cursor.execute("ALTER TABLE itens_ficha_tecnica ADD COLUMN usuario TEXT")

    cursor.execute("SELECT id FROM produtos_estoque WHERE codigo IS NULL OR codigo = '' ORDER BY id")
    produtos_sem_codigo = cursor.fetchall()

    for produto in produtos_sem_codigo:
        cursor.execute("""
            SELECT codigo
            FROM produtos_estoque
            WHERE codigo IS NOT NULL AND codigo != ''
            ORDER BY CAST(codigo AS INTEGER) DESC
            LIMIT 1
        """)
        ultimo = cursor.fetchone()
        proximo_numero = int(ultimo["codigo"]) + 1 if ultimo and str(ultimo["codigo"]).isdigit() else 1
        codigo = str(proximo_numero).zfill(4)

        cursor.execute("""
            UPDATE produtos_estoque
            SET codigo = ?
            WHERE id = ?
        """, (codigo, produto["id"]))

    # Configura automaticamente a validade dos produtos de empada já existentes,
    # sem sobrescrever configurações definidas manualmente.
    for sabor, dias in VALIDADE_POR_SABOR.items():
        cursor.execute("""
            UPDATE produtos_estoque
            SET controla_validade = 1,
                dias_validade = COALESCE(dias_validade, ?)
            WHERE nome = ?
        """, (dias, f"Empada de {sabor}"))

    if forma_abastecimento_criada:
        cursor.execute("""
            UPDATE produtos_estoque
            SET forma_abastecimento = 'Produzido internamente'
            WHERE nome LIKE 'Empada de %'
               OR categoria_id IN (SELECT id FROM categorias_estoque WHERE nome = 'Produção Própria')
        """)

    cursor.execute("PRAGMA table_info(romaneios_loja)")
    colunas_romaneios = [row["name"] for row in cursor.fetchall()]
    if "motorista_id" not in colunas_romaneios:
        cursor.execute("ALTER TABLE romaneios_loja ADD COLUMN motorista_id INTEGER")
    if "veiculo_id" not in colunas_romaneios:
        cursor.execute("ALTER TABLE romaneios_loja ADD COLUMN veiculo_id INTEGER")
    if "custo_total_snapshot" not in colunas_romaneios:
        cursor.execute("ALTER TABLE romaneios_loja ADD COLUMN custo_total_snapshot REAL NOT NULL DEFAULT 0")

    cursor.execute("PRAGMA table_info(movimentacoes_validade)")
    colunas_mov_validade = [row["name"] for row in cursor.fetchall()]
    if "motivo_perda_id" not in colunas_mov_validade:
        cursor.execute("ALTER TABLE movimentacoes_validade ADD COLUMN motivo_perda_id INTEGER")

    cursor.execute("PRAGMA table_info(itens_consumo_real_producao)")
    colunas_itens_consumo = [row["name"] for row in cursor.fetchall()]
    if "motivo_perda_id" not in colunas_itens_consumo:
        cursor.execute("ALTER TABLE itens_consumo_real_producao ADD COLUMN motivo_perda_id INTEGER")
    if "observacao_perda" not in colunas_itens_consumo:
        cursor.execute("ALTER TABLE itens_consumo_real_producao ADD COLUMN observacao_perda TEXT")

    conn.commit()
    conn.close()


def banco_vazio():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) AS total FROM sabores")
    total = cursor.fetchone()["total"]

    conn.close()
    return total == 0


def popular_banco_inicial():
    sabores_iniciais = [
        ("Bacalhau", "Prata", 15, 25, 35, 22),
        ("Banana c/ Canela", "Bronze", 5, 7, 8, 5),
        ("Calabresa", "Bronze", 8, 13, 15, 12),
        ("Camarão", "Prata", 23, 32, 45, 30),
        ("Charque", "Ouro", 75, 95, 114, 90),
        ("Chocobanana", "Bronze", 4, 6, 7, 6),
        ("Chocolate", "Ouro", 90, 115, 140, 105),
        ("Doce de Leite", "Bronze", 7, 10, 14, 12),
        ("Dois Amores", "Bronze", 15, 20, 20, 19),
        ("Frango", "Ouro", 70, 95, 114, 90),
        ("Goiabada", "Bronze", 6, 8, 10, 5),
        ("Palmito", "Bronze", 7, 9, 10, 7),
        ("Pizza", "Bronze", 6, 9, 10, 7),
        ("Queijo", "Prata", 25, 45, 50, 40),
        ("Queijo do Reino", "Prata", 25, 40, 45, 40),
        ("Tomate Seco", "Bronze", 5, 6, 8, 5),
    ]

    conn = conectar()
    cursor = conn.cursor()

    for nome, classe, chuva, normal, verao, baixa in sabores_iniciais:
        cursor.execute("""
            INSERT OR IGNORE INTO sabores
            (nome, classe, chuva, normal, verao, baixa)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nome, classe, chuva, normal, verao, baixa))

    conn.commit()

    cursor.execute("SELECT id, nome FROM sabores")
    sabores = {row["nome"]: row["id"] for row in cursor.fetchall()}

    metas_dia_iniciais = {
        "Segunda-feira": {
            "Frango": 20,
            "Charque": 20,
            "Chocolate": 20
        },
        "Quarta-feira": {
            "Frango": 20,
            "Charque": 20,
            "Chocolate": 20
        },
        "Quinta-feira": {
            "Queijo": 25,
            "Frango": 70,
            "Camarão": 23,
            "Charque": 75,
            "Queijo do Reino": 25,
            "Bacalhau": 15,
            "Chocolate": 90
        }
    }

    for dia_semana, metas in metas_dia_iniciais.items():
        for nome_sabor, meta in metas.items():
            sabor_id = sabores[nome_sabor]

            cursor.execute("""
                INSERT OR REPLACE INTO metas_dia
                (dia_semana, sabor_id, meta)
                VALUES (?, ?, ?)
            """, (dia_semana, sabor_id, meta))

    conn.commit()
    conn.close()


def popular_usuarios_iniciais():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) AS total FROM usuarios")
    total = cursor.fetchone()["total"]

    if total == 0:
        cursor.execute("""
            INSERT INTO usuarios (usuario, senha_hash, perfil)
            VALUES (?, ?, ?)
        """, (
            "admin",
            generate_password_hash("admin123"),
            "admin"
        ))

        cursor.execute("""
            INSERT INTO usuarios (usuario, senha_hash, perfil)
            VALUES (?, ?, ?)
        """, (
            "producao",
            generate_password_hash("prod123"),
            "producao"
        ))

    conn.commit()
    conn.close()



def popular_categorias_compra_iniciais():
    conn = conectar()
    cursor = conn.cursor()

    for nome_categoria in CATEGORIAS_COMPRA_INICIAIS:
        cursor.execute("""
            INSERT OR IGNORE INTO categorias_compra (nome, ativo)
            VALUES (?, 1)
        """, (nome_categoria,))

    conn.commit()
    conn.close()


def carregar_categorias_compra(apenas_ativas=True):
    conn = conectar()
    cursor = conn.cursor()

    if apenas_ativas:
        cursor.execute("""
            SELECT id, nome, ativo
            FROM categorias_compra
            WHERE ativo = 1
            ORDER BY nome
        """)
    else:
        cursor.execute("""
            SELECT id, nome, ativo
            FROM categorias_compra
            ORDER BY nome
        """)

    categorias = cursor.fetchall()
    conn.close()
    return categorias


def carregar_fornecedores(apenas_ativos=True):
    conn = conectar()
    cursor = conn.cursor()

    if apenas_ativos:
        cursor.execute("""
            SELECT id, nome, telefone, email, observacao, ativo
            FROM fornecedores
            WHERE ativo = 1
            ORDER BY nome
        """)
    else:
        cursor.execute("""
            SELECT id, nome, telefone, email, observacao, ativo
            FROM fornecedores
            ORDER BY nome
        """)

    fornecedores = cursor.fetchall()
    conn.close()
    return fornecedores


def criar_fornecedor_se_nao_existir(nome_fornecedor):
    nome_fornecedor = nome_fornecedor.strip()

    if nome_fornecedor == "":
        return None

    cursor.execute("""
        SELECT id
        FROM fornecedores
        WHERE nome = ?
    """, (nome_fornecedor,))

    fornecedor = cursor.fetchone()

    if fornecedor:
        fornecedor_id = fornecedor["id"]
    else:
        cursor.execute("""
            INSERT INTO fornecedores (nome, ativo)
            VALUES (?, 1)
        """, (nome_fornecedor,))
        fornecedor_id = cursor.lastrowid
        conn.commit()

    conn.close()
    return fornecedor_id


def contar_pedidos_pendentes():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM pedidos_compra
        WHERE status = 'Pendente'
    """)

    total = cursor.fetchone()["total"]
    conn.close()
    return total


def popular_categorias_estoque_iniciais():
    conn = conectar()
    cursor = conn.cursor()

    for nome_categoria in CATEGORIAS_ESTOQUE_INICIAIS:
        cursor.execute("""
            INSERT OR IGNORE INTO categorias_estoque (nome, ativo)
            VALUES (?, 1)
        """, (nome_categoria,))

    conn.commit()
    conn.close()


def carregar_categorias_estoque(apenas_ativas=True):
    conn = conectar()
    cursor = conn.cursor()

    if apenas_ativas:
        cursor.execute("""
            SELECT id, nome, ativo
            FROM categorias_estoque
            WHERE ativo = 1
            ORDER BY nome
        """)
    else:
        cursor.execute("""
            SELECT id, nome, ativo
            FROM categorias_estoque
            ORDER BY nome
        """)

    categorias = cursor.fetchall()
    conn.close()
    return categorias


def carregar_produtos_estoque(apenas_ativos=True):
    conn = conectar()
    cursor = conn.cursor()

    where = "WHERE produtos_estoque.ativo = 1" if apenas_ativos else ""

    cursor.execute(f"""
        SELECT
            produtos_estoque.*,
            categorias_estoque.nome AS categoria_nome
        FROM produtos_estoque
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        {where}
        ORDER BY CAST(produtos_estoque.codigo AS INTEGER), produtos_estoque.nome
    """)

    produtos = cursor.fetchall()
    conn.close()
    return produtos


def buscar_produto_estoque_cursor(cursor, produto_id):
    cursor.execute("""
        SELECT
            produtos_estoque.*,
            categorias_estoque.nome AS categoria_nome
        FROM produtos_estoque
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE produtos_estoque.id = ?
    """, (produto_id,))

    return cursor.fetchone()


def buscar_produto_estoque(produto_id):
    conn = conectar()
    cursor = conn.cursor()

    produto = buscar_produto_estoque_cursor(cursor, produto_id)
    conn.close()
    return produto


def buscar_produto_estoque_por_nome_cursor(cursor, nome):
    cursor.execute("""
        SELECT
            produtos_estoque.*,
            categorias_estoque.nome AS categoria_nome
        FROM produtos_estoque
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE produtos_estoque.nome = ?
    """, (nome,))

    produto = cursor.fetchone()
    return produto


def obter_proximo_codigo_produto(cursor):
    cursor.execute("""
        SELECT codigo
        FROM produtos_estoque
        WHERE codigo IS NOT NULL AND codigo != ''
        ORDER BY CAST(codigo AS INTEGER) DESC
        LIMIT 1
    """)

    ultimo = cursor.fetchone()

    if ultimo and str(ultimo["codigo"]).isdigit():
        proximo = int(ultimo["codigo"]) + 1
    else:
        proximo = 1

    return str(proximo).zfill(4)


def criar_produto_estoque(nome, categoria_id, unidade_padrao, estoque_minimo=0, observacao="", custo_padrao=0, ativo_venda=0, controla_validade=0, dias_validade=None):
    nome = nome.strip()
    unidade_padrao = unidade_padrao.strip()
    observacao = observacao.strip()

    if nome == "" or unidade_padrao == "":
        return None

    try:
        estoque_minimo = float(str(estoque_minimo or 0).replace(",", "."))
    except:
        estoque_minimo = 0

    try:
        custo_padrao = float(str(custo_padrao or 0).replace(",", "."))
    except:
        custo_padrao = 0

    ativo_venda = 1 if str(ativo_venda) == "1" or ativo_venda is True else 0

    conn = conectar()
    cursor = conn.cursor()

    codigo = obter_proximo_codigo_produto(cursor)

    cursor.execute("""
        INSERT INTO produtos_estoque
        (codigo, nome, categoria_id, unidade_padrao, estoque_minimo, custo_padrao, ativo, ativo_venda, controla_validade, dias_validade, data_cadastro, observacao)
        VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?, ?, ?, ?)
    """, (
        codigo,
        nome,
        int(categoria_id) if categoria_id else None,
        unidade_padrao,
        estoque_minimo,
        custo_padrao,
        ativo_venda,
        1 if str(controla_validade) == "1" or controla_validade is True else 0,
        int(dias_validade) if str(dias_validade or "").strip().isdigit() else None,
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        observacao
    ))

    produto_id = cursor.lastrowid
    conn.commit()
    conn.close()

    garantir_embalagem_unidade_produto(produto_id, unidade_padrao)

    if nome.lower().startswith("empada de "):
        conn = conectar()
        cursor = conn.cursor()
        for nome_embalagem, fator in EMBALAGENS_PADRAO_EMPADAS:
            cursor.execute("""
                INSERT OR IGNORE INTO produto_embalagens
                (produto_id, nome, fator_conversao, padrao, ativo, observacao)
                VALUES (?, ?, ?, ?, 1, ?)
            """, (
                produto_id,
                nome_embalagem,
                fator,
                1 if fator == 1 else 0,
                "Embalagem padrão de empadas criada automaticamente."
            ))
        conn.commit()
        conn.close()

    return produto_id


def popular_produtos_estoque_iniciais():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id
        FROM categorias_estoque
        WHERE nome = 'Produção Própria'
    """)

    categoria = cursor.fetchone()
    categoria_id = categoria["id"] if categoria else None

    cursor.execute("""
        SELECT nome
        FROM sabores
        ORDER BY id
    """)

    sabores = cursor.fetchall()

    for sabor in sabores:
        nome_produto = f"Empada de {sabor['nome']}"

        cursor.execute("""
            SELECT id
            FROM produtos_estoque
            WHERE nome = ?
        """, (nome_produto,))

        if cursor.fetchone():
            continue

        codigo = obter_proximo_codigo_produto(cursor)

        cursor.execute("""
            INSERT INTO produtos_estoque
            (codigo, nome, categoria_id, unidade_padrao, estoque_minimo, custo_padrao, ativo, ativo_venda, controla_validade, dias_validade, data_cadastro, observacao)
            VALUES (?, ?, ?, 'un', 0, 0, 1, 1, 1, ?, ?, ?)
        """, (
            codigo,
            nome_produto,
            categoria_id,
            obter_dias_validade(sabor["nome"]),
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Produto inicial da casa criado automaticamente."
        ))

    conn.commit()
    conn.close()

def criar_embalagem_produto(produto_id, nome, fator_conversao, padrao=0, observacao=""):
    nome = nome.strip()
    observacao = observacao.strip()

    try:
        fator_conversao = float(str(fator_conversao or 1).replace(",", "."))
    except:
        fator_conversao = 1

    if nome == "" or fator_conversao <= 0:
        return None

    conn = conectar()
    cursor = conn.cursor()

    if int(padrao) == 1:
        cursor.execute("""
            UPDATE produto_embalagens
            SET padrao = 0
            WHERE produto_id = ?
        """, (produto_id,))

    cursor.execute("""
        INSERT OR IGNORE INTO produto_embalagens
        (produto_id, nome, fator_conversao, padrao, ativo, observacao)
        VALUES (?, ?, ?, ?, 1, ?)
    """, (produto_id, nome, fator_conversao, int(padrao), observacao))

    embalagem_id = cursor.lastrowid
    conn.commit()
    conn.close()

    return embalagem_id


def carregar_embalagens_produto(produto_id, apenas_ativas=True):
    conn = conectar()
    cursor = conn.cursor()

    if apenas_ativas:
        cursor.execute("""
            SELECT id, produto_id, nome, fator_conversao, padrao, ativo, observacao, disponivel_loja
            FROM produto_embalagens
            WHERE produto_id = ? AND ativo = 1
            ORDER BY padrao DESC, fator_conversao, nome
        """, (produto_id,))
    else:
        cursor.execute("""
            SELECT id, produto_id, nome, fator_conversao, padrao, ativo, observacao, disponivel_loja
            FROM produto_embalagens
            WHERE produto_id = ?
            ORDER BY padrao DESC, fator_conversao, nome
        """, (produto_id,))

    embalagens = cursor.fetchall()
    conn.close()

    return embalagens


def carregar_embalagens_todos_produtos():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, produto_id, nome, fator_conversao, padrao, ativo, observacao, disponivel_loja
        FROM produto_embalagens
        WHERE ativo = 1
        ORDER BY padrao DESC, fator_conversao, nome
    """)

    mapa = {}

    for row in cursor.fetchall():
        produto_id = str(row["produto_id"])
        mapa.setdefault(produto_id, []).append({
            "id": row["id"],
            "nome": row["nome"],
            "fator_conversao": row["fator_conversao"],
            "padrao": row["padrao"],
        })

    conn.close()
    return mapa


def buscar_embalagem_produto_cursor(cursor, embalagem_id, produto_id=None):
    if produto_id:
        cursor.execute("""
            SELECT id, produto_id, nome, fator_conversao, padrao, ativo, observacao, disponivel_loja
            FROM produto_embalagens
            WHERE id = ? AND produto_id = ? AND ativo = 1
        """, (embalagem_id, produto_id))
    else:
        cursor.execute("""
            SELECT id, produto_id, nome, fator_conversao, padrao, ativo, observacao, disponivel_loja
            FROM produto_embalagens
            WHERE id = ? AND ativo = 1
        """, (embalagem_id,))

    return cursor.fetchone()


def buscar_embalagem_produto(embalagem_id, produto_id=None):
    conn = conectar()
    cursor = conn.cursor()

    embalagem = buscar_embalagem_produto_cursor(cursor, embalagem_id, produto_id)
    conn.close()
    return embalagem


def garantir_embalagem_unidade_produto(produto_id, unidade_padrao):
    nome_embalagem = f"Unidade base ({unidade_padrao})" if unidade_padrao != "un" else "Unidade"

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id
        FROM produto_embalagens
        WHERE produto_id = ? AND fator_conversao = 1
        LIMIT 1
    """, (produto_id,))

    if not cursor.fetchone():
        cursor.execute("""
            INSERT OR IGNORE INTO produto_embalagens
            (produto_id, nome, fator_conversao, padrao, ativo, observacao)
            VALUES (?, ?, 1, 1, 1, ?)
        """, (produto_id, nome_embalagem, "Embalagem base criada automaticamente."))

    conn.commit()
    conn.close()


def popular_embalagens_padrao_empadas():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, nome
        FROM produtos_estoque
        WHERE nome LIKE 'Empada de %'
    """)

    produtos = cursor.fetchall()

    for produto in produtos:
        for nome, fator in EMBALAGENS_PADRAO_EMPADAS:
            cursor.execute("""
                INSERT OR IGNORE INTO produto_embalagens
                (produto_id, nome, fator_conversao, padrao, ativo, observacao)
                VALUES (?, ?, ?, ?, 1, ?)
            """, (
                produto["id"],
                nome,
                fator,
                1 if fator == 1 else 0,
                "Embalagem padrão de empadas criada automaticamente."
            ))

    conn.commit()
    conn.close()


def obter_saldo_produto_cursor(cursor, produto_id):
    cursor.execute("""
        SELECT
            COALESCE(SUM(
                CASE
                    WHEN tipo = 'Entrada' THEN quantidade
                    WHEN tipo = 'Saída' THEN -quantidade
                    ELSE 0
                END
            ), 0) AS saldo_atual
        FROM movimentacoes_estoque
        WHERE produto_id = ?
    """, (produto_id,))

    saldo = cursor.fetchone()["saldo_atual"]
    return float(saldo or 0)


def obter_saldo_produto(produto_id):
    conn = conectar()
    cursor = conn.cursor()

    saldo = obter_saldo_produto_cursor(cursor, produto_id)
    conn.close()
    return saldo


def carregar_ultimas_movimentacoes_produto(produto_id, limite=8):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            movimentacoes_estoque.id,
            movimentacoes_estoque.tipo,
            movimentacoes_estoque.quantidade,
            movimentacoes_estoque.quantidade_embalagem,
            movimentacoes_estoque.embalagem_nome,
            movimentacoes_estoque.fator_conversao,
            movimentacoes_estoque.data_movimentacao,
            movimentacoes_estoque.origem,
            movimentacoes_estoque.usuario,
            movimentacoes_estoque.observacao
        FROM movimentacoes_estoque
        WHERE movimentacoes_estoque.produto_id = ?
        ORDER BY movimentacoes_estoque.id DESC
        LIMIT ?
    """, (produto_id, limite))

    movimentacoes = cursor.fetchall()
    conn.close()
    return movimentacoes


def montar_resumo_produto_estoque(produto):
    saldo_atual = obter_saldo_produto(produto["id"])
    status_saldo = montar_status_saldo(saldo_atual, float(produto["estoque_minimo"] or 0))

    return {
        "saldo_atual": saldo_atual,
        "status_saldo": status_saldo,
        "unidade": produto["unidade_padrao"]
    }



def buscar_produto_estoque_por_nome(nome):
    conn = conectar()
    cursor = conn.cursor()

    produto = buscar_produto_estoque_por_nome_cursor(cursor, nome)
    conn.close()
    return produto


def buscar_embalagem_produto_por_nome_cursor(cursor, produto_id, nome_embalagem):
    cursor.execute("""
        SELECT id, produto_id, nome, fator_conversao, padrao, ativo, observacao, disponivel_loja
        FROM produto_embalagens
        WHERE produto_id = ? AND nome = ? AND ativo = 1
        LIMIT 1
    """, (produto_id, nome_embalagem))

    return cursor.fetchone()


def buscar_embalagem_produto_por_nome(produto_id, nome_embalagem):
    conn = conectar()
    cursor = conn.cursor()

    embalagem = buscar_embalagem_produto_por_nome_cursor(cursor, produto_id, nome_embalagem)
    conn.close()
    return embalagem


def _criar_pre_producao_realizada_cursor(cursor, resultado, cenario, dia_semana, total_tabuleiros, total_empadas, origem="Etiquetas", planejamento_id=None):
    itens_validos = [item for item in resultado if float(item.get("empadas", 0) or 0) > 0]

    if not itens_validos:
        return None

    cursor.execute("""
        INSERT INTO producoes_realizadas
        (data_criacao, cenario, dia_semana, total_tabuleiros, total_unidades, status, origem, criado_por, observacao, planejamento_id)
        VALUES (?, ?, ?, ?, ?, 'Pendente', ?, ?, ?, ?)
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        cenario,
        dia_semana,
        float(total_tabuleiros or 0),
        float(total_empadas or 0),
        origem,
        session.get("usuario", "sistema") if has_request_context() else "sistema",
        (
            "Pré-produção criada automaticamente a partir do planejamento de produção. "
            "Conferir antes de confirmar entrada no estoque."
            if planejamento_id else
            "Pré-produção criada automaticamente. Conferir antes de confirmar entrada no estoque."
        ),
        planejamento_id
    ))

    producao_realizada_id = cursor.lastrowid

    for item in itens_validos:
        sabor = item["nome"]
        quantidade_tabuleiros = float(item["producao"] or 0)
        quantidade_unidades = float(item.get("empadas", quantidade_tabuleiros * EMPADAS_POR_TABULEIRO) or 0)
        nome_produto = f"Empada de {sabor}"

        produto = buscar_produto_estoque_por_nome_cursor(cursor, nome_produto)
        produto_id = produto["id"] if produto else None
        embalagem_id = None
        possui_misto = float(item.get("misto_unidades", 0) or 0) > 0
        embalagem_nome = "Unidade" if possui_misto else "Tabuleiro c/35"
        fator_conversao = 1 if possui_misto else EMPADAS_POR_TABULEIRO
        observacao_item = "Pré-lançado a partir da produção calculada. Consumo de recheio/ficha técnica ficará para etapa futura."

        if produto_id:
            embalagem = buscar_embalagem_produto_por_nome_cursor(cursor, produto_id, embalagem_nome)
            if embalagem:
                embalagem_id = embalagem["id"]
                embalagem_nome = embalagem["nome"]
                fator_conversao = float(embalagem["fator_conversao"] or EMPADAS_POR_TABULEIRO)
            else:
                observacao_item += f" Embalagem {embalagem_nome} não encontrada; confirmação usará unidade base."
        else:
            observacao_item += f" Produto oficial não encontrado: {nome_produto}."

        cursor.execute("""
            INSERT INTO itens_producao_realizada
            (producao_realizada_id, sabor, produto_estoque_id, quantidade_tabuleiros, quantidade_unidades, embalagem_id, embalagem_nome, fator_conversao, status, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', ?)
        """, (
            producao_realizada_id,
            sabor,
            produto_id,
            quantidade_tabuleiros,
            quantidade_unidades,
            embalagem_id,
            embalagem_nome,
            fator_conversao,
            observacao_item
        ))

    return producao_realizada_id


def criar_pre_producao_realizada(resultado, cenario, dia_semana, total_tabuleiros, total_empadas, origem="Etiquetas", planejamento_id=None, cursor=None):
    if cursor is not None:
        return _criar_pre_producao_realizada_cursor(
            cursor,
            resultado,
            cenario,
            dia_semana,
            total_tabuleiros,
            total_empadas,
            origem=origem,
            planejamento_id=planejamento_id
        )

    conn = conectar()
    cursor_local = conn.cursor()
    try:
        producao_realizada_id = _criar_pre_producao_realizada_cursor(
            cursor_local,
            resultado,
            cenario,
            dia_semana,
            total_tabuleiros,
            total_empadas,
            origem=origem,
            planejamento_id=planejamento_id
        )
        conn.commit()
        return producao_realizada_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def contar_pre_producoes_pendentes():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM producoes_realizadas
        WHERE status = 'Pendente'
    """)

    total = cursor.fetchone()["total"]
    conn.close()
    return total


def carregar_producoes_realizadas(status=""):
    conn = conectar()
    cursor = conn.cursor()

    if status:
        cursor.execute("""
            SELECT *
            FROM producoes_realizadas
            WHERE status = ?
            ORDER BY id DESC
        """, (status,))
    else:
        cursor.execute("""
            SELECT *
            FROM producoes_realizadas
            ORDER BY id DESC
        """)

    producoes = cursor.fetchall()
    conn.close()
    return producoes


def buscar_producao_realizada(producao_realizada_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM producoes_realizadas
        WHERE id = ?
    """, (producao_realizada_id,))

    producao = cursor.fetchone()
    conn.close()
    return producao


def carregar_itens_producao_realizada(producao_realizada_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            itens_producao_realizada.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao
        FROM itens_producao_realizada
        LEFT JOIN produtos_estoque ON produtos_estoque.id = itens_producao_realizada.produto_estoque_id
        WHERE itens_producao_realizada.producao_realizada_id = ?
        ORDER BY itens_producao_realizada.id
    """, (producao_realizada_id,))

    itens = cursor.fetchall()
    conn.close()
    return itens


def _confirmar_producao_realizada_estoque_legacy_v351(producao_realizada_id, observacao_confirmacao=""):
    producao = buscar_producao_realizada(producao_realizada_id)

    if not producao:
        raise ValueError("Pré-produção não encontrada.")

    if producao["status"] != "Pendente":
        raise ValueError("Essa pré-produção já foi finalizada ou cancelada.")

    itens = carregar_itens_producao_realizada(producao_realizada_id)
    itens_pendentes = [item for item in itens if item["status"] == "Pendente"]

    if not itens_pendentes:
        raise ValueError("Essa pré-produção não possui itens pendentes para confirmar.")

    for item in itens_pendentes:
        if not item["produto_estoque_id"]:
            raise ValueError(f"O sabor {item['sabor']} não está vinculado a um produto oficial do estoque.")

        if float(item["quantidade_unidades"] or 0) <= 0:
            raise ValueError(f"O item {item['sabor']} está com quantidade zerada ou inválida.")

    itens_movimentados = 0

    for item in itens_pendentes:
        if item["estoque_movimentacao_id"]:
            continue

        observacao = (
            f"Entrada automática por produção realizada #{producao_realizada_id}. "
            f"Sabor: {item['sabor']}. "
            f"Quantidade: {item['quantidade_tabuleiros']:.2f} tabuleiro(s), equivalente a {item['quantidade_unidades']:.2f} {item['unidade_padrao']}. "
            f"Ficha técnica/consumo de recheio será tratado em etapa futura."
        )

        if observacao_confirmacao:
            observacao = f"{observacao} Observação da confirmação: {observacao_confirmacao}"

        if item["embalagem_id"]:
            movimentacao_id = registrar_movimentacao_estoque(
                produto_id=item["produto_estoque_id"],
                tipo="Entrada",
                quantidade=item["quantidade_tabuleiros"],
                data_movimentacao=datetime.now().strftime("%Y-%m-%d"),
                origem="Produção Interna",
                embalagem_id=item["embalagem_id"],
                origem_id=producao_realizada_id,
                observacao=observacao
            )
        else:
            movimentacao_id = registrar_movimentacao_estoque(
                produto_id=item["produto_estoque_id"],
                tipo="Entrada",
                quantidade=item["quantidade_unidades"],
                data_movimentacao=datetime.now().strftime("%Y-%m-%d"),
                origem="Produção Interna",
                origem_id=producao_realizada_id,
                observacao=observacao
            )

        conn_item = conectar()
        cursor_item = conn_item.cursor()
        cursor_item.execute("""
            UPDATE itens_producao_realizada
            SET status = 'Confirmado', estoque_movimentacao_id = ?
            WHERE id = ?
        """, (movimentacao_id, item["id"]))
        conn_item.commit()
        conn_item.close()

        # A entrada do produto acabado já foi registrada acima.
        # Se o produto controla validade, cria apenas o lote vinculado, sem duplicar o saldo.
        criar_lote_validade_item_producao(
            item,
            producao,
            datetime.now().strftime("%Y-%m-%d")
        )
        itens_movimentados += 1

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE producoes_realizadas
        SET status = 'Confirmado',
            data_confirmacao = ?,
            confirmado_por = ?,
            observacao = CASE
                WHEN observacao IS NULL OR observacao = '' THEN ?
                ELSE observacao || ' | ' || ?
            END
        WHERE id = ?
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        session.get("usuario") if has_request_context() else "sistema",
        observacao_confirmacao or f"Produção confirmada. Itens movimentados: {itens_movimentados}.",
        observacao_confirmacao or f"Produção confirmada. Itens movimentados: {itens_movimentados}.",
        producao_realizada_id
    ))

    if producao["planejamento_id"]:
        cursor.execute("""
            UPDATE planejamentos_producao
            SET status = 'Produção confirmada',
                observacao = CASE
                    WHEN observacao IS NULL OR observacao = '' THEN ?
                    ELSE observacao || ' | ' || ?
                END
            WHERE id = ?
        """, (
            f"Produção realizada #{producao_realizada_id} confirmada em {_agora_texto()}.",
            f"Produção realizada #{producao_realizada_id} confirmada em {_agora_texto()}.",
            producao["planejamento_id"]
        ))

    conn.commit()
    conn.close()


def confirmar_producao_realizada_estoque(producao_realizada_id, observacao_confirmacao=""):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")

        cursor.execute("""
            SELECT *
            FROM producoes_realizadas
            WHERE id = ?
        """, (producao_realizada_id,))
        producao = cursor.fetchone()

        if not producao:
            raise ValueError("Pre-producao nao encontrada.")

        if producao["status"] != "Pendente":
            raise ValueError("Essa pre-producao ja foi finalizada ou cancelada.")

        cursor.execute("""
            SELECT
                itens_producao_realizada.*,
                produtos_estoque.codigo AS produto_codigo,
                produtos_estoque.nome AS produto_nome,
                produtos_estoque.unidade_padrao
            FROM itens_producao_realizada
            LEFT JOIN produtos_estoque ON produtos_estoque.id = itens_producao_realizada.produto_estoque_id
            WHERE itens_producao_realizada.producao_realizada_id = ?
            ORDER BY itens_producao_realizada.id
        """, (producao_realizada_id,))
        itens = cursor.fetchall()
        itens_pendentes = [item for item in itens if item["status"] == "Pendente"]

        if not itens_pendentes:
            raise ValueError("Essa pre-producao nao possui itens pendentes para confirmar.")

        for item in itens_pendentes:
            if not item["produto_estoque_id"]:
                raise ValueError(f"O sabor {item['sabor']} nao esta vinculado a um produto oficial do estoque.")

            if float(item["quantidade_unidades"] or 0) <= 0:
                raise ValueError(f"O item {item['sabor']} esta com quantidade zerada ou invalida.")

            produto = buscar_produto_estoque_cursor(cursor, item["produto_estoque_id"])
            if not produto:
                raise ValueError(f"Produto oficial nao encontrado para o sabor {item['sabor']}.")
            if int(produto["controla_validade"] or 0) == 1 and int(produto["dias_validade"] or 0) <= 0:
                raise ValueError(
                    f"O produto {produto['nome']} controla validade, mas nao possui dias de validade configurados."
                )

        itens_movimentados = 0
        data_movimentacao = datetime.now().strftime("%Y-%m-%d")

        for item in itens_pendentes:
            if item["estoque_movimentacao_id"]:
                continue

            quantidade_tabuleiros = float(item["quantidade_tabuleiros"] or 0)
            quantidade_unidades = float(item["quantidade_unidades"] or 0)
            fator_conversao = float(item["fator_conversao"] or 1)
            embalagem_id_movimento = None
            quantidade_movimento = quantidade_unidades

            if item["embalagem_id"] and quantidade_tabuleiros > 0:
                unidades_convertidas = quantidade_tabuleiros * fator_conversao
                if abs(unidades_convertidas - quantidade_unidades) <= 0.0001:
                    embalagem_id_movimento = item["embalagem_id"]
                    quantidade_movimento = quantidade_tabuleiros

            observacao = (
                f"Entrada automatica por producao realizada #{producao_realizada_id}. "
                f"Sabor: {item['sabor']}. "
                f"Quantidade: {quantidade_tabuleiros:.2f} tabuleiro(s), equivalente a {quantidade_unidades:.2f} {item['unidade_padrao']}. "
                f"Ficha tecnica/consumo de recheio sera tratado em etapa futura."
            )

            if observacao_confirmacao:
                observacao = f"{observacao} Observacao da confirmacao: {observacao_confirmacao}"

            movimentacao_id = registrar_movimentacao_estoque_cursor(
                cursor,
                produto_id=item["produto_estoque_id"],
                tipo="Entrada",
                quantidade=quantidade_movimento,
                data_movimentacao=data_movimentacao,
                origem="Produção Interna",
                embalagem_id=embalagem_id_movimento,
                origem_id=producao_realizada_id,
                observacao=observacao
            )

            cursor.execute("""
                UPDATE itens_producao_realizada
                SET status = 'Confirmado', estoque_movimentacao_id = ?
                WHERE id = ?
            """, (movimentacao_id, item["id"]))

            criar_lote_validade_item_producao(
                item,
                producao,
                data_movimentacao,
                cursor=cursor
            )
            itens_movimentados += 1

        observacao_final = observacao_confirmacao or f"Producao confirmada. Itens movimentados: {itens_movimentados}."
        usuario = session.get("usuario") if has_request_context() else "sistema"

        cursor.execute("""
            UPDATE producoes_realizadas
            SET status = 'Confirmado',
                data_confirmacao = ?,
                confirmado_por = ?,
                observacao = CASE
                    WHEN observacao IS NULL OR observacao = '' THEN ?
                    ELSE observacao || ' | ' || ?
                END
            WHERE id = ?
        """, (
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            usuario,
            observacao_final,
            observacao_final,
            producao_realizada_id
        ))

        if producao["planejamento_id"]:
            obs_planejamento = f"Producao realizada #{producao_realizada_id} confirmada em {_agora_texto()}."
            cursor.execute("""
                UPDATE planejamentos_producao
                SET status = 'Produção confirmada',
                    observacao = CASE
                        WHEN observacao IS NULL OR observacao = '' THEN ?
                        ELSE observacao || ' | ' || ?
                    END
                WHERE id = ?
            """, (
                obs_planejamento,
                obs_planejamento,
                producao["planejamento_id"]
            ))

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def recalcular_totais_producao_realizada(producao_realizada_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            COALESCE(SUM(quantidade_tabuleiros), 0) AS total_tabuleiros,
            COALESCE(SUM(quantidade_unidades), 0) AS total_unidades
        FROM itens_producao_realizada
        WHERE producao_realizada_id = ?
        AND status != 'Cancelado'
    """, (producao_realizada_id,))

    totais = cursor.fetchone()

    cursor.execute("""
        UPDATE producoes_realizadas
        SET total_tabuleiros = ?,
            total_unidades = ?
        WHERE id = ?
    """, (
        totais["total_tabuleiros"] or 0,
        totais["total_unidades"] or 0,
        producao_realizada_id
    ))

    conn.commit()
    conn.close()


def atualizar_item_pre_producao(item_id, produto_estoque_id, embalagem_id, quantidade_embalagem, observacao=""):
    try:
        produto_estoque_id = int(produto_estoque_id)
    except:
        raise ValueError("Produto oficial inválido.")

    try:
        embalagem_id = int(embalagem_id)
    except:
        embalagem_id = None

    try:
        quantidade_embalagem = float(str(quantidade_embalagem or 0).replace(",", "."))
    except:
        quantidade_embalagem = 0

    if quantidade_embalagem <= 0:
        raise ValueError("A quantidade precisa ser maior que zero.")

    produto = buscar_produto_estoque(produto_estoque_id)
    if not produto:
        raise ValueError("Produto oficial não encontrado.")

    embalagem = buscar_embalagem_produto(embalagem_id, produto_estoque_id) if embalagem_id else None

    if embalagem:
        embalagem_nome = embalagem["nome"]
        fator_conversao = float(embalagem["fator_conversao"] or 1)
    else:
        embalagem_nome = f"Unidade base ({produto['unidade_padrao']})"
        fator_conversao = 1
        embalagem_id = None

    quantidade_unidades = quantidade_embalagem * fator_conversao

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT producao_realizada_id, status, estoque_movimentacao_id
        FROM itens_producao_realizada
        WHERE id = ?
    """, (item_id,))

    item = cursor.fetchone()

    if not item:
        conn.close()
        raise ValueError("Item da pré-produção não encontrado.")

    if item["status"] != "Pendente" or item["estoque_movimentacao_id"]:
        conn.close()
        raise ValueError("Somente itens pendentes e ainda não movimentados podem ser alterados.")

    cursor.execute("""
        UPDATE itens_producao_realizada
        SET produto_estoque_id = ?,
            quantidade_tabuleiros = ?,
            quantidade_unidades = ?,
            embalagem_id = ?,
            embalagem_nome = ?,
            fator_conversao = ?,
            observacao = ?
        WHERE id = ?
    """, (
        produto_estoque_id,
        quantidade_embalagem,
        quantidade_unidades,
        embalagem_id,
        embalagem_nome,
        fator_conversao,
        observacao.strip(),
        item_id
    ))

    producao_realizada_id = item["producao_realizada_id"]

    conn.commit()
    conn.close()

    recalcular_totais_producao_realizada(producao_realizada_id)
    return producao_realizada_id


def cancelar_item_pre_producao(item_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT producao_realizada_id, status, estoque_movimentacao_id
        FROM itens_producao_realizada
        WHERE id = ?
    """, (item_id,))

    item = cursor.fetchone()

    if not item:
        conn.close()
        raise ValueError("Item da pré-produção não encontrado.")

    if item["status"] != "Pendente" or item["estoque_movimentacao_id"]:
        conn.close()
        raise ValueError("Somente itens pendentes e ainda não movimentados podem ser cancelados.")

    cursor.execute("""
        UPDATE itens_producao_realizada
        SET status = 'Cancelado'
        WHERE id = ?
    """, (item_id,))

    producao_realizada_id = item["producao_realizada_id"]

    conn.commit()
    conn.close()

    recalcular_totais_producao_realizada(producao_realizada_id)
    return producao_realizada_id


def adicionar_item_pre_producao(producao_realizada_id, produto_estoque_id, embalagem_id, quantidade_embalagem, sabor="", observacao=""):
    producao = buscar_producao_realizada(producao_realizada_id)

    if not producao:
        raise ValueError("Pré-produção não encontrada.")

    if producao["status"] != "Pendente":
        raise ValueError("Somente pré-produções pendentes podem receber novos itens.")

    try:
        produto_estoque_id = int(produto_estoque_id)
    except:
        raise ValueError("Produto oficial inválido.")

    try:
        embalagem_id = int(embalagem_id)
    except:
        embalagem_id = None

    try:
        quantidade_embalagem = float(str(quantidade_embalagem or 0).replace(",", "."))
    except:
        quantidade_embalagem = 0

    if quantidade_embalagem <= 0:
        raise ValueError("A quantidade precisa ser maior que zero.")

    produto = buscar_produto_estoque(produto_estoque_id)
    if not produto:
        raise ValueError("Produto oficial não encontrado.")

    embalagem = buscar_embalagem_produto(embalagem_id, produto_estoque_id) if embalagem_id else None

    if embalagem:
        embalagem_nome = embalagem["nome"]
        fator_conversao = float(embalagem["fator_conversao"] or 1)
    else:
        embalagem_nome = f"Unidade base ({produto['unidade_padrao']})"
        fator_conversao = 1
        embalagem_id = None

    quantidade_unidades = quantidade_embalagem * fator_conversao
    sabor = sabor.strip() or produto["nome"].replace("Empada de ", "")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO itens_producao_realizada
        (producao_realizada_id, sabor, produto_estoque_id, quantidade_tabuleiros, quantidade_unidades, embalagem_id, embalagem_nome, fator_conversao, status, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', ?)
    """, (
        producao_realizada_id,
        sabor,
        produto_estoque_id,
        quantidade_embalagem,
        quantidade_unidades,
        embalagem_id,
        embalagem_nome,
        fator_conversao,
        observacao.strip()
    ))

    conn.commit()
    conn.close()

    recalcular_totais_producao_realizada(producao_realizada_id)
    return producao_realizada_id


def registrar_movimentacao_estoque_cursor(
    cursor,
    produto_id,
    tipo,
    quantidade,
    data_movimentacao,
    origem,
    embalagem_id=None,
    origem_id=None,
    item_compra_id=None,
    nota_entrada_item_id=None,
    fornecedor_id=None,
    observacao="",
    permitir_saldo_negativo=False,
    justificativa_saldo_negativo=""
):
    try:
        quantidade = float(str(quantidade).replace(",", "."))
    except:
        quantidade = 0

    if quantidade <= 0:
        raise ValueError("Quantidade inválida para movimentação de estoque.")

    if tipo not in TIPOS_MOVIMENTACAO_ESTOQUE:
        raise ValueError("Tipo de movimentação inválido.")

    produto = buscar_produto_estoque_cursor(cursor, produto_id)

    if not produto:
        raise ValueError("Produto de estoque não encontrado.")

    quantidade_embalagem = quantidade
    embalagem_nome = produto["unidade_padrao"]
    fator_conversao = 1

    if embalagem_id:
        embalagem = buscar_embalagem_produto_cursor(cursor, embalagem_id, produto_id)

        if not embalagem:
            raise ValueError("Embalagem inválida para o produto selecionado.")

        embalagem_nome = embalagem["nome"]
        fator_conversao = float(embalagem["fator_conversao"] or 1)
        quantidade = quantidade_embalagem * fator_conversao

    if tipo == "Saída":
        saldo_atual = obter_saldo_produto_cursor(cursor, produto_id)

        if quantidade > saldo_atual:
            if not permitir_saldo_negativo:
                raise ValueError(
                    f"Saldo insuficiente. Saldo atual de {produto['nome']}: {saldo_atual:.2f} {produto['unidade_padrao']}."
                )

            justificativa_saldo_negativo = justificativa_saldo_negativo.strip()

            perfil_usuario = session.get("perfil") if has_request_context() else None
            if perfil_usuario != "admin":
                raise ValueError("Somente admin pode autorizar saída com saldo negativo.")

            if justificativa_saldo_negativo == "":
                raise ValueError("Informe uma justificativa para autorizar saída com saldo negativo.")

            observacao_extra = (
                f"SAÍDA COM SALDO INSUFICIENTE AUTORIZADA. "
                f"Saldo antes: {saldo_atual:.2f} {produto['unidade_padrao']}. "
                f"Quantidade retirada: {quantidade:.2f} {produto['unidade_padrao']}. "
                f"Justificativa: {justificativa_saldo_negativo}"
            )

            if observacao:
                observacao = f"{observacao} | {observacao_extra}"
            else:
                observacao = observacao_extra

    if data_movimentacao == "":
        data_movimentacao = datetime.now().strftime("%d/%m/%Y")
    else:
        try:
            data_movimentacao = datetime.strptime(data_movimentacao, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            pass

    cursor.execute("""
        INSERT INTO movimentacoes_estoque
        (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_id, embalagem_nome, fator_conversao, data_movimentacao, origem, origem_id, item_compra_id, nota_entrada_item_id, fornecedor_id, usuario, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        produto_id,
        tipo,
        quantidade,
        quantidade_embalagem,
        embalagem_id,
        embalagem_nome,
        fator_conversao,
        data_movimentacao,
        origem,
        origem_id,
        item_compra_id,
        nota_entrada_item_id,
        fornecedor_id,
        session.get("usuario") if has_request_context() else "sistema",
        observacao
    ))

    movimentacao_id = cursor.lastrowid
    return movimentacao_id


def registrar_movimentacao_estoque(
    produto_id,
    tipo,
    quantidade,
    data_movimentacao,
    origem,
    embalagem_id=None,
    origem_id=None,
    item_compra_id=None,
    nota_entrada_item_id=None,
    fornecedor_id=None,
    observacao="",
    permitir_saldo_negativo=False,
    justificativa_saldo_negativo=""
):
    conn = conectar()
    cursor = conn.cursor()
    try:
        movimentacao_id = registrar_movimentacao_estoque_cursor(
            cursor,
            produto_id=produto_id,
            tipo=tipo,
            quantidade=quantidade,
            data_movimentacao=data_movimentacao,
            origem=origem,
            embalagem_id=embalagem_id,
            origem_id=origem_id,
            item_compra_id=item_compra_id,
            nota_entrada_item_id=nota_entrada_item_id,
            fornecedor_id=fornecedor_id,
            observacao=observacao,
            permitir_saldo_negativo=permitir_saldo_negativo,
            justificativa_saldo_negativo=justificativa_saldo_negativo
        )
        conn.commit()
        return movimentacao_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def montar_status_saldo(saldo_atual, estoque_minimo):
    if saldo_atual < 0:
        return "Negativo"

    if saldo_atual == 0:
        return "Zerado"

    if estoque_minimo > 0 and saldo_atual <= estoque_minimo:
        return "Baixo"

    return "OK"


def carregar_saldo_estoque(busca="", categoria_id="", status=""):
    conn = conectar()
    cursor = conn.cursor()

    filtros = ["produtos_estoque.ativo = 1"]
    parametros = []

    busca = str(busca or "").strip()
    categoria_id = str(categoria_id or "").strip()
    status = str(status or "").strip()

    if busca:
        filtros.append("""
            (produtos_estoque.nome LIKE ?
            OR produtos_estoque.codigo LIKE ?
            OR categorias_estoque.nome LIKE ?)
        """)
        termo = f"%{busca}%"
        parametros.extend([termo, termo, termo])

    if categoria_id:
        filtros.append("produtos_estoque.categoria_id = ?")
        parametros.append(categoria_id)

    where_sql = " AND ".join(filtros)

    cursor.execute(f"""
        SELECT
            produtos_estoque.id,
            produtos_estoque.codigo,
            produtos_estoque.nome,
            produtos_estoque.unidade_padrao,
            produtos_estoque.estoque_minimo,
            produtos_estoque.custo_padrao,
            produtos_estoque.ativo,
            produtos_estoque.ativo_venda,
            categorias_estoque.nome AS categoria_nome,
            categorias_estoque.id AS categoria_id,
            COALESCE(SUM(
                CASE
                    WHEN movimentacoes_estoque.tipo = 'Entrada' THEN movimentacoes_estoque.quantidade
                    WHEN movimentacoes_estoque.tipo = 'Saída' THEN -movimentacoes_estoque.quantidade
                    ELSE 0
                END
            ), 0) AS saldo_atual
        FROM produtos_estoque
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        LEFT JOIN movimentacoes_estoque ON movimentacoes_estoque.produto_id = produtos_estoque.id
        WHERE {where_sql}
        GROUP BY produtos_estoque.id
        ORDER BY categorias_estoque.nome, CAST(produtos_estoque.codigo AS INTEGER), produtos_estoque.nome
    """, parametros)

    saldos = []
    for row in cursor.fetchall():
        item = dict(row)
        item["status_saldo"] = montar_status_saldo(item["saldo_atual"], item["estoque_minimo"])

        if status and item["status_saldo"] != status:
            continue

        saldos.append(item)

    conn.close()
    return saldos


def carregar_resumo_saldo_por_categoria():
    saldos = carregar_saldo_estoque()
    resumo = {}

    for item in saldos:
        categoria = item["categoria_nome"] or "Sem categoria"

        if categoria not in resumo:
            resumo[categoria] = {
                "categoria": categoria,
                "produtos": 0,
                "ok": 0,
                "baixo": 0,
                "zerado": 0,
                "negativo": 0
            }

        resumo[categoria]["produtos"] += 1

        if item["status_saldo"] == "OK":
            resumo[categoria]["ok"] += 1
        elif item["status_saldo"] == "Baixo":
            resumo[categoria]["baixo"] += 1
        elif item["status_saldo"] == "Zerado":
            resumo[categoria]["zerado"] += 1
        elif item["status_saldo"] == "Negativo":
            resumo[categoria]["negativo"] += 1

    return list(resumo.values())


def contar_produtos_baixo_estoque():
    saldos = carregar_saldo_estoque()
    total = 0

    for item in saldos:
        if item["estoque_minimo"] > 0 and item["saldo_atual"] <= item["estoque_minimo"]:
            total += 1

    return total


def buscar_item_compra_para_estoque(item_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            itens_compra.*,
            pedidos_compra.id AS pedido_id,
            pedidos_compra.solicitante,
            categorias_compra.nome AS categoria_nome,
            fornecedores.nome AS fornecedor_nome
        FROM itens_compra
        INNER JOIN pedidos_compra ON pedidos_compra.id = itens_compra.pedido_id
        INNER JOIN categorias_compra ON categorias_compra.id = itens_compra.categoria_id
        LEFT JOIN fornecedores ON fornecedores.id = itens_compra.fornecedor_id
        WHERE itens_compra.id = ?
    """, (item_id,))

    item = cursor.fetchone()
    conn.close()
    return item


def buscar_pedido_compra(pedido_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            pedidos_compra.*,
            fornecedores.nome AS fornecedor_nome
        FROM pedidos_compra
        LEFT JOIN fornecedores ON fornecedores.id = pedidos_compra.fornecedor_id
        WHERE pedidos_compra.id = ?
    """, (pedido_id,))

    pedido = cursor.fetchone()
    conn.close()
    return pedido


def carregar_itens_pedido_compra(pedido_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            itens_compra.*,
            categorias_compra.nome AS categoria_nome,
            fornecedores.nome AS fornecedor_nome
        FROM itens_compra
        INNER JOIN categorias_compra ON categorias_compra.id = itens_compra.categoria_id
        LEFT JOIN fornecedores ON fornecedores.id = itens_compra.fornecedor_id
        WHERE itens_compra.pedido_id = ?
        ORDER BY itens_compra.id
    """, (pedido_id,))

    itens = cursor.fetchall()
    conn.close()
    return itens


def atualizar_status_pedido_compra(pedido_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN status = 'Comprado' THEN 1 ELSE 0 END) AS comprados,
            SUM(CASE WHEN status = 'Cancelado' THEN 1 ELSE 0 END) AS cancelados
        FROM itens_compra
        WHERE pedido_id = ?
    """, (pedido_id,))

    resumo = cursor.fetchone()
    total = resumo["total"] or 0
    comprados = resumo["comprados"] or 0
    cancelados = resumo["cancelados"] or 0

    if total > 0 and comprados == total:
        novo_status = "Comprado"
    elif total > 0 and cancelados == total:
        novo_status = "Cancelado"
    else:
        novo_status = "Pendente"

    cursor.execute("""
        UPDATE pedidos_compra
        SET status = ?
        WHERE id = ?
    """, (novo_status, pedido_id))

    conn.commit()
    conn.close()


# ============================================================
# FICHA TÉCNICA V2.0
# ============================================================

def _agora_texto():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def _usuario_atual():
    if has_request_context():
        return session.get("usuario", "sistema")
    return "sistema"


def _linha_para_dict(linha):
    return dict(linha) if linha is not None else None


def contar_fichas_tecnicas_ativas():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM fichas_tecnicas
        WHERE ativo = 1
    """)

    total = cursor.fetchone()["total"]
    conn.close()
    return total


def carregar_fichas_tecnicas(busca="", ativo=""):
    conn = conectar()
    cursor = conn.cursor()

    filtros = []
    parametros = []

    busca = str(busca or "").strip()
    ativo = str(ativo or "").strip()

    if busca:
        filtros.append("""
            (produtos_estoque.nome LIKE ?
            OR produtos_estoque.codigo LIKE ?
            OR fichas_tecnicas.tipo LIKE ?)
        """)
        termo = f"%{busca}%"
        parametros.extend([termo, termo, termo])

    if ativo in ["0", "1"]:
        filtros.append("fichas_tecnicas.ativo = ?")
        parametros.append(int(ativo))

    where_sql = "WHERE " + " AND ".join(filtros) if filtros else ""

    cursor.execute(f"""
        SELECT
            fichas_tecnicas.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS produto_unidade,
            categorias_estoque.nome AS categoria_nome,
            COUNT(DISTINCT CASE WHEN itens_ficha_tecnica.ativo = 1 THEN itens_ficha_tecnica.id END) AS total_itens,
            COUNT(DISTINCT ficha_tecnica_utilizacoes.id) AS total_utilizacoes
        FROM fichas_tecnicas
        INNER JOIN produtos_estoque ON produtos_estoque.id = fichas_tecnicas.produto_final_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        LEFT JOIN itens_ficha_tecnica ON itens_ficha_tecnica.ficha_id = fichas_tecnicas.id
        LEFT JOIN ficha_tecnica_utilizacoes ON ficha_tecnica_utilizacoes.ficha_id = fichas_tecnicas.id
        {where_sql}
        GROUP BY fichas_tecnicas.id
        ORDER BY CAST(produtos_estoque.codigo AS INTEGER), produtos_estoque.nome
    """, parametros)

    fichas = cursor.fetchall()
    conn.close()
    return fichas


def buscar_ficha_tecnica(ficha_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            fichas_tecnicas.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS produto_unidade,
            categorias_estoque.nome AS categoria_nome,
            (
                SELECT COUNT(*)
                FROM ficha_tecnica_utilizacoes
                WHERE ficha_tecnica_utilizacoes.ficha_id = fichas_tecnicas.id
            ) AS total_utilizacoes
        FROM fichas_tecnicas
        INNER JOIN produtos_estoque ON produtos_estoque.id = fichas_tecnicas.produto_final_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE fichas_tecnicas.id = ?
    """, (ficha_id,))

    ficha = cursor.fetchone()
    conn.close()
    return ficha


def buscar_ficha_tecnica_por_produto(produto_final_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id
        FROM fichas_tecnicas
        WHERE produto_final_id = ?
        LIMIT 1
    """, (produto_final_id,))

    ficha = cursor.fetchone()
    conn.close()
    return ficha


def carregar_produtos_sem_ficha(excluir_produto_id=None):
    conn = conectar()
    cursor = conn.cursor()

    parametros = []
    filtro_exclusao = ""

    if excluir_produto_id:
        filtro_exclusao = "AND produtos_estoque.id != ?"
        parametros.append(int(excluir_produto_id))

    cursor.execute(f"""
        SELECT
            produtos_estoque.id,
            produtos_estoque.codigo,
            produtos_estoque.nome,
            produtos_estoque.unidade_padrao,
            categorias_estoque.nome AS categoria_nome
        FROM produtos_estoque
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        LEFT JOIN fichas_tecnicas ON fichas_tecnicas.produto_final_id = produtos_estoque.id
        WHERE produtos_estoque.ativo = 1
          AND fichas_tecnicas.id IS NULL
          {filtro_exclusao}
        ORDER BY CAST(produtos_estoque.codigo AS INTEGER), produtos_estoque.nome
    """, parametros)

    produtos = cursor.fetchall()
    conn.close()
    return produtos


def carregar_itens_ficha_tecnica(ficha_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            itens_ficha_tecnica.*,
            produtos_estoque.codigo AS insumo_codigo,
            produtos_estoque.nome AS insumo_nome,
            produtos_estoque.unidade_padrao AS insumo_unidade,
            categorias_estoque.nome AS categoria_nome
        FROM itens_ficha_tecnica
        INNER JOIN produtos_estoque ON produtos_estoque.id = itens_ficha_tecnica.insumo_produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE itens_ficha_tecnica.ficha_id = ?
        ORDER BY itens_ficha_tecnica.ativo DESC, produtos_estoque.nome
    """, (ficha_id,))

    itens = cursor.fetchall()
    conn.close()
    return itens


def buscar_item_ficha_tecnica(item_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            itens_ficha_tecnica.*,
            produtos_estoque.codigo AS insumo_codigo,
            produtos_estoque.nome AS insumo_nome,
            produtos_estoque.unidade_padrao AS insumo_unidade,
            fichas_tecnicas.produto_final_id,
            produtos_finais.codigo AS produto_final_codigo,
            produtos_finais.nome AS produto_final_nome
        FROM itens_ficha_tecnica
        INNER JOIN produtos_estoque ON produtos_estoque.id = itens_ficha_tecnica.insumo_produto_id
        INNER JOIN fichas_tecnicas ON fichas_tecnicas.id = itens_ficha_tecnica.ficha_id
        INNER JOIN produtos_estoque AS produtos_finais ON produtos_finais.id = fichas_tecnicas.produto_final_id
        WHERE itens_ficha_tecnica.id = ?
    """, (item_id,))

    item = cursor.fetchone()
    conn.close()
    return item


def montar_snapshot_ficha(ficha_id):
    ficha = buscar_ficha_tecnica(ficha_id)
    if not ficha:
        return None

    itens = carregar_itens_ficha_tecnica(ficha_id)

    return {
        "ficha": _linha_para_dict(ficha),
        "itens": [_linha_para_dict(item) for item in itens]
    }


def registrar_revisao_ficha(ficha_id, acao):
    snapshot = montar_snapshot_ficha(ficha_id)
    if not snapshot:
        return

    ficha = snapshot["ficha"]
    versao = int(ficha.get("versao") or 1)

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT OR REPLACE INTO ficha_tecnica_revisoes
        (ficha_id, versao, acao, dados_json, data_hora, usuario)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        ficha_id,
        versao,
        str(acao or "Atualização").strip(),
        json.dumps(snapshot, ensure_ascii=False),
        _agora_texto(),
        _usuario_atual()
    ))

    conn.commit()
    conn.close()


def garantir_revisoes_iniciais_fichas():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("SELECT id FROM fichas_tecnicas ORDER BY id")
    fichas = cursor.fetchall()
    conn.close()

    for ficha in fichas:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM ficha_tecnica_revisoes
            WHERE ficha_id = ?
        """, (ficha["id"],))
        total = cursor.fetchone()["total"]
        conn.close()

        if total == 0:
            registrar_revisao_ficha(ficha["id"], "Versão inicial registrada")


def carregar_revisoes_ficha(ficha_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, ficha_id, versao, acao, data_hora, usuario
        FROM ficha_tecnica_revisoes
        WHERE ficha_id = ?
        ORDER BY versao DESC, id DESC
    """, (ficha_id,))

    revisoes = cursor.fetchall()
    conn.close()
    return revisoes


def buscar_revisao_ficha(ficha_id, revisao_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, ficha_id, versao, acao, dados_json, data_hora, usuario
        FROM ficha_tecnica_revisoes
        WHERE ficha_id = ? AND id = ?
    """, (ficha_id, revisao_id))

    revisao = cursor.fetchone()
    conn.close()

    if not revisao:
        return None, None

    try:
        dados = json.loads(revisao["dados_json"])
    except Exception:
        dados = {"ficha": {}, "itens": []}

    return revisao, dados


def contar_utilizacoes_ficha(ficha_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM ficha_tecnica_utilizacoes
        WHERE ficha_id = ?
    """, (ficha_id,))

    total = cursor.fetchone()["total"]
    conn.close()
    return total


def registrar_utilizacao_ficha_cursor(cursor, ficha_id, origem, origem_id=None, observacao=""):
    snapshot = montar_snapshot_ficha(ficha_id)
    if not snapshot:
        raise ValueError("Ficha tecnica nao encontrada.")

    versao = int(snapshot["ficha"].get("versao") or 1)

    cursor.execute("""
        INSERT INTO ficha_tecnica_utilizacoes
        (ficha_id, versao_ficha, origem, origem_id, dados_json, data_hora, usuario, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        ficha_id,
        versao,
        str(origem or "Uso nao informado").strip(),
        origem_id,
        json.dumps(snapshot, ensure_ascii=False),
        _agora_texto(),
        _usuario_atual(),
        str(observacao or "").strip()
    ))

    return cursor.lastrowid


def registrar_utilizacao_ficha(ficha_id, origem, origem_id=None, observacao=""):
    """Prepara a ficha para o futuro planejamento/produção.

    Quando uma ficha for usada por um módulo futuro, esta função congela a
    versão e a composição utilizadas. Isso preserva o histórico mesmo que a
    receita seja alterada depois.
    """
    snapshot = montar_snapshot_ficha(ficha_id)
    if not snapshot:
        raise ValueError("Ficha técnica não encontrada.")

    versao = int(snapshot["ficha"].get("versao") or 1)

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO ficha_tecnica_utilizacoes
        (ficha_id, versao_ficha, origem, origem_id, dados_json, data_hora, usuario, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        ficha_id,
        versao,
        str(origem or "Uso não informado").strip(),
        origem_id,
        json.dumps(snapshot, ensure_ascii=False),
        _agora_texto(),
        _usuario_atual(),
        str(observacao or "").strip()
    ))

    utilizacao_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return utilizacao_id


def _incrementar_versao_ficha(cursor, ficha_id):
    cursor.execute("""
        UPDATE fichas_tecnicas
        SET versao = COALESCE(versao, 1) + 1,
            atualizado_em = ?,
            usuario = ?
        WHERE id = ?
    """, (_agora_texto(), _usuario_atual(), ficha_id))


def criar_ficha_tecnica(produto_final_id, tipo, rendimento_quantidade, rendimento_unidade, observacao=""):
    try:
        produto_final_id = int(produto_final_id)
    except Exception:
        raise ValueError("Produto final inválido.")

    produto = buscar_produto_estoque(produto_final_id)
    if not produto:
        raise ValueError("Produto oficial não encontrado.")

    if buscar_ficha_tecnica_por_produto(produto_final_id):
        raise ValueError("Este produto já possui ficha técnica cadastrada.")

    try:
        rendimento_quantidade = float(str(rendimento_quantidade or 1).replace(",", "."))
    except Exception:
        rendimento_quantidade = 1

    if rendimento_quantidade <= 0:
        raise ValueError("O rendimento precisa ser maior que zero.")

    tipo = str(tipo or "Produto acabado").strip()
    rendimento_unidade = str(rendimento_unidade or produto["unidade_padrao"] or "un").strip()
    observacao = str(observacao or "").strip()
    agora = _agora_texto()

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO fichas_tecnicas
        (produto_final_id, tipo, rendimento_quantidade, rendimento_unidade, ativo, versao, data_cadastro, atualizado_em, usuario, observacao)
        VALUES (?, ?, ?, ?, 1, 1, ?, ?, ?, ?)
    """, (
        produto_final_id,
        tipo,
        rendimento_quantidade,
        rendimento_unidade,
        agora,
        agora,
        _usuario_atual(),
        observacao
    ))

    ficha_id = cursor.lastrowid
    conn.commit()
    conn.close()

    registrar_revisao_ficha(ficha_id, "Ficha criada")
    return ficha_id


def atualizar_ficha_tecnica(ficha_id, tipo, rendimento_quantidade, rendimento_unidade, observacao=""):
    ficha = buscar_ficha_tecnica(ficha_id)
    if not ficha:
        raise ValueError("Ficha técnica não encontrada.")

    try:
        rendimento_quantidade = float(str(rendimento_quantidade or 1).replace(",", "."))
    except Exception:
        rendimento_quantidade = 1

    if rendimento_quantidade <= 0:
        raise ValueError("O rendimento precisa ser maior que zero.")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE fichas_tecnicas
        SET tipo = ?,
            rendimento_quantidade = ?,
            rendimento_unidade = ?,
            observacao = ?
        WHERE id = ?
    """, (
        str(tipo or "Produto acabado").strip(),
        rendimento_quantidade,
        str(rendimento_unidade or "un").strip(),
        str(observacao or "").strip(),
        ficha_id
    ))

    _incrementar_versao_ficha(cursor, ficha_id)
    conn.commit()
    conn.close()

    registrar_revisao_ficha(ficha_id, "Dados gerais atualizados")


def adicionar_item_ficha_tecnica(ficha_id, insumo_produto_id, quantidade, unidade, observacao=""):
    ficha = buscar_ficha_tecnica(ficha_id)
    if not ficha:
        raise ValueError("Ficha técnica não encontrada.")

    try:
        insumo_produto_id = int(insumo_produto_id)
    except Exception:
        raise ValueError("Insumo inválido.")

    if insumo_produto_id == ficha["produto_final_id"]:
        raise ValueError("O produto final não pode ser componente dele mesmo.")

    insumo = buscar_produto_estoque(insumo_produto_id)
    if not insumo:
        raise ValueError("Produto/Insumo não encontrado.")

    try:
        quantidade = float(str(quantidade or 0).replace(",", "."))
    except Exception:
        quantidade = 0

    if quantidade <= 0:
        raise ValueError("A quantidade do item precisa ser maior que zero.")

    unidade = str(unidade or insumo["unidade_padrao"] or "un").strip()
    agora = _agora_texto()

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO itens_ficha_tecnica
        (ficha_id, insumo_produto_id, quantidade, unidade, ativo, atualizado_em, usuario, observacao)
        VALUES (?, ?, ?, ?, 1, ?, ?, ?)
    """, (
        ficha_id,
        insumo_produto_id,
        quantidade,
        unidade,
        agora,
        _usuario_atual(),
        str(observacao or "").strip()
    ))

    _incrementar_versao_ficha(cursor, ficha_id)
    conn.commit()
    conn.close()

    registrar_revisao_ficha(ficha_id, f"Componente adicionado: {insumo['nome']}")


def atualizar_item_ficha_tecnica(item_id, insumo_produto_id, quantidade, unidade, observacao=""):
    item = buscar_item_ficha_tecnica(item_id)
    if not item:
        raise ValueError("Item da ficha técnica não encontrado.")

    try:
        insumo_produto_id = int(insumo_produto_id)
    except Exception:
        raise ValueError("Insumo inválido.")

    if insumo_produto_id == item["produto_final_id"]:
        raise ValueError("O produto final não pode ser componente dele mesmo.")

    insumo = buscar_produto_estoque(insumo_produto_id)
    if not insumo:
        raise ValueError("Produto/Insumo não encontrado.")

    try:
        quantidade = float(str(quantidade or 0).replace(",", "."))
    except Exception:
        quantidade = 0

    if quantidade <= 0:
        raise ValueError("A quantidade precisa ser maior que zero.")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE itens_ficha_tecnica
        SET insumo_produto_id = ?,
            quantidade = ?,
            unidade = ?,
            atualizado_em = ?,
            usuario = ?,
            observacao = ?
        WHERE id = ?
    """, (
        insumo_produto_id,
        quantidade,
        str(unidade or insumo["unidade_padrao"] or "un").strip(),
        _agora_texto(),
        _usuario_atual(),
        str(observacao or "").strip(),
        item_id
    ))

    _incrementar_versao_ficha(cursor, item["ficha_id"])
    conn.commit()
    conn.close()

    registrar_revisao_ficha(item["ficha_id"], f"Componente editado: {insumo['nome']}")
    return item["ficha_id"]


def alternar_item_ficha_tecnica(item_id):
    item = buscar_item_ficha_tecnica(item_id)
    if not item:
        raise ValueError("Item da ficha técnica não encontrado.")

    novo_status = 0 if item["ativo"] == 1 else 1

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE itens_ficha_tecnica
        SET ativo = ?, atualizado_em = ?, usuario = ?
        WHERE id = ?
    """, (novo_status, _agora_texto(), _usuario_atual(), item_id))

    _incrementar_versao_ficha(cursor, item["ficha_id"])
    conn.commit()
    conn.close()

    acao = "Componente ativado" if novo_status == 1 else "Componente inativado"
    registrar_revisao_ficha(item["ficha_id"], f"{acao}: {item['insumo_nome']}")
    return item["ficha_id"]


def remover_item_ficha_tecnica(item_id):
    item = buscar_item_ficha_tecnica(item_id)
    if not item:
        raise ValueError("Item da ficha técnica não encontrado.")

    if contar_utilizacoes_ficha(item["ficha_id"]) > 0:
        raise ValueError("Esta ficha já foi utilizada. Para preservar o histórico, inative o componente em vez de removê-lo.")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM itens_ficha_tecnica WHERE id = ?", (item_id,))
    _incrementar_versao_ficha(cursor, item["ficha_id"])

    conn.commit()
    conn.close()

    registrar_revisao_ficha(item["ficha_id"], f"Componente removido: {item['insumo_nome']}")
    return item["ficha_id"]


def alternar_ficha_tecnica(ficha_id):
    ficha = buscar_ficha_tecnica(ficha_id)
    if not ficha:
        raise ValueError("Ficha técnica não encontrada.")

    novo_status = 0 if ficha["ativo"] == 1 else 1

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE fichas_tecnicas
        SET ativo = ?
        WHERE id = ?
    """, (novo_status, ficha_id))

    _incrementar_versao_ficha(cursor, ficha_id)
    conn.commit()
    conn.close()

    acao = "Ficha ativada" if novo_status == 1 else "Ficha inativada"
    registrar_revisao_ficha(ficha_id, acao)


def duplicar_ficha_tecnica(ficha_id, produto_destino_id, observacao_extra=""):
    ficha = buscar_ficha_tecnica(ficha_id)
    if not ficha:
        raise ValueError("Ficha técnica de origem não encontrada.")

    try:
        produto_destino_id = int(produto_destino_id)
    except Exception:
        raise ValueError("Produto de destino inválido.")

    produto_destino = buscar_produto_estoque(produto_destino_id)
    if not produto_destino:
        raise ValueError("Produto de destino não encontrado.")

    if buscar_ficha_tecnica_por_produto(produto_destino_id):
        raise ValueError("O produto de destino já possui ficha técnica.")

    itens = carregar_itens_ficha_tecnica(ficha_id)
    agora = _agora_texto()
    observacao_base = str(ficha["observacao"] or "").strip()
    observacao_extra = str(observacao_extra or "").strip()
    observacao_final = observacao_base

    if observacao_extra:
        observacao_final = f"{observacao_base} | {observacao_extra}" if observacao_base else observacao_extra

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO fichas_tecnicas
        (produto_final_id, tipo, rendimento_quantidade, rendimento_unidade, ativo, versao, data_cadastro, atualizado_em, usuario, observacao)
        VALUES (?, ?, ?, ?, 1, 1, ?, ?, ?, ?)
    """, (
        produto_destino_id,
        ficha["tipo"],
        ficha["rendimento_quantidade"],
        ficha["rendimento_unidade"],
        agora,
        agora,
        _usuario_atual(),
        observacao_final
    ))

    nova_ficha_id = cursor.lastrowid

    for item in itens:
        cursor.execute("""
            INSERT INTO itens_ficha_tecnica
            (ficha_id, insumo_produto_id, quantidade, unidade, ativo, atualizado_em, usuario, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            nova_ficha_id,
            item["insumo_produto_id"],
            item["quantidade"],
            item["unidade"],
            item["ativo"],
            agora,
            _usuario_atual(),
            item["observacao"]
        ))

    conn.commit()
    conn.close()

    registrar_revisao_ficha(
        nova_ficha_id,
        f"Ficha duplicada de {ficha['produto_codigo']} - {ficha['produto_nome']}"
    )
    return nova_ficha_id


def excluir_ficha_tecnica(ficha_id):
    ficha = buscar_ficha_tecnica(ficha_id)
    if not ficha:
        raise ValueError("Ficha técnica não encontrada.")

    if contar_utilizacoes_ficha(ficha_id) > 0:
        raise ValueError("Esta ficha já foi utilizada em outro processo e não pode ser excluída. Inative-a para preservar o histórico.")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM itens_ficha_tecnica WHERE ficha_id = ?", (ficha_id,))
    cursor.execute("DELETE FROM ficha_tecnica_revisoes WHERE ficha_id = ?", (ficha_id,))
    cursor.execute("DELETE FROM ficha_tecnica_utilizacoes WHERE ficha_id = ?", (ficha_id,))
    cursor.execute("DELETE FROM fichas_tecnicas WHERE id = ?", (ficha_id,))

    conn.commit()
    conn.close()


def simular_consumo_ficha(ficha, itens, quantidade_produzida):
    try:
        quantidade_produzida = float(str(quantidade_produzida or 0).replace(",", "."))
    except Exception:
        quantidade_produzida = 0

    rendimento = float(ficha["rendimento_quantidade"] or 1)
    fator = quantidade_produzida / rendimento if rendimento > 0 else 0

    resultado = []

    for item in itens:
        if item["ativo"] != 1:
            continue

        resultado.append({
            "insumo_codigo": item["insumo_codigo"],
            "insumo_nome": item["insumo_nome"],
            "quantidade_base": float(item["quantidade"] or 0),
            "unidade": item["unidade"],
            "quantidade_prevista": float(item["quantidade"] or 0) * fator,
            "observacao": item["observacao"] or ""
        })

    return {
        "quantidade_produzida": quantidade_produzida,
        "fator": fator,
        "itens": resultado
    }

def inicializar_banco():
    criar_banco()
    garantir_colunas_novas()

    if banco_vazio():
        popular_banco_inicial()

    popular_usuarios_iniciais()
    popular_categorias_compra_iniciais()
    popular_categorias_estoque_iniciais()
    popular_centros_custo_iniciais()
    popular_produtos_estoque_iniciais()
    popular_embalagens_padrao_empadas()
    popular_origens_expedicao_iniciais()
    popular_motivos_perda_iniciais()
    garantir_perdas_legadas()
    migrar_pedidos_existentes_para_rodadas()
    garantir_revisoes_iniciais_fichas()
    garantir_separacoes_planejamentos_existentes()


def carregar_sabores():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, nome, classe, chuva, normal, verao, baixa
        FROM sabores
        ORDER BY id
    """)

    sabores = []

    for row in cursor.fetchall():
        sabores.append({
            "id": row["id"],
            "nome": row["nome"],
            "classe": row["classe"],
            "metas": {
                "Chuva": row["chuva"],
                "Normal": row["normal"],
                "Verão": row["verao"],
                "Baixa": row["baixa"]
            }
        })

    conn.close()
    return sabores


def carregar_metas_por_dia():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            metas_dia.dia_semana,
            sabores.nome AS nome_sabor,
            metas_dia.meta
        FROM metas_dia
        INNER JOIN sabores ON sabores.id = metas_dia.sabor_id
    """)

    metas_por_dia = {dia: {} for dia in DIAS_SEMANA}

    for row in cursor.fetchall():
        metas_por_dia[row["dia_semana"]][row["nome_sabor"]] = row["meta"]

    conn.close()
    return metas_por_dia


# ============================================================
# LOGIN / PERMISSÕES
# ============================================================

def buscar_usuario(usuario):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, usuario, senha_hash, perfil, loja_id
        FROM usuarios
        WHERE usuario = ?
    """, (usuario,))

    usuario_encontrado = cursor.fetchone()
    conn.close()

    return usuario_encontrado


def login_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        return funcao(*args, **kwargs)

    return wrapper


def admin_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") != "admin":
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para acessar esta página."
            )

        return funcao(*args, **kwargs)

    return wrapper


def compras_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "compras"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para acessar o módulo de compras."
            )

        return funcao(*args, **kwargs)

    return wrapper


def admin_ou_compras_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "compras"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para acessar esta página."
            )

        return funcao(*args, **kwargs)

    return wrapper


def producao_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "producao", "estoque"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para acessar o módulo de produção."
            )

        return funcao(*args, **kwargs)

    return wrapper


def pode_solicitar_compra_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "estoque"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Somente admin e estoque podem solicitar compras."
            )

        return funcao(*args, **kwargs)

    return wrapper


def pode_ver_compras_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "compras", "estoque"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para acessar o módulo de compras."
            )

        return funcao(*args, **kwargs)

    return wrapper


def estoque_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "estoque", "compras"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para acessar o módulo de estoque."
            )

        return funcao(*args, **kwargs)

    return wrapper



def operacao_estoque_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "estoque"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Somente admin e estoque podem separar ou confirmar solicitações internas."
            )

        return funcao(*args, **kwargs)

    return wrapper



def consumo_producao_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))

        if session.get("perfil") not in ["admin", "producao", "estoque"]:
            return render_template(
                "acesso_negado.html",
                mensagem="Você não tem permissão para registrar consumo real da produção."
            )

        return funcao(*args, **kwargs)

    return wrapper


def loja_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))
        if session.get("perfil") not in ["loja", "admin"]:
            return render_template("acesso_negado.html", mensagem="Acesso exclusivo para lojas.")
        if session.get("perfil") == "loja" and not session.get("loja_id"):
            return render_template("acesso_negado.html", mensagem="Este usuário não está vinculado a uma loja.")
        return funcao(*args, **kwargs)
    return wrapper


def expedicao_obrigatorio(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))
        if session.get("perfil") not in ["admin", "estoque"]:
            return render_template("acesso_negado.html", mensagem="Somente admin e estoque podem operar a expedição.")
        return funcao(*args, **kwargs)
    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    erro = None

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        senha = request.form.get("senha", "").strip()

        usuario_banco = buscar_usuario(usuario)

        if usuario_banco and check_password_hash(usuario_banco["senha_hash"], senha):
            session["usuario"] = usuario_banco["usuario"]
            session["perfil"] = usuario_banco["perfil"]
            session["loja_id"] = usuario_banco["loja_id"]

            if usuario_banco["perfil"] == "loja":
                return redirect(url_for("loja_painel"))
            return redirect(url_for("dashboard"))

        erro = "Usuário ou senha inválidos."

    return render_template("login.html", erro=erro)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))



# ============================================================
# PLANEJAMENTO DE PRODUÇÃO / CONSUMO PREVISTO
# ============================================================

def normalizar_unidade_planejamento(unidade):
    valor = str(unidade or "").strip()
    chave = valor.lower().replace(".", "")

    aliases = {
        "u": "un",
        "un": "un",
        "und": "un",
        "unid": "un",
        "unidade": "un",
        "unidades": "un",
        "g": "g",
        "grama": "g",
        "gramas": "g",
        "kg": "kg",
        "quilo": "kg",
        "quilos": "kg",
        "quilograma": "kg",
        "quilogramas": "kg",
        "ml": "ml",
        "mililitro": "ml",
        "mililitros": "ml",
        "l": "L",
        "litro": "L",
        "litros": "L",
        "pct": "pct",
        "pacote": "pct",
        "pacotes": "pct",
        "cx": "cx",
        "caixa": "cx",
        "caixas": "cx",
        "fardo": "fardo",
        "fardos": "fardo",
        "saco": "saco",
        "sacos": "saco",
        "balde": "balde",
        "baldes": "balde"
    }

    return aliases.get(chave, valor)


def converter_quantidade_planejamento(quantidade, unidade_origem, unidade_destino):
    quantidade = float(quantidade or 0)
    origem = normalizar_unidade_planejamento(unidade_origem)
    destino = normalizar_unidade_planejamento(unidade_destino)

    if origem == destino:
        return quantidade

    conversoes = {
        "g": ("massa", 1.0),
        "kg": ("massa", 1000.0),
        "ml": ("volume", 1.0),
        "L": ("volume", 1000.0),
        "un": ("contagem", 1.0)
    }

    if origem not in conversoes or destino not in conversoes:
        raise ValueError(f"Não há conversão automática entre {unidade_origem} e {unidade_destino}.")

    grupo_origem, fator_origem = conversoes[origem]
    grupo_destino, fator_destino = conversoes[destino]

    if grupo_origem != grupo_destino:
        raise ValueError(f"Unidades incompatíveis: {unidade_origem} e {unidade_destino}.")

    quantidade_base = quantidade * fator_origem
    return quantidade_base / fator_destino


def buscar_ficha_tecnica_ativa_por_produto(produto_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            fichas_tecnicas.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS produto_unidade
        FROM fichas_tecnicas
        INNER JOIN produtos_estoque ON produtos_estoque.id = fichas_tecnicas.produto_final_id
        WHERE fichas_tecnicas.produto_final_id = ?
          AND fichas_tecnicas.ativo = 1
        LIMIT 1
    """, (produto_id,))

    ficha = cursor.fetchone()
    conn.close()
    return ficha


def carregar_itens_ativos_ficha_planejamento(ficha_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            itens_ficha_tecnica.*,
            produtos_estoque.codigo AS insumo_codigo,
            produtos_estoque.nome AS insumo_nome,
            produtos_estoque.unidade_padrao AS insumo_unidade,
            categorias_estoque.nome AS categoria_nome
        FROM itens_ficha_tecnica
        INNER JOIN produtos_estoque ON produtos_estoque.id = itens_ficha_tecnica.insumo_produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE itens_ficha_tecnica.ficha_id = ?
          AND itens_ficha_tecnica.ativo = 1
        ORDER BY produtos_estoque.nome
    """, (ficha_id,))

    itens = cursor.fetchall()
    conn.close()
    return itens


def contar_planejamentos_abertos():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM planejamentos_producao
        WHERE status IN ('Rascunho', 'Com alertas')
    """)
    total = cursor.fetchone()["total"]
    conn.close()
    return total


def _explodir_ficha_planejamento(ficha, fator, caminho, nivel, pilha, fichas_usadas, alertas):
    registros = []
    ficha_id = int(ficha["id"])
    versao_ficha = int(ficha["versao"] or 1)

    if ficha_id in pilha:
        mensagem = f"Ciclo de ficha técnica detectado em {ficha['produto_nome']}."
        alertas.append(mensagem)
        return registros

    nova_pilha = set(pilha)
    nova_pilha.add(ficha_id)
    fichas_usadas[ficha_id] = {
        "versao": versao_ficha,
        "caminhos": fichas_usadas.get(ficha_id, {}).get("caminhos", []) + [caminho]
    }

    itens = carregar_itens_ativos_ficha_planejamento(ficha_id)
    if not itens:
        mensagem = f"A ficha de {ficha['produto_nome']} não possui componentes ativos."
        alertas.append(mensagem)
        return registros

    for item in itens:
        quantidade_prevista_original = float(item["quantidade"] or 0) * float(fator or 0)
        caminho_item = f"{caminho} > {item['insumo_nome']}"
        ficha_componente = buscar_ficha_tecnica_ativa_por_produto(item["insumo_produto_id"])
        alerta_item = None

        try:
            quantidade_prevista = converter_quantidade_planejamento(
                quantidade_prevista_original,
                item["unidade"],
                item["insumo_unidade"]
            )
            unidade_prevista = normalizar_unidade_planejamento(item["insumo_unidade"])
        except Exception as e:
            quantidade_prevista = quantidade_prevista_original
            unidade_prevista = normalizar_unidade_planejamento(item["unidade"])
            alerta_item = (
                f"{item['insumo_nome']}: não foi possível converter {item['unidade']} "
                f"para a unidade base {item['insumo_unidade']} ({str(e)})."
            )
            alertas.append(alerta_item)

        registro = {
            "produto_id": item["insumo_produto_id"],
            "produto_codigo": item["insumo_codigo"],
            "produto_nome": item["insumo_nome"],
            "tipo": "Intermediário" if ficha_componente else "Insumo",
            "quantidade": quantidade_prevista,
            "unidade": unidade_prevista,
            "nivel": nivel,
            "caminho": caminho_item,
            "ficha_origem_id": ficha_id,
            "versao_ficha_origem": versao_ficha,
            "ficha_componente_id": ficha_componente["id"] if ficha_componente else None,
            "versao_ficha_componente": int(ficha_componente["versao"] or 1) if ficha_componente else None,
            "alerta": alerta_item,
            "observacao": item["observacao"] or ""
        }

        registros.append(registro)

        if ficha_componente:
            try:
                rendimento = float(ficha_componente["rendimento_quantidade"] or 0)
                if rendimento <= 0:
                    raise ValueError("o rendimento da ficha é zero")

                quantidade_na_unidade_rendimento = converter_quantidade_planejamento(
                    quantidade_prevista,
                    unidade_prevista,
                    ficha_componente["rendimento_unidade"]
                )
                fator_componente = quantidade_na_unidade_rendimento / rendimento

                registros.extend(
                    _explodir_ficha_planejamento(
                        ficha_componente,
                        fator_componente,
                        caminho_item,
                        nivel + 1,
                        nova_pilha,
                        fichas_usadas,
                        alertas
                    )
                )
            except Exception as e:
                alerta_item = (
                    f"Não foi possível abrir a ficha de {item['insumo_nome']}: {str(e)}"
                )
                registro["alerta"] = alerta_item
                alertas.append(alerta_item)

    return registros


def _criar_planejamento_producao_legacy_v351(resultado, cenario, dia_semana, total_tabuleiros, total_empadas, rodada_id=None):
    itens_validos = [item for item in resultado if float(item.get("producao", 0) or 0) > 0]
    if not itens_validos:
        raise ValueError("Não há itens com produção maior que zero para planejar.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO planejamentos_producao
        (data_criacao, cenario, dia_semana, total_tabuleiros, total_unidades, status, origem, criado_por, observacao, rodada_id)
        VALUES (?, ?, ?, ?, ?, 'Rascunho', ?, ?, ?, ?)
    """, (
        _agora_texto(),
        cenario,
        dia_semana,
        float(total_tabuleiros or 0),
        float(total_empadas or 0),
        "Pedidos das Lojas" if rodada_id else "Meta de Produção",
        _usuario_atual(),
        "Planejamento criado a partir da demanda fechada das lojas e da meta de segurança." if rodada_id else "Planejamento criado a partir da ordem calculada. Não movimenta estoque; serve para previsão, separação e conferência.",
        int(rodada_id) if rodada_id else None,
    ))
    planejamento_id = cursor.lastrowid
    conn.commit()
    conn.close()

    itens_processados = []
    fichas_usadas = {}
    alertas_gerais = []

    for item in itens_validos:
        sabor = item["nome"]
        quantidade_tabuleiros = float(item["producao"] or 0)
        quantidade_unidades = float(item["empadas"] or (quantidade_tabuleiros * EMPADAS_POR_TABULEIRO))
        nome_produto = f"Empada de {sabor}"
        produto = buscar_produto_estoque_por_nome(nome_produto)
        alerta_item = []
        consumos = []
        ficha = None
        snapshot = None

        if not produto:
            alerta_item.append(f"Produto oficial não encontrado: {nome_produto}.")
        else:
            ficha = buscar_ficha_tecnica_ativa_por_produto(produto["id"])
            if not ficha:
                alerta_item.append(f"O produto {nome_produto} não possui ficha técnica ativa.")
            else:
                snapshot = montar_snapshot_ficha(ficha["id"])
                try:
                    rendimento = float(ficha["rendimento_quantidade"] or 0)
                    if rendimento <= 0:
                        raise ValueError("rendimento da ficha igual a zero")

                    quantidade_na_unidade_rendimento = converter_quantidade_planejamento(
                        quantidade_unidades,
                        produto["unidade_padrao"],
                        ficha["rendimento_unidade"]
                    )
                    fator = quantidade_na_unidade_rendimento / rendimento
                    consumos = _explodir_ficha_planejamento(
                        ficha,
                        fator,
                        nome_produto,
                        1,
                        set(),
                        fichas_usadas,
                        alerta_item
                    )
                except Exception as e:
                    alerta_item.append(f"Não foi possível calcular a ficha de {nome_produto}: {str(e)}")

        alertas_gerais.extend(alerta_item)
        itens_processados.append({
            "sabor": sabor,
            "produto": produto,
            "quantidade_tabuleiros": quantidade_tabuleiros,
            "quantidade_unidades": quantidade_unidades,
            "ficha": ficha,
            "snapshot": snapshot,
            "alertas": alerta_item,
            "consumos": consumos
        })

    utilizacoes = {}
    for ficha_id, dados in fichas_usadas.items():
        try:
            utilizacoes[ficha_id] = registrar_utilizacao_ficha(
                ficha_id,
                "Planejamento de Produção",
                planejamento_id,
                observacao=" | ".join(sorted(set(dados.get("caminhos", []))))
            )
        except Exception as e:
            alertas_gerais.append(f"Não foi possível congelar a ficha #{ficha_id}: {str(e)}")

    conn = conectar()
    cursor = conn.cursor()

    for item in itens_processados:
        ficha = item["ficha"]
        alerta_texto = " | ".join(item["alertas"])
        status_item = "Alerta" if item["alertas"] else "Planejado"

        cursor.execute("""
            INSERT INTO itens_planejamento_producao
            (planejamento_id, sabor, produto_final_id, quantidade_tabuleiros, quantidade_unidades,
             ficha_id, versao_ficha, utilizacao_ficha_id, status, alerta, snapshot_json, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            planejamento_id,
            item["sabor"],
            item["produto"]["id"] if item["produto"] else None,
            item["quantidade_tabuleiros"],
            item["quantidade_unidades"],
            ficha["id"] if ficha else None,
            int(ficha["versao"] or 1) if ficha else None,
            utilizacoes.get(ficha["id"]) if ficha else None,
            status_item,
            alerta_texto or None,
            json.dumps(item["snapshot"], ensure_ascii=False) if item["snapshot"] else None,
            "Quantidade prevista pela meta. Conferir antes de produzir."
        ))
        item_planejamento_id = cursor.lastrowid

        for consumo in item["consumos"]:
            cursor.execute("""
                INSERT INTO consumos_previstos_planejamento
                (planejamento_id, item_planejamento_id, produto_id, tipo, quantidade, unidade,
                 nivel, caminho, ficha_origem_id, versao_ficha_origem,
                 ficha_componente_id, versao_ficha_componente, alerta, observacao)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                planejamento_id,
                item_planejamento_id,
                consumo["produto_id"],
                consumo["tipo"],
                consumo["quantidade"],
                consumo["unidade"],
                consumo["nivel"],
                consumo["caminho"],
                consumo["ficha_origem_id"],
                consumo["versao_ficha_origem"],
                consumo["ficha_componente_id"],
                consumo["versao_ficha_componente"],
                consumo["alerta"],
                consumo["observacao"]
            ))

    status_planejamento = "Com alertas" if alertas_gerais else "Rascunho"
    cursor.execute("""
        UPDATE planejamentos_producao
        SET status = ?, observacao = CASE
            WHEN ? = '' THEN observacao
            ELSE observacao || ' | Alertas: ' || ?
        END
        WHERE id = ?
    """, (
        status_planejamento,
        " | ".join(alertas_gerais),
        " | ".join(alertas_gerais),
        planejamento_id
    ))

    conn.commit()
    conn.close()

    separacao_id = criar_separacao_estoque_planejamento(planejamento_id)

    pre_producao_id = criar_pre_producao_realizada(
        resultado,
        cenario,
        dia_semana,
        total_tabuleiros,
        total_empadas,
        origem=f"Planejamento #{planejamento_id}",
        planejamento_id=planejamento_id
    )

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE planejamentos_producao
        SET producao_realizada_id = ?
        WHERE id = ?
    """, (pre_producao_id, planejamento_id))

    if rodada_id:
        cursor.execute("""
            UPDATE rodadas_pedidos_loja
            SET status = 'Planejada', planejamento_id = ?
            WHERE id = ? AND status = 'Fechada'
        """, (planejamento_id, int(rodada_id)))

    conn.commit()
    conn.close()

    return planejamento_id, pre_producao_id


def criar_planejamento_producao(resultado, cenario, dia_semana, total_tabuleiros, total_empadas, rodada_id=None):
    itens_validos = [item for item in resultado if float(item.get("empadas", 0) or 0) > 0]
    if not itens_validos:
        raise ValueError("Nao ha itens com producao maior que zero para planejar.")

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("""
            INSERT INTO planejamentos_producao
            (data_criacao, cenario, dia_semana, total_tabuleiros, total_unidades, status, origem, criado_por, observacao, rodada_id)
            VALUES (?, ?, ?, ?, ?, 'Rascunho', ?, ?, ?, ?)
        """, (
            _agora_texto(),
            cenario,
            dia_semana,
            float(total_tabuleiros or 0),
            float(total_empadas or 0),
            "Pedidos das Lojas" if rodada_id else "Meta de Produção",
            _usuario_atual(),
            "Planejamento criado a partir da demanda fechada das lojas e da meta de seguranca." if rodada_id else "Planejamento criado a partir da ordem calculada. Nao movimenta estoque; serve para previsao, separacao e conferencia.",
            int(rodada_id) if rodada_id else None,
        ))
        planejamento_id = cursor.lastrowid

        itens_processados = []
        fichas_usadas = {}
        alertas_gerais = []

        for item in itens_validos:
            sabor = item["nome"]
            quantidade_tabuleiros = float(item["producao"] or 0)
            quantidade_unidades = float(item["empadas"] or (quantidade_tabuleiros * EMPADAS_POR_TABULEIRO))
            nome_produto = f"Empada de {sabor}"
            produto = buscar_produto_estoque_por_nome_cursor(cursor, nome_produto)
            alerta_item = []
            consumos = []
            ficha = None
            snapshot = None

            if not produto:
                alerta_item.append(f"Produto oficial nao encontrado: {nome_produto}.")
            else:
                ficha = buscar_ficha_tecnica_ativa_por_produto(produto["id"])
                if not ficha:
                    alerta_item.append(f"O produto {nome_produto} nao possui ficha tecnica ativa.")
                else:
                    snapshot = montar_snapshot_ficha(ficha["id"])
                    try:
                        rendimento = float(ficha["rendimento_quantidade"] or 0)
                        if rendimento <= 0:
                            raise ValueError("rendimento da ficha igual a zero")

                        quantidade_na_unidade_rendimento = converter_quantidade_planejamento(
                            quantidade_unidades,
                            produto["unidade_padrao"],
                            ficha["rendimento_unidade"]
                        )
                        fator = quantidade_na_unidade_rendimento / rendimento
                        consumos = _explodir_ficha_planejamento(
                            ficha,
                            fator,
                            nome_produto,
                            1,
                            set(),
                            fichas_usadas,
                            alerta_item
                        )
                    except Exception as e:
                        alerta_item.append(f"Nao foi possivel calcular a ficha de {nome_produto}: {str(e)}")

            alertas_gerais.extend(alerta_item)
            itens_processados.append({
                "sabor": sabor,
                "produto": produto,
                "quantidade_tabuleiros": quantidade_tabuleiros,
                "quantidade_unidades": quantidade_unidades,
                "ficha": ficha,
                "snapshot": snapshot,
                "alertas": alerta_item,
                "consumos": consumos
            })

        utilizacoes = {}
        for ficha_id, dados in fichas_usadas.items():
            try:
                utilizacoes[ficha_id] = registrar_utilizacao_ficha_cursor(
                    cursor,
                    ficha_id,
                    "Planejamento de Produção",
                    planejamento_id,
                    observacao=" | ".join(sorted(set(dados.get("caminhos", []))))
                )
            except Exception as e:
                alertas_gerais.append(f"Nao foi possivel congelar a ficha #{ficha_id}: {str(e)}")

        for item in itens_processados:
            ficha = item["ficha"]
            alerta_texto = " | ".join(item["alertas"])
            status_item = "Alerta" if item["alertas"] else "Planejado"

            cursor.execute("""
                INSERT INTO itens_planejamento_producao
                (planejamento_id, sabor, produto_final_id, quantidade_tabuleiros, quantidade_unidades,
                 ficha_id, versao_ficha, utilizacao_ficha_id, status, alerta, snapshot_json, observacao)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                planejamento_id,
                item["sabor"],
                item["produto"]["id"] if item["produto"] else None,
                item["quantidade_tabuleiros"],
                item["quantidade_unidades"],
                ficha["id"] if ficha else None,
                int(ficha["versao"] or 1) if ficha else None,
                utilizacoes.get(ficha["id"]) if ficha else None,
                status_item,
                alerta_texto or None,
                json.dumps(item["snapshot"], ensure_ascii=False) if item["snapshot"] else None,
                "Quantidade prevista pela meta. Conferir antes de produzir."
            ))
            item_planejamento_id = cursor.lastrowid

            for consumo in item["consumos"]:
                cursor.execute("""
                    INSERT INTO consumos_previstos_planejamento
                    (planejamento_id, item_planejamento_id, produto_id, tipo, quantidade, unidade,
                     nivel, caminho, ficha_origem_id, versao_ficha_origem,
                     ficha_componente_id, versao_ficha_componente, alerta, observacao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    planejamento_id,
                    item_planejamento_id,
                    consumo["produto_id"],
                    consumo["tipo"],
                    consumo["quantidade"],
                    consumo["unidade"],
                    consumo["nivel"],
                    consumo["caminho"],
                    consumo["ficha_origem_id"],
                    consumo["versao_ficha_origem"],
                    consumo["ficha_componente_id"],
                    consumo["versao_ficha_componente"],
                    consumo["alerta"],
                    consumo["observacao"]
                ))

        status_planejamento = "Com alertas" if alertas_gerais else "Rascunho"
        cursor.execute("""
            UPDATE planejamentos_producao
            SET status = ?, observacao = CASE
                WHEN ? = '' THEN observacao
                ELSE observacao || ' | Alertas: ' || ?
            END
            WHERE id = ?
        """, (
            status_planejamento,
            " | ".join(alertas_gerais),
            " | ".join(alertas_gerais),
            planejamento_id
        ))

        criar_separacao_estoque_planejamento_cursor(cursor, planejamento_id)

        pre_producao_id = criar_pre_producao_realizada(
            resultado,
            cenario,
            dia_semana,
            total_tabuleiros,
            total_empadas,
            origem=f"Planejamento #{planejamento_id}",
            planejamento_id=planejamento_id,
            cursor=cursor
        )

        cursor.execute("""
            UPDATE planejamentos_producao
            SET producao_realizada_id = ?
            WHERE id = ?
        """, (pre_producao_id, planejamento_id))

        if rodada_id:
            cursor.execute("""
                UPDATE rodadas_pedidos_loja
                SET status = 'Planejada', planejamento_id = ?
                WHERE id = ? AND status = 'Fechada'
            """, (planejamento_id, int(rodada_id)))

        conn.commit()
        return planejamento_id, pre_producao_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def carregar_planejamentos_producao(status=""):
    conn = conectar()
    cursor = conn.cursor()
    parametros = []
    filtro = ""

    if status:
        filtro = "WHERE planejamentos_producao.status = ?"
        parametros.append(status)

    cursor.execute(f"""
        SELECT
            planejamentos_producao.*,
            separacoes_estoque.id AS separacao_id,
            separacoes_estoque.status AS separacao_status,
            COUNT(DISTINCT itens_planejamento_producao.id) AS total_itens,
            SUM(CASE WHEN itens_planejamento_producao.status = 'Alerta' THEN 1 ELSE 0 END) AS total_alertas
        FROM planejamentos_producao
        LEFT JOIN itens_planejamento_producao ON itens_planejamento_producao.planejamento_id = planejamentos_producao.id
        LEFT JOIN separacoes_estoque ON separacoes_estoque.planejamento_id = planejamentos_producao.id
        {filtro}
        GROUP BY planejamentos_producao.id
        ORDER BY planejamentos_producao.id DESC
    """, parametros)

    planejamentos = cursor.fetchall()
    conn.close()
    return planejamentos


def buscar_planejamento_producao(planejamento_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT *
        FROM planejamentos_producao
        WHERE id = ?
    """, (planejamento_id,))
    planejamento = cursor.fetchone()
    conn.close()
    return planejamento


def carregar_itens_planejamento_producao(planejamento_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            itens_planejamento_producao.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS produto_unidade
        FROM itens_planejamento_producao
        LEFT JOIN produtos_estoque ON produtos_estoque.id = itens_planejamento_producao.produto_final_id
        WHERE itens_planejamento_producao.planejamento_id = ?
        ORDER BY itens_planejamento_producao.id
    """, (planejamento_id,))
    itens = cursor.fetchall()
    conn.close()
    return itens


def carregar_consumos_consolidados_planejamento(planejamento_id, tipo):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            produtos_estoque.id AS produto_id,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            categorias_estoque.nome AS categoria_nome,
            consumos_previstos_planejamento.tipo,
            consumos_previstos_planejamento.unidade,
            SUM(consumos_previstos_planejamento.quantidade) AS quantidade_total,
            MAX(consumos_previstos_planejamento.alerta) AS alerta,
            GROUP_CONCAT(DISTINCT consumos_previstos_planejamento.caminho) AS caminhos
        FROM consumos_previstos_planejamento
        INNER JOIN produtos_estoque ON produtos_estoque.id = consumos_previstos_planejamento.produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE consumos_previstos_planejamento.planejamento_id = ?
          AND consumos_previstos_planejamento.tipo = ?
        GROUP BY produtos_estoque.id, consumos_previstos_planejamento.unidade, consumos_previstos_planejamento.tipo
        ORDER BY categorias_estoque.nome, produtos_estoque.nome
    """, (planejamento_id, tipo))
    consumos = cursor.fetchall()
    conn.close()
    return consumos


def carregar_consumos_consolidados_planejamento_cursor(cursor, planejamento_id, tipo):
    cursor.execute("""
        SELECT
            produtos_estoque.id AS produto_id,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            categorias_estoque.nome AS categoria_nome,
            consumos_previstos_planejamento.tipo,
            consumos_previstos_planejamento.unidade,
            SUM(consumos_previstos_planejamento.quantidade) AS quantidade_total,
            MAX(consumos_previstos_planejamento.alerta) AS alerta,
            GROUP_CONCAT(DISTINCT consumos_previstos_planejamento.caminho) AS caminhos
        FROM consumos_previstos_planejamento
        INNER JOIN produtos_estoque ON produtos_estoque.id = consumos_previstos_planejamento.produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE consumos_previstos_planejamento.planejamento_id = ?
          AND consumos_previstos_planejamento.tipo = ?
        GROUP BY produtos_estoque.id, consumos_previstos_planejamento.unidade, consumos_previstos_planejamento.tipo
        ORDER BY categorias_estoque.nome, produtos_estoque.nome
    """, (planejamento_id, tipo))
    return cursor.fetchall()


def carregar_consumos_detalhados_planejamento(planejamento_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            consumos_previstos_planejamento.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            itens_planejamento_producao.sabor
        FROM consumos_previstos_planejamento
        INNER JOIN produtos_estoque ON produtos_estoque.id = consumos_previstos_planejamento.produto_id
        INNER JOIN itens_planejamento_producao ON itens_planejamento_producao.id = consumos_previstos_planejamento.item_planejamento_id
        WHERE consumos_previstos_planejamento.planejamento_id = ?
        ORDER BY itens_planejamento_producao.id, consumos_previstos_planejamento.nivel, consumos_previstos_planejamento.id
    """, (planejamento_id,))
    consumos = cursor.fetchall()
    conn.close()
    return consumos



def criar_separacao_estoque_planejamento_cursor(cursor, planejamento_id):
    cursor.execute("""
        SELECT *
        FROM planejamentos_producao
        WHERE id = ?
    """, (planejamento_id,))
    planejamento = cursor.fetchone()
    if not planejamento:
        raise ValueError("Planejamento nao encontrado para criar a separacao.")

    cursor.execute("""
        SELECT id
        FROM separacoes_estoque
        WHERE planejamento_id = ?
    """, (planejamento_id,))
    existente = cursor.fetchone()

    if existente:
        return existente["id"]

    insumos = carregar_consumos_consolidados_planejamento_cursor(cursor, planejamento_id, "Insumo")

    if planejamento["status"] == "Cancelado":
        status_inicial = "Cancelado"
    elif planejamento["status"] == "Com alertas":
        status_inicial = "Aguardando revisão"
    elif planejamento["status"] == "Produção confirmada":
        status_inicial = "Encerrado"
    elif not insumos:
        status_inicial = "Sem itens"
    else:
        status_inicial = "Pendente"

    cursor.execute("""
        INSERT INTO separacoes_estoque
        (planejamento_id, data_criacao, status, criado_por, observacao)
        VALUES (?, ?, ?, ?, ?)
    """, (
        planejamento_id,
        _agora_texto(),
        status_inicial,
        _usuario_atual(),
        "Lista criada automaticamente pelo planejamento. Nao movimenta saldo; registra apenas a separacao fisica prevista."
    ))
    separacao_id = cursor.lastrowid

    for item in insumos:
        quantidade_prevista = float(item["quantidade_total"] or 0)
        saldo_atual = obter_saldo_produto_cursor(cursor, item["produto_id"])
        alertas = []
        if item["alerta"]:
            alertas.append(str(item["alerta"]))
        if saldo_atual < quantidade_prevista:
            alertas.append(
                f"Saldo atual insuficiente na criacao: {saldo_atual:.4f} {item['unidade']} para {quantidade_prevista:.4f} {item['unidade']} previstos."
            )

        cursor.execute("""
            INSERT OR IGNORE INTO itens_separacao_estoque
            (separacao_id, produto_id, quantidade_prevista, unidade,
             quantidade_separada, saldo_no_planejamento, status, alerta, observacao)
            VALUES (?, ?, ?, ?, 0, ?, 'Pendente', ?, ?)
        """, (
            separacao_id,
            item["produto_id"],
            quantidade_prevista,
            item["unidade"],
            saldo_atual,
            " | ".join(alertas) or None,
            "Quantidade consolidada a partir das fichas tecnicas congeladas."
        ))

    return separacao_id


def criar_separacao_estoque_planejamento(planejamento_id):
    planejamento = buscar_planejamento_producao(planejamento_id)
    if not planejamento:
        raise ValueError("Planejamento não encontrado para criar a separação.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id
        FROM separacoes_estoque
        WHERE planejamento_id = ?
    """, (planejamento_id,))
    existente = cursor.fetchone()
    conn.close()

    if existente:
        return existente["id"]

    insumos = carregar_consumos_consolidados_planejamento(planejamento_id, "Insumo")

    if planejamento["status"] == "Cancelado":
        status_inicial = "Cancelado"
    elif planejamento["status"] == "Com alertas":
        status_inicial = "Aguardando revisão"
    elif planejamento["status"] == "Produção confirmada":
        status_inicial = "Encerrado"
    elif not insumos:
        status_inicial = "Sem itens"
    else:
        status_inicial = "Pendente"

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO separacoes_estoque
        (planejamento_id, data_criacao, status, criado_por, observacao)
        VALUES (?, ?, ?, ?, ?)
    """, (
        planejamento_id,
        _agora_texto(),
        status_inicial,
        _usuario_atual(),
        "Lista criada automaticamente pelo planejamento. Não movimenta saldo; registra apenas a separação física prevista."
    ))
    separacao_id = cursor.lastrowid

    for item in insumos:
        quantidade_prevista = float(item["quantidade_total"] or 0)
        saldo_atual = obter_saldo_produto(item["produto_id"])
        alertas = []
        if item["alerta"]:
            alertas.append(str(item["alerta"]))
        if saldo_atual < quantidade_prevista:
            alertas.append(
                f"Saldo atual insuficiente na criação: {saldo_atual:.4f} {item['unidade']} para {quantidade_prevista:.4f} {item['unidade']} previstos."
            )

        cursor.execute("""
            INSERT OR IGNORE INTO itens_separacao_estoque
            (separacao_id, produto_id, quantidade_prevista, unidade,
             quantidade_separada, saldo_no_planejamento, status, alerta, observacao)
            VALUES (?, ?, ?, ?, 0, ?, 'Pendente', ?, ?)
        """, (
            separacao_id,
            item["produto_id"],
            quantidade_prevista,
            item["unidade"],
            saldo_atual,
            " | ".join(alertas) or None,
            "Quantidade consolidada a partir das fichas técnicas congeladas."
        ))

    conn.commit()
    conn.close()
    return separacao_id


def garantir_separacoes_planejamentos_existentes():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT planejamentos_producao.id
        FROM planejamentos_producao
        LEFT JOIN separacoes_estoque
            ON separacoes_estoque.planejamento_id = planejamentos_producao.id
        WHERE separacoes_estoque.id IS NULL
        ORDER BY planejamentos_producao.id
    """)
    ids = [row["id"] for row in cursor.fetchall()]
    conn.close()

    for planejamento_id in ids:
        try:
            criar_separacao_estoque_planejamento(planejamento_id)
        except Exception:
            # Não impede o sistema de iniciar por causa de um planejamento antigo incompleto.
            pass


def buscar_separacao_por_planejamento(planejamento_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT *
        FROM separacoes_estoque
        WHERE planejamento_id = ?
    """, (planejamento_id,))
    separacao = cursor.fetchone()
    conn.close()
    return separacao


def buscar_separacao_estoque(separacao_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            separacoes_estoque.*,
            planejamentos_producao.cenario,
            planejamentos_producao.dia_semana,
            planejamentos_producao.total_tabuleiros,
            planejamentos_producao.total_unidades,
            planejamentos_producao.status AS planejamento_status,
            planejamentos_producao.producao_realizada_id
        FROM separacoes_estoque
        INNER JOIN planejamentos_producao
            ON planejamentos_producao.id = separacoes_estoque.planejamento_id
        WHERE separacoes_estoque.id = ?
    """, (separacao_id,))
    separacao = cursor.fetchone()
    conn.close()
    return separacao


def carregar_itens_separacao_estoque(separacao_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            itens_separacao_estoque.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS produto_unidade,
            categorias_estoque.nome AS categoria_nome
        FROM itens_separacao_estoque
        INNER JOIN produtos_estoque
            ON produtos_estoque.id = itens_separacao_estoque.produto_id
        LEFT JOIN categorias_estoque
            ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE itens_separacao_estoque.separacao_id = ?
        ORDER BY categorias_estoque.nome, produtos_estoque.nome
    """, (separacao_id,))
    rows = cursor.fetchall()
    conn.close()

    itens = []
    for row in rows:
        item = dict(row)
        item["saldo_atual"] = obter_saldo_produto(item["produto_id"])
        item["quantidade_faltante"] = max(
            float(item["quantidade_prevista"] or 0) - float(item["quantidade_separada"] or 0),
            0
        )
        item["diferenca"] = float(item["quantidade_separada"] or 0) - float(item["quantidade_prevista"] or 0)
        item["quantidade_transferida"] = quantidade_transferida_separacao_item(item["id"])
        item["quantidade_a_transferir"] = max(
            float(item["quantidade_separada"] or 0) - float(item["quantidade_transferida"] or 0),
            0
        )
        item["disponibilidade"] = "Disponível" if item["saldo_atual"] >= item["quantidade_faltante"] else "Saldo insuficiente"
        itens.append(item)

    return itens


def recalcular_status_separacao(separacao_id):
    separacao = buscar_separacao_estoque(separacao_id)
    if not separacao:
        raise ValueError("Separação não encontrada.")

    if separacao["status"] in ["Cancelado", "Encerrado"]:
        return separacao["status"]

    itens = carregar_itens_separacao_estoque(separacao_id)
    if not itens:
        novo_status = "Sem itens"
    else:
        total_completo = sum(
            1 for item in itens
            if float(item["quantidade_separada"] or 0) >= float(item["quantidade_prevista"] or 0)
        )
        total_iniciado = sum(1 for item in itens if float(item["quantidade_separada"] or 0) > 0)

        if total_completo == len(itens):
            novo_status = "Separado"
        elif total_iniciado > 0:
            novo_status = "Parcial"
        elif separacao["planejamento_status"] == "Com alertas":
            novo_status = "Aguardando revisão"
        else:
            novo_status = "Pendente"

    agora = _agora_texto()
    usuario = _usuario_atual()
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE separacoes_estoque
        SET status = ?, data_atualizacao = ?, atualizado_por = ?,
            data_conclusao = CASE WHEN ? = 'Separado' THEN COALESCE(data_conclusao, ?) ELSE NULL END,
            concluido_por = CASE WHEN ? = 'Separado' THEN COALESCE(concluido_por, ?) ELSE NULL END
        WHERE id = ?
    """, (
        novo_status, agora, usuario,
        novo_status, agora,
        novo_status, usuario,
        separacao_id
    ))
    conn.commit()
    conn.close()
    return novo_status


def atualizar_separacao_estoque(separacao_id, form):
    separacao = buscar_separacao_estoque(separacao_id)
    if not separacao:
        raise ValueError("Separação não encontrada.")
    if separacao["status"] in ["Cancelado", "Encerrado"]:
        raise ValueError("Esta separação está encerrada e não pode ser alterada.")
    if separacao["planejamento_status"] == "Cancelado":
        raise ValueError("O planejamento vinculado foi cancelado.")

    itens = carregar_itens_separacao_estoque(separacao_id)
    conn = conectar()
    cursor = conn.cursor()
    agora = _agora_texto()
    usuario = _usuario_atual()

    for item in itens:
        campo_quantidade = f"quantidade_separada_{item['id']}"
        campo_observacao = f"observacao_{item['id']}"
        valor_bruto = str(form.get(campo_quantidade, item["quantidade_separada"]) or "0").strip().replace(",", ".")
        try:
            quantidade_separada = float(valor_bruto)
        except ValueError:
            conn.close()
            raise ValueError(f"Quantidade inválida para {item['produto_nome']}.")

        if quantidade_separada < 0:
            conn.close()
            raise ValueError(f"A quantidade separada de {item['produto_nome']} não pode ser negativa.")

        prevista = float(item["quantidade_prevista"] or 0)
        if quantidade_separada <= 0:
            status_item = "Pendente"
        elif quantidade_separada < prevista:
            status_item = "Parcial"
        else:
            status_item = "Separado"

        observacao = str(form.get(campo_observacao, item["observacao"] or "") or "").strip()
        cursor.execute("""
            UPDATE itens_separacao_estoque
            SET quantidade_separada = ?, status = ?, observacao = ?,
                atualizado_em = ?, atualizado_por = ?
            WHERE id = ? AND separacao_id = ?
        """, (
            quantidade_separada, status_item, observacao,
            agora, usuario, item["id"], separacao_id
        ))

    observacao_geral = str(form.get("observacao_geral", separacao["observacao"] or "") or "").strip()
    cursor.execute("""
        UPDATE separacoes_estoque
        SET observacao = ?, data_atualizacao = ?, atualizado_por = ?
        WHERE id = ?
    """, (observacao_geral, agora, usuario, separacao_id))
    conn.commit()
    conn.close()

    return recalcular_status_separacao(separacao_id)


def carregar_separacoes_estoque(status="", busca=""):
    filtros = []
    parametros = []
    if status:
        filtros.append("separacoes_estoque.status = ?")
        parametros.append(status)
    if busca:
        filtros.append("(CAST(separacoes_estoque.id AS TEXT) LIKE ? OR CAST(separacoes_estoque.planejamento_id AS TEXT) LIKE ? OR planejamentos_producao.dia_semana LIKE ?)")
        termo = f"%{busca}%"
        parametros.extend([termo, termo, termo])

    where_sql = "WHERE " + " AND ".join(filtros) if filtros else ""
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT
            separacoes_estoque.*,
            planejamentos_producao.data_criacao AS planejamento_data,
            planejamentos_producao.cenario,
            planejamentos_producao.dia_semana,
            planejamentos_producao.status AS planejamento_status,
            COUNT(itens_separacao_estoque.id) AS total_itens,
            SUM(CASE WHEN itens_separacao_estoque.status = 'Separado' THEN 1 ELSE 0 END) AS itens_separados,
            SUM(CASE WHEN itens_separacao_estoque.alerta IS NOT NULL AND itens_separacao_estoque.alerta != '' THEN 1 ELSE 0 END) AS total_alertas
        FROM separacoes_estoque
        INNER JOIN planejamentos_producao
            ON planejamentos_producao.id = separacoes_estoque.planejamento_id
        LEFT JOIN itens_separacao_estoque
            ON itens_separacao_estoque.separacao_id = separacoes_estoque.id
        {where_sql}
        GROUP BY separacoes_estoque.id
        ORDER BY separacoes_estoque.id DESC
    """, parametros)
    separacoes = cursor.fetchall()
    conn.close()
    return separacoes


def contar_separacoes_pendentes():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM separacoes_estoque
        WHERE status IN ('Pendente', 'Parcial', 'Aguardando revisão')
    """)
    total = cursor.fetchone()["total"]
    conn.close()
    return int(total or 0)

def cancelar_planejamento_producao(planejamento_id):
    planejamento = buscar_planejamento_producao(planejamento_id)
    if not planejamento:
        raise ValueError("Planejamento não encontrado.")

    if planejamento["status"] not in ["Rascunho", "Com alertas"]:
        raise ValueError("Apenas planejamentos abertos podem ser cancelados.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE planejamentos_producao
        SET status = 'Cancelado',
            observacao = CASE
                WHEN observacao IS NULL OR observacao = '' THEN ?
                ELSE observacao || ' | ' || ?
            END
        WHERE id = ?
    """, (
        f"Cancelado por {_usuario_atual()} em {_agora_texto()}.",
        f"Cancelado por {_usuario_atual()} em {_agora_texto()}.",
        planejamento_id
    ))

    if planejamento["producao_realizada_id"]:
        cursor.execute("""
            UPDATE producoes_realizadas
            SET status = 'Cancelado', data_cancelamento = ?, cancelado_por = ?
            WHERE id = ? AND status = 'Pendente'
        """, (_agora_texto(), _usuario_atual(), planejamento["producao_realizada_id"]))
        cursor.execute("""
            UPDATE itens_producao_realizada
            SET status = 'Cancelado'
            WHERE producao_realizada_id = ? AND status = 'Pendente'
        """, (planejamento["producao_realizada_id"],))

    cursor.execute("""
        UPDATE separacoes_estoque
        SET status = 'Cancelado', data_atualizacao = ?, atualizado_por = ?
        WHERE planejamento_id = ?
    """, (_agora_texto(), _usuario_atual(), planejamento_id))

    conn.commit()
    conn.close()

# ============================================================
# CÁLCULO DE PRODUÇÃO
# ============================================================

def calcular_producao(form):
    sabores = carregar_sabores()
    metas_por_dia = carregar_metas_por_dia()

    resultado = []
    cenario = form.get("cenario", "Baixa")
    dia_semana = form.get("dia_semana", "Segunda-feira")
    rodada_id = str(form.get("rodada_id", "") or "").strip()
    rodada = buscar_rodada_pedidos(int(rodada_id)) if rodada_id.isdigit() else None
    demanda = demanda_sabores_rodada(int(rodada_id)) if rodada and rodada["status"] in ["Fechada", "Planejada"] else {}

    if rodada:
        cenario = rodada["cenario"] or cenario
        dia_semana = rodada["dia_semana"] or dia_semana

    total_tabuleiros_normais = 0.0
    total_empadas = 0.0
    valores_digitados = {}

    for i, sabor in enumerate(sabores):
        nome_sabor = sabor["nome"]
        inicio = int(float(form.get(f"inicio_{i}", 0) or 0))
        info_demanda = demanda.get(nome_sabor, {})

        if rodada:
            pedido_normal_unidades = float(info_demanda.get("normal_unidades", 0) or 0)
            pedido_misto_unidades = float(info_demanda.get("misto_unidades", 0) or 0)
            pedido_normal_tabuleiros = pedido_normal_unidades / EMPADAS_POR_TABULEIRO
        else:
            pedido_normal_tabuleiros = float(form.get(f"pedido_{i}", 0) or 0)
            pedido_normal_unidades = pedido_normal_tabuleiros * EMPADAS_POR_TABULEIRO
            pedido_misto_unidades = 0.0

        valores_digitados[i] = {
            "inicio": inicio,
            "pedido": pedido_normal_tabuleiros,
            "pedido_unidades": pedido_normal_unidades,
            "misto_unidades": pedido_misto_unidades,
        }

        meta_cenario = sabor["metas"][cenario]
        metas_do_dia = metas_por_dia.get(dia_semana, {})
        meta_dia = metas_do_dia.get(nome_sabor)

        if dia_semana in DIAS_EXCLUSIVOS:
            if meta_dia is not None:
                meta = meta_dia
            else:
                meta = 0
            origem_meta = dia_semana
            meta_especial = True
        else:
            if meta_dia is not None:
                meta = meta_dia
                origem_meta = dia_semana
                meta_especial = True
            else:
                meta = meta_cenario
                origem_meta = cenario
                meta_especial = False

        necessidade_normal_unidades = max(
            (float(meta) * EMPADAS_POR_TABULEIRO)
            - (float(inicio) * EMPADAS_POR_TABULEIRO)
            + pedido_normal_unidades,
            0.0,
        )
        producao_normal_tabuleiros = int(math.ceil((necessidade_normal_unidades - 1e-9) / EMPADAS_POR_TABULEIRO)) if necessidade_normal_unidades > 0 else 0
        producao_normal_unidades = producao_normal_tabuleiros * EMPADAS_POR_TABULEIRO
        quantidade_total_unidades = producao_normal_unidades + pedido_misto_unidades
        estoque_final_tabuleiros = float(inicio) + producao_normal_tabuleiros - pedido_normal_tabuleiros

        total_tabuleiros_normais += producao_normal_tabuleiros
        total_empadas += quantidade_total_unidades

        resultado.append({
            "nome": nome_sabor,
            "classe": sabor["classe"],
            "inicio": inicio,
            "pedido": pedido_normal_tabuleiros,
            "pedido_unidades": pedido_normal_unidades,
            "misto_unidades": pedido_misto_unidades,
            "meta": meta,
            "origem_meta": origem_meta,
            "meta_especial": meta_especial,
            "producao": producao_normal_tabuleiros,
            "producao_normal_unidades": producao_normal_unidades,
            "estoque_final": estoque_final_tabuleiros,
            "empadas": quantidade_total_unidades,
        })

    total_mistos = total_mistos_rodada(int(rodada_id)) if rodada else 0.0
    total_producao = total_tabuleiros_normais + total_mistos
    return resultado, cenario, dia_semana, total_producao, total_empadas, valores_digitados


# ============================================================
# IMPRESSÃO DIRETA - ARGOX
# ============================================================

def mm_para_px(mm, dpi=PRINTER_DPI):
    return int((mm / 25.4) * dpi)


ETIQUETA_LARGURA_PX = mm_para_px(ETIQUETA_LARGURA_MM)
ETIQUETA_ALTURA_PX = mm_para_px(ETIQUETA_ALTURA_MM)


def carregar_fontes_etiqueta():
    try:
        fonte_empresa = ImageFont.truetype("arialbd.ttf", 24)
        fonte_sabor = ImageFont.truetype("arialbd.ttf", 44)
        fonte_info_titulo = ImageFont.truetype("arialbd.ttf", 24)
        fonte_info = ImageFont.truetype("arial.ttf", 24)
        fonte_pequena = ImageFont.truetype("arial.ttf", 20)
    except:
        fonte_empresa = ImageFont.load_default()
        fonte_sabor = ImageFont.load_default()
        fonte_info_titulo = ImageFont.load_default()
        fonte_info = ImageFont.load_default()
        fonte_pequena = ImageFont.load_default()

    return fonte_empresa, fonte_sabor, fonte_info_titulo, fonte_info, fonte_pequena


def quebrar_texto_etiqueta(texto, limite=14):
    palavras = texto.split()
    linhas = []
    atual = ""

    for palavra in palavras:
        teste = palavra if atual == "" else atual + " " + palavra

        if len(teste) <= limite:
            atual = teste
        else:
            if atual:
                linhas.append(atual)
            atual = palavra

    if atual:
        linhas.append(atual)

    return linhas[:3]


def texto_centralizado(draw, x, y, texto, fonte, fill="black"):
    draw.text((x, y), texto, font=fonte, fill=fill, anchor="mm")


def desenhar_linhas_centralizadas(draw, x, y_inicial, linhas, fonte, espacamento=26, fill="black"):
    for i, linha in enumerate(linhas):
        draw.text(
            (x, y_inicial + i * espacamento),
            linha,
            font=fonte,
            fill=fill,
            anchor="mm"
        )


def criar_imagem_etiqueta(etiqueta):
    img = Image.new("RGB", (ETIQUETA_LARGURA_PX, ETIQUETA_ALTURA_PX), "white")
    draw = ImageDraw.Draw(img)

    fonte_empresa, fonte_sabor, fonte_info_titulo, fonte_info, fonte_pequena = carregar_fontes_etiqueta()

    margem = 10

    # Borda
    draw.rectangle(
        [2, 2, ETIQUETA_LARGURA_PX - 3, ETIQUETA_ALTURA_PX - 3],
        outline="black",
        width=2
    )

    # Cabeçalho
    texto_centralizado(
        draw,
        ETIQUETA_LARGURA_PX // 2,
        22,
        "EMPADINHAS BARNABÉ",
        fonte_empresa
    )

    draw.line(
        [(margem, 38), (ETIQUETA_LARGURA_PX - margem, 38)],
        fill="black",
        width=1
    )

    # Controle de qualidade
    if etiqueta["tipo"] == "controle_qualidade":
        texto_centralizado(draw, ETIQUETA_LARGURA_PX // 2, 80, "CONTROLE DE", fonte_info_titulo)
        texto_centralizado(draw, ETIQUETA_LARGURA_PX // 2, 108, "QUALIDADE", fonte_info_titulo)
        texto_centralizado(draw, ETIQUETA_LARGURA_PX // 2, 138, "Tabuleiro com todos os sabores", fonte_pequena)

        draw.text((18, 185), "Data:", font=fonte_info_titulo, fill="black")
        draw.text((120, 185), etiqueta["data"], font=fonte_info, fill="black")

        draw.text((18, 213), "Validade:", font=fonte_info_titulo, fill="black")
        draw.text((120, 213), etiqueta["validade"], font=fonte_info, fill="black")

        draw.text((18, 241), "Qtd:", font=fonte_info_titulo, fill="black")
        draw.text((120, 241), f'{etiqueta["quantidade_empadas"]} unidades', font=fonte_info, fill="black")

        return img

    # Tabuleiro misto
    if etiqueta["tipo"] == "misto":
        texto_centralizado(draw, ETIQUETA_LARGURA_PX // 2, 80, "TAB: MISTO", fonte_info_titulo) #Aqui da pra definir o que é escrito antes

        sabor_1 = etiqueta["sabor_1"].upper()
        sabor_2 = etiqueta["sabor_2"].upper()

        draw.text((18, 120), f"{sabor_1}: 18", font=fonte_pequena, fill="black")
        draw.text((18, 148), f"{sabor_2}: 17", font=fonte_pequena, fill="black")

        draw.text((18, 190), "Data:", font=fonte_info_titulo, fill="black")
        draw.text((120, 190), etiqueta["data"], font=fonte_info, fill="black")

        draw.text((18, 218), "Validade:", font=fonte_info_titulo, fill="black")
        draw.text((120, 218), etiqueta["validade"], font=fonte_info, fill="black")

        if "numero" in etiqueta and "total" in etiqueta:
            draw.text(
                (18, 246),
                f'Tab. {etiqueta["numero"]}/{etiqueta["total"]}',
                font=fonte_pequena,
                fill="black"
            )

        return img

    # Tabuleiro normal
    sabor = f"TAB: {etiqueta['sabor']}".upper() #usando o f"TAB: defini a palavra tab pra sempre aparecer na etiqueta
    linhas_sabor = quebrar_texto_etiqueta(sabor, 16)

    if len(linhas_sabor) == 1:
        y_sabor = 90
        espacamento = 24
    elif len(linhas_sabor) == 2:
        y_sabor = 78
        espacamento = 22
    else:
        y_sabor = 70
        espacamento = 20

    desenhar_linhas_centralizadas(
        draw,
        ETIQUETA_LARGURA_PX // 2,
        y_sabor,
        linhas_sabor,
        fonte_sabor,
        espacamento=espacamento
    )

    draw.text((18, 185), "Data:", font=fonte_info_titulo, fill="black")
    draw.text((120, 185), etiqueta["data"], font=fonte_info, fill="black")

    draw.text((18, 213), "Validade:", font=fonte_info_titulo, fill="black")
    draw.text((120, 213), etiqueta["validade"], font=fonte_info, fill="black")

    draw.text((18, 241), "Qtd:", font=fonte_info_titulo, fill="black")
    draw.text((120, 241), f'{etiqueta["quantidade_empadas"]} unidades', font=fonte_info, fill="black")

    if "numero" in etiqueta and "total" in etiqueta:
        draw.text(
            (18, 269),
            f'Tab. {etiqueta["numero"]}/{etiqueta["total"]}',
            font=fonte_pequena,
            fill="black"
        )

    return img


def imprimir_etiquetas_direto(etiquetas):
    if not etiquetas:
        raise ValueError("Nenhuma etiqueta para imprimir.")

    if win32ui is None:
        raise RuntimeError("Impressão direta da Argox disponível apenas no Windows com pywin32 instalado.")
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(PRINTER_NAME)

    try:
        hdc.StartDoc("Etiquetas FoodOps Demo")

        for etiqueta in etiquetas:
            img = criar_imagem_etiqueta(etiqueta)

            if ROTACAO_IMPRESSAO != 0:
                img = img.rotate(ROTACAO_IMPRESSAO, expand=True)

            largura_img, altura_img = img.size

            largura_destino = int(largura_img * ESCALA_IMPRESSAO)
            altura_destino = int(altura_img * ESCALA_IMPRESSAO)

            dib = ImageWin.Dib(img)

            hdc.StartPage()

            area_impressao = (
                0,
                0,
                largura_destino,
                altura_destino
            )

            dib.draw(hdc.GetHandleOutput(), area_impressao)

            hdc.EndPage()

        hdc.EndDoc()

    finally:
        hdc.DeleteDC()


def montar_etiquetas_producao(resultado, mistos=None):
    etiquetas_geradas = []

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    validade_controle = (datetime.now() + timedelta(days=7)).strftime("%d/%m/%Y")

    etiquetas_geradas.append({
        "tipo": "controle_qualidade",
        "titulo": "CONTROLE DE QUALIDADE",
        "descricao": "Tabuleiro com todos os sabores",
        "data": data_hoje,
        "validade": validade_controle,
        "dias_validade": 7,
        "quantidade_empadas": EMPADAS_POR_TABULEIRO,
        "numero": 1,
        "total": 1
    })

    for item in resultado:
        quantidade = item["producao"]

        if quantidade > 0:
            for numero in range(1, quantidade + 1):
                validade, dias_validade = calcular_data_validade(item["nome"])

                etiquetas_geradas.append({
                    "tipo": "tabuleiro",
                    "sabor": item["nome"],
                    "classe": item["classe"],
                    "data": data_hoje,
                    "validade": validade,
                    "dias_validade": dias_validade,
                    "quantidade_empadas": EMPADAS_POR_TABULEIRO,
                    "numero": numero,
                    "total": quantidade
                })

    for misto in (mistos or []):
        quantidade = int(float(misto.get("quantidade_tabuleiros", 0) or 0))
        if quantidade <= 0:
            continue
        sabor_1 = misto.get("sabor_1_nome", "").replace("Empada de ", "", 1)
        sabor_2 = misto.get("sabor_2_nome", "").replace("Empada de ", "", 1)
        dias_validade = min(obter_dias_validade(sabor_1), obter_dias_validade(sabor_2))
        validade = (datetime.now() + timedelta(days=dias_validade)).strftime("%d/%m/%Y")
        for numero in range(1, quantidade + 1):
            etiquetas_geradas.append({
                "tipo": "misto",
                "sabor_1": sabor_1,
                "sabor_2": sabor_2,
                "qtd_1": 18,
                "qtd_2": 17,
                "data": data_hoje,
                "validade": validade,
                "dias_validade": dias_validade,
                "quantidade_empadas": EMPADAS_POR_TABULEIRO,
                "numero": numero,
                "total": quantidade,
            })

    return etiquetas_geradas


def montar_etiquetas_avulsas(form):
    tipo_etiqueta = form.get("tipo_etiqueta")
    quantidade = int(form.get("quantidade", 0) or 0)

    if quantidade < 1:
        quantidade = 1

    etiquetas_geradas = []
    data_hoje = datetime.now().strftime("%d/%m/%Y")

    if tipo_etiqueta == "tabuleiro":
        sabor = form.get("sabor_tabuleiro", "").strip()

        validade, dias_validade = calcular_data_validade(sabor)

        for numero in range(1, quantidade + 1):
            etiquetas_geradas.append({
                "tipo": "tabuleiro",
                "sabor": sabor,
                "data": data_hoje,
                "validade": validade,
                "dias_validade": dias_validade,
                "quantidade_empadas": EMPADAS_POR_TABULEIRO,
                "numero": numero,
                "total": quantidade
            })

    elif tipo_etiqueta == "misto":
        sabor_1 = form.get("sabor_1", "").strip()
        sabor_2 = form.get("sabor_2", "").strip()

        dias_1 = obter_dias_validade(sabor_1)
        dias_2 = obter_dias_validade(sabor_2)

        dias_validade = min(dias_1, dias_2)
        validade = (datetime.now() + timedelta(days=dias_validade)).strftime("%d/%m/%Y")

        for numero in range(1, quantidade + 1):
            etiquetas_geradas.append({
                "tipo": "misto",
                "sabor_1": sabor_1,
                "sabor_2": sabor_2,
                "qtd_1": 18,
                "qtd_2": 17,
                "data": data_hoje,
                "validade": validade,
                "dias_validade": dias_validade,
                "quantidade_empadas": EMPADAS_POR_TABULEIRO,
                "numero": numero,
                "total": quantidade
            })

    return etiquetas_geradas


# ============================================================
# BACKUP
# ============================================================

def garantir_pasta_backups():
    if not os.path.exists("backups"):
        os.makedirs("backups")


def criar_backup_banco():
    garantir_pasta_backups()

    data_backup = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_backup = f"backup_database_{data_backup}.db"
    caminho_backup = os.path.join("backups", nome_backup)

    origem = sqlite3.connect(DB_NAME)
    destino = sqlite3.connect(caminho_backup)

    origem.backup(destino)

    destino.close()
    origem.close()

    return nome_backup


def listar_backups():
    garantir_pasta_backups()

    arquivos = []

    for nome_arquivo in os.listdir("backups"):
        if nome_arquivo.endswith(".db"):
            caminho = os.path.join("backups", nome_arquivo)
            tamanho = os.path.getsize(caminho)

            arquivos.append({
                "nome": nome_arquivo,
                "tamanho_kb": round(tamanho / 1024, 2)
            })

    arquivos.sort(key=lambda item: item["nome"], reverse=True)
    return arquivos


# ============================================================
# SOLICITAÇÕES INTERNAS / CENTROS DE CUSTO - V2.7
# ============================================================

def popular_centros_custo_iniciais():
    conn = conectar()
    cursor = conn.cursor()
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    for nome in CENTROS_CUSTO_INICIAIS:
        cursor.execute("""
            INSERT OR IGNORE INTO centros_custo
            (nome, descricao, ativo, data_cadastro, criado_por)
            VALUES (?, '', 1, ?, 'sistema')
        """, (nome, agora))

    conn.commit()
    conn.close()


def carregar_centros_custo(apenas_ativos=True):
    conn = conectar()
    cursor = conn.cursor()

    if apenas_ativos:
        cursor.execute("""
            SELECT *
            FROM centros_custo
            WHERE ativo = 1
            ORDER BY nome
        """)
    else:
        cursor.execute("""
            SELECT *
            FROM centros_custo
            ORDER BY ativo DESC, nome
        """)

    centros = cursor.fetchall()
    conn.close()
    return centros


def buscar_centro_custo(centro_custo_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM centros_custo WHERE id = ?", (centro_custo_id,))
    centro = cursor.fetchone()
    conn.close()
    return centro


def _data_html_para_br(valor):
    valor = str(valor or "").strip()
    if not valor:
        return ""
    try:
        return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return valor


def _data_br_para_objeto(valor):
    valor = str(valor or "").strip()
    if not valor:
        return None
    for formato in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(valor, formato)
        except ValueError:
            continue
    return None


def gerar_codigo_solicitacao_interna(cursor):
    prefixo = datetime.now().strftime("SI-%Y%m%d-")
    cursor.execute("""
        SELECT codigo
        FROM solicitacoes_internas
        WHERE codigo LIKE ?
        ORDER BY id DESC
        LIMIT 1
    """, (f"{prefixo}%",))
    ultimo = cursor.fetchone()

    sequencia = 1
    if ultimo:
        try:
            sequencia = int(str(ultimo["codigo"]).split("-")[-1]) + 1
        except (ValueError, IndexError):
            sequencia = 1

    return f"{prefixo}{sequencia:03d}"


def contar_solicitacoes_internas_pendentes(somente_usuario=None):
    conn = conectar()
    cursor = conn.cursor()
    filtros = ["status IN ('Pendente', 'Em separação', 'Pronta para confirmar')"]
    parametros = []

    if somente_usuario:
        filtros.append("solicitante = ?")
        parametros.append(somente_usuario)

    cursor.execute(f"""
        SELECT COUNT(*) AS total
        FROM solicitacoes_internas
        WHERE {' AND '.join(filtros)}
    """, parametros)
    total = int(cursor.fetchone()["total"] or 0)
    conn.close()
    return total


def criar_solicitacao_interna(form):
    destino_tipo = str(form.get("destino_tipo", "") or "").strip()
    centro_custo_id = str(form.get("centro_custo_id", "") or "").strip()
    celula_id = str(form.get("celula_id", "") or "").strip()
    data_necessidade = _data_html_para_br(form.get("data_necessidade", ""))
    prioridade = str(form.get("prioridade", "Normal") or "Normal").strip()
    observacao = str(form.get("observacao", "") or "").strip()

    if destino_tipo not in TIPOS_DESTINO_SOLICITACAO:
        raise ValueError("Selecione um destino válido para a solicitação.")

    if prioridade not in PRIORIDADES_SOLICITACAO:
        prioridade = "Normal"

    centro = None
    celula = None

    if destino_tipo == "Centro de Custo":
        if not centro_custo_id:
            raise ValueError("Selecione o centro de custo de destino.")
        centro = buscar_centro_custo(int(centro_custo_id))
        if not centro or not int(centro["ativo"]):
            raise ValueError("Centro de custo inválido ou inativo.")
        celula_id = None
    else:
        if not celula_id:
            raise ValueError("Selecione a célula de produção de destino.")
        celula = buscar_celula_producao(int(celula_id))
        if not celula or not int(celula["ativo"]):
            raise ValueError("Célula de produção inválida ou inativa.")
        centro_custo_id = None

    produtos_ids = form.getlist("produto_id[]")
    embalagens_ids = form.getlist("embalagem_id[]")
    quantidades = form.getlist("quantidade[]")
    observacoes_itens = form.getlist("observacao_item[]")

    itens = []
    produtos_usados = set()

    for indice, produto_valor in enumerate(produtos_ids):
        produto_valor = str(produto_valor or "").strip()
        if not produto_valor:
            continue

        produto_id = int(produto_valor)
        if produto_id in produtos_usados:
            raise ValueError("O mesmo produto foi informado mais de uma vez na solicitação.")
        produtos_usados.add(produto_id)

        produto = buscar_produto_estoque(produto_id)
        if not produto or not int(produto["ativo"]):
            raise ValueError("Um dos produtos selecionados não existe ou está inativo.")

        quantidade_texto = quantidades[indice] if indice < len(quantidades) else "0"
        try:
            quantidade_embalagem = float(str(quantidade_texto or 0).replace(",", "."))
        except ValueError:
            raise ValueError(f"Quantidade inválida para {produto['nome']}.")

        if quantidade_embalagem <= 0:
            raise ValueError(f"A quantidade de {produto['nome']} deve ser maior que zero.")

        embalagem_id = None
        embalagem_nome = produto["unidade_padrao"]
        fator_conversao = 1.0
        embalagem_valor = embalagens_ids[indice] if indice < len(embalagens_ids) else ""

        if str(embalagem_valor or "").strip():
            embalagem = buscar_embalagem_produto(int(embalagem_valor), produto_id)
            if not embalagem:
                raise ValueError(f"Embalagem inválida para {produto['nome']}.")
            embalagem_id = embalagem["id"]
            embalagem_nome = embalagem["nome"]
            fator_conversao = float(embalagem["fator_conversao"] or 1)

        quantidade_base = quantidade_embalagem * fator_conversao
        observacao_item = (
            str(observacoes_itens[indice] or "").strip()
            if indice < len(observacoes_itens)
            else ""
        )

        itens.append({
            "produto_id": produto_id,
            "embalagem_id": embalagem_id,
            "embalagem_nome": embalagem_nome,
            "quantidade_embalagem": quantidade_embalagem,
            "fator_conversao": fator_conversao,
            "quantidade_base": quantidade_base,
            "unidade": produto["unidade_padrao"],
            "observacao": observacao_item,
        })

    if not itens:
        raise ValueError("Adicione pelo menos um produto à solicitação.")

    conn = conectar()
    cursor = conn.cursor()

    try:
        codigo = gerar_codigo_solicitacao_interna(cursor)
        agora = _agora_texto()
        usuario = _usuario_atual()

        cursor.execute("""
            INSERT INTO solicitacoes_internas
            (codigo, solicitante, destino_tipo, centro_custo_id, celula_id,
             data_solicitacao, data_necessidade, prioridade, status, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', ?)
        """, (
            codigo,
            usuario,
            destino_tipo,
            int(centro_custo_id) if centro_custo_id else None,
            int(celula_id) if celula_id else None,
            agora,
            data_necessidade,
            prioridade,
            observacao,
        ))
        solicitacao_id = cursor.lastrowid

        for item in itens:
            cursor.execute("""
                INSERT INTO itens_solicitacao_interna
                (solicitacao_id, produto_id, embalagem_id, embalagem_nome,
                 quantidade_solicitada_embalagem, fator_conversao,
                 quantidade_solicitada_base, quantidade_separada_base,
                 quantidade_atendida_base, unidade, status, recusado,
                 observacao, atualizado_em, atualizado_por)
                VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0, ?, 'Pendente', 0, ?, ?, ?)
            """, (
                solicitacao_id,
                item["produto_id"],
                item["embalagem_id"],
                item["embalagem_nome"],
                item["quantidade_embalagem"],
                item["fator_conversao"],
                item["quantidade_base"],
                item["unidade"],
                item["observacao"],
                agora,
                usuario,
            ))

        conn.commit()
        return solicitacao_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def carregar_solicitacoes_internas(status="", destino_tipo="", centro_custo_id="", celula_id="", busca="", somente_usuario=None):
    filtros = []
    parametros = []

    if status:
        filtros.append("si.status = ?")
        parametros.append(status)
    if destino_tipo:
        filtros.append("si.destino_tipo = ?")
        parametros.append(destino_tipo)
    if centro_custo_id:
        filtros.append("si.centro_custo_id = ?")
        parametros.append(centro_custo_id)
    if celula_id:
        filtros.append("si.celula_id = ?")
        parametros.append(celula_id)
    if busca:
        filtros.append("(si.codigo LIKE ? OR si.solicitante LIKE ? OR si.observacao LIKE ?)")
        termo = f"%{busca}%"
        parametros.extend([termo, termo, termo])
    if somente_usuario:
        filtros.append("si.solicitante = ?")
        parametros.append(somente_usuario)

    where = f"WHERE {' AND '.join(filtros)}" if filtros else ""

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT
            si.*,
            cc.nome AS centro_custo_nome,
            cp.nome AS celula_nome,
            COUNT(isi.id) AS total_itens,
            COALESCE(SUM(isi.quantidade_solicitada_base), 0) AS total_solicitado,
            COALESCE(SUM(isi.quantidade_separada_base), 0) AS total_separado,
            COALESCE(SUM(isi.custo_total_snapshot), 0) AS custo_itens
        FROM solicitacoes_internas si
        LEFT JOIN centros_custo cc ON cc.id = si.centro_custo_id
        LEFT JOIN celulas_producao cp ON cp.id = si.celula_id
        LEFT JOIN itens_solicitacao_interna isi ON isi.solicitacao_id = si.id
        {where}
        GROUP BY si.id
        ORDER BY si.id DESC
    """, parametros)
    linhas = cursor.fetchall()
    conn.close()
    return linhas


def buscar_solicitacao_interna(solicitacao_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            si.*,
            cc.nome AS centro_custo_nome,
            cp.nome AS celula_nome
        FROM solicitacoes_internas si
        LEFT JOIN centros_custo cc ON cc.id = si.centro_custo_id
        LEFT JOIN celulas_producao cp ON cp.id = si.celula_id
        WHERE si.id = ?
    """, (solicitacao_id,))
    solicitacao = cursor.fetchone()
    conn.close()
    return solicitacao


def carregar_itens_solicitacao_interna(solicitacao_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            isi.*,
            pe.codigo AS produto_codigo,
            pe.nome AS produto_nome,
            pe.unidade_padrao,
            pe.custo_padrao,
            ce.nome AS categoria_nome
        FROM itens_solicitacao_interna isi
        INNER JOIN produtos_estoque pe ON pe.id = isi.produto_id
        LEFT JOIN categorias_estoque ce ON ce.id = pe.categoria_id
        WHERE isi.solicitacao_id = ?
        ORDER BY isi.id
    """, (solicitacao_id,))

    itens = []
    for row in cursor.fetchall():
        item = dict(row)
        item["saldo_atual"] = obter_saldo_produto(item["produto_id"])
        item["quantidade_faltante"] = max(
            float(item["quantidade_solicitada_base"] or 0) - float(item["quantidade_separada_base"] or 0),
            0,
        )
        itens.append(item)

    conn.close()
    return itens


def _recalcular_status_solicitacao(cursor, solicitacao_id):
    cursor.execute("SELECT status FROM solicitacoes_internas WHERE id = ?", (solicitacao_id,))
    linha_solicitacao = cursor.fetchone()
    if not linha_solicitacao or linha_solicitacao["status"] in ("Atendida", "Atendida parcialmente", "Cancelada"):
        return

    cursor.execute("""
        SELECT status, quantidade_separada_base, recusado
        FROM itens_solicitacao_interna
        WHERE solicitacao_id = ?
    """, (solicitacao_id,))
    itens = cursor.fetchall()

    if not itens:
        novo_status = "Pendente"
    else:
        decidiu_todos = all(
            item["status"] in ("Separado", "Parcial", "Recusado")
            for item in itens
        )
        tem_progresso = any(
            float(item["quantidade_separada_base"] or 0) > 0 or int(item["recusado"] or 0) == 1
            for item in itens
        )

        if decidiu_todos and tem_progresso:
            novo_status = "Pronta para confirmar"
        elif tem_progresso:
            novo_status = "Em separação"
        else:
            novo_status = "Pendente"

    cursor.execute("""
        UPDATE solicitacoes_internas
        SET status = ?, analisado_por = ?, data_analise = ?
        WHERE id = ?
    """, (novo_status, _usuario_atual(), _agora_texto(), solicitacao_id))


def salvar_separacao_solicitacao_interna(solicitacao_id, form):
    solicitacao = buscar_solicitacao_interna(solicitacao_id)
    if not solicitacao:
        raise ValueError("Solicitação interna não encontrada.")
    if solicitacao["status"] in ("Atendida", "Atendida parcialmente", "Cancelada"):
        raise ValueError("Esta solicitação não pode mais ser alterada.")

    itens = carregar_itens_solicitacao_interna(solicitacao_id)
    conn = conectar()
    cursor = conn.cursor()

    try:
        agora = _agora_texto()
        usuario = _usuario_atual()

        for item in itens:
            item_id = item["id"]
            recusado = form.get(f"recusado_{item_id}") == "1"
            motivo_recusa = str(form.get(f"motivo_recusa_{item_id}", "") or "").strip()
            observacao = str(form.get(f"observacao_{item_id}", "") or "").strip()

            try:
                quantidade_separada = float(
                    str(form.get(f"separado_{item_id}", 0) or 0).replace(",", ".")
                )
            except ValueError:
                raise ValueError(f"Quantidade separada inválida para {item['produto_nome']}.")

            if quantidade_separada < 0:
                raise ValueError("Quantidade separada não pode ser negativa.")

            solicitado = float(item["quantidade_solicitada_base"] or 0)
            if quantidade_separada > solicitado + 1e-9:
                raise ValueError(
                    f"A quantidade separada de {item['produto_nome']} não pode superar {solicitado:.4f} {item['unidade']} solicitados."
                )

            if recusado:
                if not motivo_recusa:
                    raise ValueError(f"Informe o motivo da recusa de {item['produto_nome']}.")
                quantidade_separada = 0
                status_item = "Recusado"
            elif quantidade_separada <= 0:
                status_item = "Pendente"
                motivo_recusa = ""
            elif quantidade_separada + 1e-9 < solicitado:
                status_item = "Parcial"
                motivo_recusa = ""
            else:
                quantidade_separada = solicitado
                status_item = "Separado"
                motivo_recusa = ""

            cursor.execute("""
                UPDATE itens_solicitacao_interna
                SET quantidade_separada_base = ?, status = ?, recusado = ?,
                    motivo_recusa = ?, observacao = ?, atualizado_em = ?, atualizado_por = ?
                WHERE id = ? AND solicitacao_id = ?
            """, (
                quantidade_separada,
                status_item,
                1 if recusado else 0,
                motivo_recusa,
                observacao,
                agora,
                usuario,
                item_id,
                solicitacao_id,
            ))

        _recalcular_status_solicitacao(cursor, solicitacao_id)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _saldo_estoque_cursor(cursor, produto_id):
    cursor.execute("""
        SELECT COALESCE(SUM(
            CASE WHEN tipo = 'Entrada' THEN quantidade ELSE -quantidade END
        ), 0) AS saldo
        FROM movimentacoes_estoque
        WHERE produto_id = ?
    """, (produto_id,))
    return float(cursor.fetchone()["saldo"] or 0)


def confirmar_solicitacao_interna(solicitacao_id):
    solicitacao = buscar_solicitacao_interna(solicitacao_id)
    if not solicitacao:
        raise ValueError("Solicitação interna não encontrada.")
    if solicitacao["status"] != "Pronta para confirmar":
        raise ValueError("Conclua a separação de todos os itens antes de confirmar o atendimento.")

    conn = conectar()
    cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT
                isi.*,
                pe.nome AS produto_nome,
                pe.unidade_padrao,
                pe.custo_padrao
            FROM itens_solicitacao_interna isi
            INNER JOIN produtos_estoque pe ON pe.id = isi.produto_id
            WHERE isi.solicitacao_id = ?
            ORDER BY isi.id
        """, (solicitacao_id,))
        itens = cursor.fetchall()

        itens_atendidos = [item for item in itens if float(item["quantidade_separada_base"] or 0) > 0]

        for item in itens_atendidos:
            saldo = _saldo_estoque_cursor(cursor, item["produto_id"])
            quantidade = float(item["quantidade_separada_base"] or 0)
            if saldo + 1e-9 < quantidade:
                raise ValueError(
                    f"Saldo insuficiente de {item['produto_nome']}. Disponível: {saldo:.4f} {item['unidade_padrao']}; separado: {quantidade:.4f} {item['unidade_padrao']}."
                )

        agora = _agora_texto()
        data_mov = datetime.now().strftime("%d/%m/%Y")
        usuario = _usuario_atual()
        custo_total_solicitacao = 0.0
        teve_parcial_ou_recusa = False

        for item in itens:
            quantidade = float(item["quantidade_separada_base"] or 0)
            solicitado = float(item["quantidade_solicitada_base"] or 0)

            if int(item["recusado"] or 0) == 1:
                teve_parcial_ou_recusa = True
                continue
            if quantidade <= 0:
                teve_parcial_ou_recusa = True
                continue

            custo_unitario = float(item["custo_padrao"] or 0)
            custo_total = quantidade * custo_unitario
            custo_total_solicitacao += custo_total

            quantidade_embalagem = quantidade / float(item["fator_conversao"] or 1)
            observacao_mov = (
                f"Solicitação interna {solicitacao['codigo']} para "
                f"{solicitacao['centro_custo_nome'] or solicitacao['celula_nome']}."
            )
            if item["observacao"]:
                observacao_mov += f" {item['observacao']}"

            cursor.execute("""
                INSERT INTO movimentacoes_estoque
                (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_id,
                 embalagem_nome, fator_conversao, data_movimentacao, origem,
                 origem_id, usuario, observacao)
                VALUES (?, 'Saída', ?, ?, ?, ?, ?, ?, 'Solicitação Interna', ?, ?, ?)
            """, (
                item["produto_id"],
                quantidade,
                quantidade_embalagem,
                item["embalagem_id"],
                item["embalagem_nome"],
                item["fator_conversao"],
                data_mov,
                solicitacao_id,
                usuario,
                observacao_mov,
            ))
            estoque_movimentacao_id = cursor.lastrowid
            celula_movimentacao_id = None

            if solicitacao["destino_tipo"] == "Célula":
                cursor.execute("""
                    INSERT INTO movimentacoes_celula
                    (celula_id, produto_id, tipo, quantidade, data_movimentacao,
                     origem, origem_id, usuario, observacao)
                    VALUES (?, ?, 'Entrada', ?, ?, 'Solicitação Interna', ?, ?, ?)
                """, (
                    solicitacao["celula_id"],
                    item["produto_id"],
                    quantidade,
                    data_mov,
                    solicitacao_id,
                    usuario,
                    observacao_mov,
                ))
                celula_movimentacao_id = cursor.lastrowid

            status_item = "Atendido"
            if quantidade + 1e-9 < solicitado:
                status_item = "Atendido parcialmente"
                teve_parcial_ou_recusa = True

            cursor.execute("""
                UPDATE itens_solicitacao_interna
                SET quantidade_atendida_base = ?, status = ?,
                    custo_unitario_snapshot = ?, custo_total_snapshot = ?,
                    estoque_movimentacao_id = ?, celula_movimentacao_id = ?,
                    atualizado_em = ?, atualizado_por = ?
                WHERE id = ?
            """, (
                quantidade,
                status_item,
                custo_unitario,
                custo_total,
                estoque_movimentacao_id,
                celula_movimentacao_id,
                agora,
                usuario,
                item["id"],
            ))

        status_final = "Atendida parcialmente" if teve_parcial_ou_recusa else "Atendida"
        cursor.execute("""
            UPDATE solicitacoes_internas
            SET status = ?, confirmado_por = ?, data_confirmacao = ?, custo_total = ?
            WHERE id = ?
        """, (status_final, usuario, agora, custo_total_solicitacao, solicitacao_id))

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def cancelar_solicitacao_interna(solicitacao_id, motivo=""):
    solicitacao = buscar_solicitacao_interna(solicitacao_id)
    if not solicitacao:
        raise ValueError("Solicitação interna não encontrada.")
    if solicitacao["status"] in ("Atendida", "Atendida parcialmente", "Cancelada"):
        raise ValueError("Esta solicitação não pode ser cancelada.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE solicitacoes_internas
        SET status = 'Cancelada', cancelado_por = ?, data_cancelamento = ?, motivo_cancelamento = ?
        WHERE id = ?
    """, (_usuario_atual(), _agora_texto(), str(motivo or "").strip(), solicitacao_id))
    conn.commit()
    conn.close()


def carregar_relatorio_centros_custo(data_inicio="", data_fim="", centro_custo_id=""):
    inicio = _data_br_para_objeto(data_inicio)
    fim = _data_br_para_objeto(data_fim)
    if fim:
        fim = fim.replace(hour=23, minute=59, second=59)

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            si.id,
            si.codigo,
            si.data_confirmacao,
            si.solicitante,
            si.confirmado_por,
            si.centro_custo_id,
            cc.nome AS centro_custo_nome,
            isi.quantidade_atendida_base,
            isi.unidade,
            isi.custo_unitario_snapshot,
            isi.custo_total_snapshot,
            pe.codigo AS produto_codigo,
            pe.nome AS produto_nome
        FROM itens_solicitacao_interna isi
        INNER JOIN solicitacoes_internas si ON si.id = isi.solicitacao_id
        INNER JOIN centros_custo cc ON cc.id = si.centro_custo_id
        INNER JOIN produtos_estoque pe ON pe.id = isi.produto_id
        WHERE si.destino_tipo = 'Centro de Custo'
          AND si.status IN ('Atendida', 'Atendida parcialmente')
          AND isi.quantidade_atendida_base > 0
        ORDER BY si.id DESC, isi.id
    """)

    detalhes = []
    resumo = {}
    total_geral = 0.0

    for row in cursor.fetchall():
        if centro_custo_id and str(row["centro_custo_id"]) != str(centro_custo_id):
            continue
        data_confirmacao = _data_br_para_objeto(row["data_confirmacao"])
        if inicio and (not data_confirmacao or data_confirmacao < inicio):
            continue
        if fim and (not data_confirmacao or data_confirmacao > fim):
            continue

        item = dict(row)
        detalhes.append(item)
        nome = item["centro_custo_nome"]
        valor = float(item["custo_total_snapshot"] or 0)
        resumo.setdefault(nome, {"centro_custo": nome, "total": 0.0, "itens": 0})
        resumo[nome]["total"] += valor
        resumo[nome]["itens"] += 1
        total_geral += valor

    # Inclui despesas/serviços registrados em notas sem movimentação de estoque.
    cursor.execute("""
        SELECT ne.id, ne.numero AS codigo, ne.data_lancamento AS data_confirmacao,
               ne.criado_por AS solicitante, ne.lancado_por AS confirmado_por,
               ine.centro_custo_id, cc.nome AS centro_custo_nome,
               ine.quantidade_faturada AS quantidade_atendida_base,
               ine.unidade, ine.custo_efetivo_unitario AS custo_unitario_snapshot,
               ine.custo_total_entrada AS custo_total_snapshot,
               '' AS produto_codigo, ine.descricao AS produto_nome
        FROM itens_nota_entrada ine
        INNER JOIN notas_entrada ne ON ne.id = ine.nota_id
        INNER JOIN centros_custo cc ON cc.id = ine.centro_custo_id
        WHERE ne.status = 'Lançada' AND ine.movimenta_estoque = 0
        ORDER BY ne.id DESC, ine.id
    """)
    for row in cursor.fetchall():
        if centro_custo_id and str(row["centro_custo_id"]) != str(centro_custo_id):
            continue
        data_confirmacao = _data_br_para_objeto(row["data_confirmacao"])
        if inicio and (not data_confirmacao or data_confirmacao < inicio):
            continue
        if fim and (not data_confirmacao or data_confirmacao > fim):
            continue
        item = dict(row); detalhes.append(item)
        nome = item["centro_custo_nome"]; valor = float(item["custo_total_snapshot"] or 0)
        resumo.setdefault(nome, {"centro_custo": nome, "total": 0.0, "itens": 0})
        resumo[nome]["total"] += valor; resumo[nome]["itens"] += 1; total_geral += valor

    conn.close()
    resumo_lista = sorted(resumo.values(), key=lambda x: x["total"], reverse=True)
    return resumo_lista, detalhes, total_geral




# ============================================================
# PEDIDOS DAS LOJAS / EXPEDIÇÃO - V2.8
# ============================================================

def popular_origens_expedicao_iniciais():
    conn = conectar()
    cursor = conn.cursor()

    for nome in ORIGENS_EXPEDICAO_INICIAIS:
        cursor.execute("""
            INSERT OR IGNORE INTO origens_expedicao (nome, ativo)
            VALUES (?, 1)
        """, (nome,))

    # Vincula automaticamente produtos sem origem com base na categoria.
    regras = {
        "Produção Própria": "Produção",
        "Recheios": "Produção",
        "Bebidas": "Bebidas",
        "Descartáveis": "Descartáveis e Apoio",
        "Refeitório": "Descartáveis e Apoio",
        "Material de Limpeza": "Limpeza",
        "Material de Escritório": "Escritório",
    }

    for categoria, origem in regras.items():
        cursor.execute("SELECT id FROM origens_expedicao WHERE nome = ?", (origem,))
        origem_row = cursor.fetchone()
        if not origem_row:
            continue
        cursor.execute("""
            UPDATE produtos_estoque
            SET origem_expedicao_id = ?
            WHERE origem_expedicao_id IS NULL
              AND categoria_id IN (SELECT id FROM categorias_estoque WHERE nome = ?)
        """, (origem_row["id"], categoria))

    cursor.execute("SELECT id FROM origens_expedicao WHERE nome = 'Almoxarifado Geral'")
    geral = cursor.fetchone()
    if geral:
        cursor.execute("""
            UPDATE produtos_estoque
            SET origem_expedicao_id = ?
            WHERE origem_expedicao_id IS NULL
        """, (geral["id"],))

    # Para empadas, somente o tabuleiro c/35 fica liberado no catálogo normal.
    cursor.execute("""
        UPDATE produto_embalagens
        SET disponivel_loja = CASE WHEN nome = 'Tabuleiro c/35' THEN 1 ELSE 0 END
        WHERE produto_id IN (
            SELECT id FROM produtos_estoque WHERE nome LIKE 'Empada de %'
        )
    """)

    # Para outros produtos ativos para venda, libera a embalagem padrão.
    cursor.execute("""
        UPDATE produto_embalagens
        SET disponivel_loja = 1
        WHERE produto_id IN (
            SELECT id FROM produtos_estoque
            WHERE ativo_venda = 1 AND nome NOT LIKE 'Empada de %'
        )
          AND padrao = 1
          AND NOT EXISTS (
              SELECT 1 FROM produto_embalagens pe2
              WHERE pe2.produto_id = produto_embalagens.produto_id
                AND pe2.disponivel_loja = 1
          )
    """)

    conn.commit()
    conn.close()


def carregar_origens_expedicao(apenas_ativas=True):
    conn = conectar()
    cursor = conn.cursor()
    where = "WHERE ativo = 1" if apenas_ativas else ""
    cursor.execute(f"""
        SELECT id, nome, ativo, observacao
        FROM origens_expedicao
        {where}
        ORDER BY nome
    """)
    rows = cursor.fetchall()
    conn.close()
    return rows


def carregar_lojas(apenas_ativas=True):
    conn = conectar()
    cursor = conn.cursor()
    where = "WHERE ativo = 1" if apenas_ativas else ""
    cursor.execute(f"""
        SELECT id, codigo, nome, ativo, observacao, data_cadastro
        FROM lojas
        {where}
        ORDER BY nome
    """)
    rows = cursor.fetchall()
    conn.close()
    return rows


def buscar_loja(loja_id):
    if not loja_id:
        return None
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM lojas WHERE id = ?", (loja_id,))
    row = cursor.fetchone()
    conn.close()
    return row


def carregar_catalogo_loja():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            p.id,
            p.codigo,
            p.nome,
            p.unidade_padrao,
            p.categoria_id,
            c.nome AS categoria_nome,
            p.origem_expedicao_id,
            oe.nome AS origem_nome,
            p.ordem_loja
        FROM produtos_estoque p
        LEFT JOIN categorias_estoque c ON c.id = p.categoria_id
        LEFT JOIN origens_expedicao oe ON oe.id = p.origem_expedicao_id
        WHERE p.ativo = 1
          AND p.ativo_venda = 1
          AND EXISTS (
              SELECT 1 FROM produto_embalagens pe
              WHERE pe.produto_id = p.id
                AND pe.ativo = 1
                AND pe.disponivel_loja = 1
          )
        ORDER BY COALESCE(c.nome, 'Sem categoria'), p.ordem_loja, p.nome
    """)
    produtos = [dict(row) for row in cursor.fetchall()]

    cursor.execute("""
        SELECT id, produto_id, nome, fator_conversao, padrao
        FROM produto_embalagens
        WHERE ativo = 1 AND disponivel_loja = 1
        ORDER BY padrao DESC, fator_conversao, nome
    """)
    embalagens = {}
    for row in cursor.fetchall():
        embalagens.setdefault(row["produto_id"], []).append(dict(row))

    cursor.execute("""
        SELECT p.id, p.codigo, p.nome, p.origem_expedicao_id
        FROM produtos_estoque p
        WHERE p.ativo = 1
          AND p.ativo_venda = 1
          AND p.nome LIKE 'Empada de %'
        ORDER BY p.nome
    """)
    sabores_misto = [dict(row) for row in cursor.fetchall()]
    conn.close()

    categorias = {}
    for produto in produtos:
        produto["embalagens"] = embalagens.get(produto["id"], [])
        categoria = produto["categoria_nome"] or "Sem categoria"
        categorias.setdefault(categoria, []).append(produto)

    return categorias, sabores_misto


def buscar_pedido_loja(pedido_id, loja_id=None):
    conn = conectar()
    cursor = conn.cursor()
    sql = """
        SELECT p.*, l.nome AS loja_nome, l.codigo AS loja_codigo
        FROM pedidos_loja p
        JOIN lojas l ON l.id = p.loja_id
        WHERE p.id = ?
    """
    params = [pedido_id]
    if loja_id:
        sql += " AND p.loja_id = ?"
        params.append(loja_id)
    cursor.execute(sql, params)
    pedido = cursor.fetchone()
    conn.close()
    return pedido


def carregar_itens_pedido_loja(pedido_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            i.*,
            p.codigo AS produto_codigo,
            p.nome AS produto_nome,
            p.unidade_padrao,
            s1.codigo AS sabor_1_codigo,
            s1.nome AS sabor_1_nome,
            s2.codigo AS sabor_2_codigo,
            s2.nome AS sabor_2_nome,
            oe.nome AS origem_nome
        FROM itens_pedido_loja i
        LEFT JOIN produtos_estoque p ON p.id = i.produto_id
        LEFT JOIN produtos_estoque s1 ON s1.id = i.sabor_1_id
        LEFT JOIN produtos_estoque s2 ON s2.id = i.sabor_2_id
        LEFT JOIN origens_expedicao oe ON oe.id = i.origem_expedicao_id
        WHERE i.pedido_id = ?
        ORDER BY COALESCE(oe.nome, ''), i.id
    """, (pedido_id,))
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


def carregar_pedidos_loja(loja_id=None, status="", limite=100):
    conn = conectar()
    cursor = conn.cursor()
    condicoes = []
    params = []
    if loja_id:
        condicoes.append("p.loja_id = ?")
        params.append(loja_id)
    if status:
        condicoes.append("p.status = ?")
        params.append(status)
    where = "WHERE " + " AND ".join(condicoes) if condicoes else ""
    params.append(int(limite))
    cursor.execute(f"""
        SELECT
            p.*,
            l.nome AS loja_nome,
            COUNT(i.id) AS total_itens,
            SUM(CASE WHEN i.status_expedicao = 'Separado' THEN 1 ELSE 0 END) AS itens_separados
        FROM pedidos_loja p
        JOIN lojas l ON l.id = p.loja_id
        LEFT JOIN itens_pedido_loja i ON i.pedido_id = p.id
        {where}
        GROUP BY p.id
        ORDER BY p.id DESC
        LIMIT ?
    """, params)
    rows = cursor.fetchall()
    conn.close()
    return rows


def contar_pedidos_loja_pendentes(loja_id=None):
    conn = conectar()
    cursor = conn.cursor()
    sql = "SELECT COUNT(*) AS total FROM pedidos_loja WHERE status NOT IN ('Finalizado', 'Cancelado')"
    params = []
    if loja_id:
        sql += " AND loja_id = ?"
        params.append(loja_id)
    cursor.execute(sql, params)
    total = cursor.fetchone()["total"]
    conn.close()
    return int(total or 0)


def atualizar_status_pedido_loja(cursor, pedido_id):
    cursor.execute("SELECT status FROM pedidos_loja WHERE id = ?", (pedido_id,))
    pedido = cursor.fetchone()
    if not pedido or pedido["status"] in ["Finalizado", "Cancelado"]:
        return

    cursor.execute("""
        SELECT status_expedicao
        FROM itens_pedido_loja
        WHERE pedido_id = ?
    """, (pedido_id,))
    statuses = [row["status_expedicao"] for row in cursor.fetchall()]
    if not statuses or all(s == "Pendente" for s in statuses):
        novo = "Recebido"
    elif all(s in ["Separado", "Indisponível"] for s in statuses):
        novo = "Em expedição"
    else:
        novo = "Em separação"
    cursor.execute("UPDATE pedidos_loja SET status = ? WHERE id = ?", (novo, pedido_id))


def criar_pedido_loja(loja_id, data_entrega, observacao, produtos_ids, embalagens_ids, quantidades,
                       misto_sabor_1=None, misto_sabor_2=None, misto_quantidade=0):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id FROM lojas WHERE id = ? AND ativo = 1", (loja_id,))
        if not cursor.fetchone():
            raise ValueError("Loja inválida ou inativa.")

        itens_normais = []
        for produto_id, embalagem_id, quantidade in zip(produtos_ids, embalagens_ids, quantidades):
            try:
                quantidade = float(str(quantidade or 0).replace(',', '.'))
                produto_id = int(produto_id)
                embalagem_id = int(embalagem_id)
            except Exception:
                continue
            if quantidade <= 0:
                continue
            cursor.execute("""
                SELECT
                    p.id AS produto_id, p.nome, p.origem_expedicao_id,
                    pe.id AS embalagem_id, pe.nome AS embalagem_nome, pe.fator_conversao
                FROM produtos_estoque p
                JOIN produto_embalagens pe ON pe.produto_id = p.id
                WHERE p.id = ? AND pe.id = ?
                  AND p.ativo = 1 AND p.ativo_venda = 1
                  AND pe.ativo = 1 AND pe.disponivel_loja = 1
            """, (produto_id, embalagem_id))
            row = cursor.fetchone()
            if not row:
                raise ValueError("Um item ou embalagem selecionada não está disponível para a loja.")
            itens_normais.append((row, quantidade))

        try:
            misto_quantidade = float(str(misto_quantidade or 0).replace(',', '.'))
        except Exception:
            misto_quantidade = 0

        item_misto = None
        if misto_quantidade > 0:
            if not misto_sabor_1 or not misto_sabor_2:
                raise ValueError("Selecione os dois sabores do tabuleiro misto.")
            if int(misto_sabor_1) == int(misto_sabor_2):
                raise ValueError("Os sabores do tabuleiro misto precisam ser diferentes.")
            cursor.execute("""
                SELECT id, nome, origem_expedicao_id
                FROM produtos_estoque
                WHERE id IN (?, ?)
                  AND ativo = 1 AND ativo_venda = 1
                  AND nome LIKE 'Empada de %'
                ORDER BY id
            """, (int(misto_sabor_1), int(misto_sabor_2)))
            sabores = {row["id"]: row for row in cursor.fetchall()}
            if int(misto_sabor_1) not in sabores or int(misto_sabor_2) not in sabores:
                raise ValueError("Um dos sabores do misto não está disponível.")
            origem_1 = sabores[int(misto_sabor_1)]["origem_expedicao_id"]
            origem_2 = sabores[int(misto_sabor_2)]["origem_expedicao_id"]
            if origem_1 != origem_2:
                raise ValueError("Os sabores do misto precisam pertencer à mesma origem de expedição.")
            item_misto = (sabores, int(misto_sabor_1), int(misto_sabor_2), misto_quantidade, origem_1)

        if not itens_normais and not item_misto:
            raise ValueError("Informe ao menos um item com quantidade maior que zero.")

        rodada = buscar_ou_criar_rodada(data_entrega, cursor)
        if rodada["status"] != "Aberta":
            raise ValueError("Os pedidos desta data já foram fechados para a produção. Escolha outra data de entrega.")

        cursor.execute("""
            INSERT INTO pedidos_loja
            (loja_id, criado_por, data_criacao, data_entrega_desejada, status, observacao, rodada_id)
            VALUES (?, ?, ?, ?, 'Recebido', ?, ?)
        """, (
            loja_id,
            session.get("usuario", "loja"),
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            data_entrega or None,
            observacao.strip(),
            rodada["id"],
        ))
        pedido_id = cursor.lastrowid

        for row, quantidade in itens_normais:
            cursor.execute("""
                INSERT INTO itens_pedido_loja
                (pedido_id, tipo_item, produto_id, embalagem_id, embalagem_nome, fator_conversao,
                 quantidade_comercial, quantidade_base, origem_expedicao_id, status_expedicao)
                VALUES (?, 'Normal', ?, ?, ?, ?, ?, ?, ?, 'Pendente')
            """, (
                pedido_id, row["produto_id"], row["embalagem_id"], row["embalagem_nome"],
                row["fator_conversao"], quantidade, quantidade * float(row["fator_conversao"] or 1),
                row["origem_expedicao_id"],
            ))

        if item_misto:
            sabores, sabor_1, sabor_2, quantidade, origem_id = item_misto
            cursor.execute("""
                INSERT INTO itens_pedido_loja
                (pedido_id, tipo_item, embalagem_nome, fator_conversao, quantidade_comercial, quantidade_base,
                 sabor_1_id, sabor_2_id, quantidade_sabor_1_base, quantidade_sabor_2_base,
                 origem_expedicao_id, status_expedicao)
                VALUES (?, 'Misto', 'Tabuleiro Misto 18/17', 35, ?, ?, ?, ?, ?, ?, ?, 'Pendente')
            """, (
                pedido_id, quantidade, quantidade * 35,
                sabor_1, sabor_2, quantidade * 18, quantidade * 17, origem_id,
            ))

        conn.commit()
        return pedido_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def registrar_separacao_item_loja(item_id, quantidade_comercial, indisponivel=False, observacao=""):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT * FROM itens_pedido_loja WHERE id = ?", (item_id,))
        item = cursor.fetchone()
        if not item:
            raise ValueError("Item do pedido não encontrado.")
        cursor.execute("SELECT status FROM pedidos_loja WHERE id = ?", (item["pedido_id"],))
        pedido = cursor.fetchone()
        if not pedido or pedido["status"] in ["Finalizado", "Cancelado", "Em rota"]:
            raise ValueError("Este pedido não pode mais ser alterado.")
        cursor.execute("SELECT id, status FROM romaneios_loja WHERE pedido_id = ?", (item["pedido_id"],))
        romaneio_existente = cursor.fetchone()
        if romaneio_existente:
            raise ValueError("O romaneio já foi preparado. Cancele a preparação do romaneio antes de alterar a separação.")

        if indisponivel:
            quantidade = 0.0
            status = "Indisponível"
        else:
            try:
                quantidade = float(str(quantidade_comercial or 0).replace(',', '.'))
            except Exception:
                quantidade = 0.0
            if quantidade < 0 or quantidade > float(item["quantidade_comercial"] or 0):
                raise ValueError("A quantidade separada deve ficar entre zero e a quantidade pedida.")
            if quantidade == 0:
                status = "Pendente"
            elif quantidade < float(item["quantidade_comercial"] or 0):
                status = "Parcial"
            else:
                status = "Separado"

        fator = float(item["fator_conversao"] or 1)
        quantidade_base = quantidade * fator
        cursor.execute("""
            UPDATE itens_pedido_loja
            SET quantidade_separada_comercial = ?,
                quantidade_separada_base = ?,
                status_expedicao = ?,
                observacao_expedicao = ?,
                data_separacao = ?,
                separado_por = ?
            WHERE id = ?
        """, (
            quantidade, quantidade_base, status, observacao.strip(),
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"), session.get("usuario", "sistema"), item_id,
        ))
        atualizar_status_pedido_loja(cursor, item["pedido_id"])
        conn.commit()
        return item["pedido_id"]
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def registrar_separacao_pedido_loja(pedido_id, form):
    """Salva toda a separação do pedido em uma única transação."""
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("SELECT * FROM pedidos_loja WHERE id = ?", (pedido_id,))
        pedido = cursor.fetchone()
        if not pedido:
            raise ValueError("Pedido não encontrado.")
        if pedido["status"] in ["Finalizado", "Cancelado", "Em rota"]:
            raise ValueError("Este pedido não pode mais ser alterado.")

        cursor.execute("SELECT id FROM romaneios_loja WHERE pedido_id = ?", (pedido_id,))
        if cursor.fetchone():
            raise ValueError("O romaneio já foi preparado. Volte o romaneio para a separação antes de alterar a lista.")

        cursor.execute("SELECT * FROM itens_pedido_loja WHERE pedido_id = ? ORDER BY id", (pedido_id,))
        itens = cursor.fetchall()
        if not itens:
            raise ValueError("O pedido não possui itens para separar.")

        agora = _agora_texto()
        usuario = _usuario_atual()
        for item in itens:
            indisponivel = str(form.get(f"indisponivel_{item['id']}", "")) == "1"
            observacao = str(form.get(f"observacao_{item['id']}", item["observacao_expedicao"] or "") or "").strip()
            if indisponivel:
                quantidade = 0.0
                status = "Indisponível"
            else:
                bruto = str(form.get(f"quantidade_separada_{item['id']}", item["quantidade_separada_comercial"] or 0) or "0").replace(",", ".")
                try:
                    quantidade = float(bruto)
                except ValueError:
                    raise ValueError(f"Quantidade inválida no item #{item['id']}.")
                maximo = float(item["quantidade_comercial"] or 0)
                if quantidade < 0 or quantidade > maximo + 1e-9:
                    raise ValueError(f"A quantidade separada do item #{item['id']} deve ficar entre 0 e {maximo:g}.")
                if quantidade <= 0:
                    status = "Pendente"
                elif quantidade < maximo - 1e-9:
                    status = "Parcial"
                else:
                    status = "Separado"

            fator = float(item["fator_conversao"] or 1)
            cursor.execute("""
                UPDATE itens_pedido_loja
                SET quantidade_separada_comercial = ?,
                    quantidade_separada_base = ?,
                    status_expedicao = ?,
                    observacao_expedicao = ?,
                    data_separacao = ?,
                    separado_por = ?
                WHERE id = ?
            """, (quantidade, quantidade * fator, status, observacao, agora, usuario, item["id"]))

        atualizar_status_pedido_loja(cursor, pedido_id)
        conn.commit()
        return pedido_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def finalizar_pedido_loja(pedido_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT * FROM pedidos_loja WHERE id = ?", (pedido_id,))
        pedido = cursor.fetchone()
        if not pedido:
            raise ValueError("Pedido não encontrado.")
        if pedido["status"] != "Em expedição":
            raise ValueError("Todos os itens precisam estar separados ou marcados como indisponíveis antes da finalização.")

        cursor.execute("SELECT * FROM itens_pedido_loja WHERE pedido_id = ? ORDER BY id", (pedido_id,))
        itens = cursor.fetchall()

        # Valida todos os saldos antes de gerar qualquer movimento.
        necessidades = {}
        for item in itens:
            if item["status_expedicao"] == "Indisponível":
                continue
            qtd_comercial = float(item["quantidade_separada_comercial"] or 0)
            if qtd_comercial <= 0:
                continue
            if item["tipo_item"] == "Misto":
                necessidades[item["sabor_1_id"]] = necessidades.get(item["sabor_1_id"], 0) + qtd_comercial * 18
                necessidades[item["sabor_2_id"]] = necessidades.get(item["sabor_2_id"], 0) + qtd_comercial * 17
            else:
                necessidades[item["produto_id"]] = necessidades.get(item["produto_id"], 0) + float(item["quantidade_separada_base"] or 0)

        for produto_id, quantidade in necessidades.items():
            saldo = obter_saldo_produto(produto_id)
            if quantidade > saldo + 1e-9:
                produto = buscar_produto_estoque(produto_id)
                raise ValueError(f"Saldo insuficiente de {produto['nome']}: disponível {saldo:.2f}, necessário {quantidade:.2f}.")

        data_mov = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        for item in itens:
            if item["status_expedicao"] == "Indisponível":
                continue
            qtd_comercial = float(item["quantidade_separada_comercial"] or 0)
            if qtd_comercial <= 0:
                continue
            if item["tipo_item"] == "Misto":
                ids_mov = []
                for produto_id, qtd_un in [(item["sabor_1_id"], qtd_comercial * 18), (item["sabor_2_id"], qtd_comercial * 17)]:
                    cursor.execute("""
                        INSERT INTO movimentacoes_estoque
                        (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_nome, fator_conversao,
                         data_movimentacao, origem, origem_id, usuario, observacao)
                        VALUES (?, 'Saída', ?, ?, 'Tabuleiro Misto', 1, ?, 'Expedição Loja', ?, ?, ?)
                    """, (
                        produto_id, qtd_un, qtd_un, data_mov, pedido_id,
                        session.get("usuario", "sistema"), f"Pedido da loja #{pedido_id} - parte de tabuleiro misto.",
                    ))
                    ids_mov.append(cursor.lastrowid)
                cursor.execute("""
                    UPDATE itens_pedido_loja
                    SET movimento_estoque_id = ?, movimento_estoque_2_id = ?
                    WHERE id = ?
                """, (ids_mov[0], ids_mov[1], item["id"]))
            else:
                cursor.execute("""
                    INSERT INTO movimentacoes_estoque
                    (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_id, embalagem_nome,
                     fator_conversao, data_movimentacao, origem, origem_id, usuario, observacao)
                    VALUES (?, 'Saída', ?, ?, ?, ?, ?, ?, 'Expedição Loja', ?, ?, ?)
                """, (
                    item["produto_id"], item["quantidade_separada_base"], item["quantidade_separada_comercial"],
                    item["embalagem_id"], item["embalagem_nome"], item["fator_conversao"], data_mov,
                    pedido_id, session.get("usuario", "sistema"), f"Pedido da loja #{pedido_id}.",
                ))
                cursor.execute("UPDATE itens_pedido_loja SET movimento_estoque_id = ? WHERE id = ?", (cursor.lastrowid, item["id"]))

        cursor.execute("""
            UPDATE pedidos_loja
            SET status = 'Finalizado', data_finalizacao = ?, finalizado_por = ?
            WHERE id = ?
        """, (data_mov, session.get("usuario", "sistema"), pedido_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def carregar_pedidos_expedicao(origem_id=None, status=""):
    conn = conectar()
    cursor = conn.cursor()
    condicoes = ["p.status NOT IN ('Cancelado')"]
    params = []
    if origem_id:
        condicoes.append("i.origem_expedicao_id = ?")
        params.append(origem_id)
    if status:
        condicoes.append("p.status = ?")
        params.append(status)
    where = " AND ".join(condicoes)
    cursor.execute(f"""
        SELECT
            p.id, p.status, p.data_criacao, p.data_entrega_desejada,
            l.nome AS loja_nome,
            COUNT(i.id) AS total_itens,
            SUM(CASE WHEN i.status_expedicao = 'Separado' THEN 1 ELSE 0 END) AS separados,
            SUM(CASE WHEN i.status_expedicao = 'Indisponível' THEN 1 ELSE 0 END) AS indisponiveis,
            r.id AS romaneio_id,
            r.codigo AS romaneio_codigo,
            r.status AS romaneio_status
        FROM pedidos_loja p
        JOIN lojas l ON l.id = p.loja_id
        JOIN itens_pedido_loja i ON i.pedido_id = p.id
        LEFT JOIN romaneios_loja r ON r.pedido_id = p.id
        WHERE {where}
        GROUP BY p.id
        ORDER BY CASE p.status WHEN 'Recebido' THEN 1 WHEN 'Em separação' THEN 2 WHEN 'Em expedição' THEN 3 WHEN 'Em rota' THEN 4 ELSE 5 END,
                 p.id DESC
    """, params)
    rows = cursor.fetchall()
    conn.close()
    return rows


# ============================================================
# FECHAMENTO DE PEDIDOS / DEMANDA CONSOLIDADA - V2.9
# ============================================================

def _data_iso_valida(valor):
    try:
        return datetime.strptime(str(valor or ""), "%Y-%m-%d").strftime("%Y-%m-%d")
    except Exception:
        return None


def migrar_pedidos_existentes_para_rodadas():
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT id, data_entrega_desejada
            FROM pedidos_loja
            WHERE rodada_id IS NULL
              AND data_entrega_desejada IS NOT NULL
              AND data_entrega_desejada != ''
            ORDER BY id
        """)
        pedidos = cursor.fetchall()
        for pedido in pedidos:
            data_entrega = _data_iso_valida(pedido["data_entrega_desejada"])
            if not data_entrega:
                continue
            rodada = buscar_ou_criar_rodada(data_entrega, cursor)
            cursor.execute("UPDATE pedidos_loja SET rodada_id = ? WHERE id = ?", (rodada["id"], pedido["id"]))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def buscar_ou_criar_rodada(data_entrega, cursor=None):
    data_entrega = _data_iso_valida(data_entrega)
    if not data_entrega:
        raise ValueError("Informe uma data de entrega válida.")
    proprio = cursor is None
    conn = None
    if proprio:
        conn = conectar()
        cursor = conn.cursor()
    cursor.execute("SELECT * FROM rodadas_pedidos_loja WHERE data_entrega = ?", (data_entrega,))
    rodada = cursor.fetchone()
    if not rodada:
        cursor.execute("""
            INSERT INTO rodadas_pedidos_loja
            (data_entrega, status, data_criacao, criado_por, observacao)
            VALUES (?, 'Aberta', ?, ?, ?)
        """, (data_entrega, _agora_texto(), _usuario_atual(), "Rodada criada automaticamente pelo primeiro pedido da data."))
        rodada_id = cursor.lastrowid
        cursor.execute("SELECT * FROM rodadas_pedidos_loja WHERE id = ?", (rodada_id,))
        rodada = cursor.fetchone()
        if proprio:
            conn.commit()
    if proprio:
        conn.close()
    return rodada


def buscar_rodada_pedidos(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT r.*,
               COUNT(DISTINCT p.id) AS total_pedidos,
               COUNT(DISTINCT CASE WHEN p.status != 'Cancelado' THEN p.id END) AS pedidos_validos,
               COUNT(DISTINCT p.loja_id) AS total_lojas
        FROM rodadas_pedidos_loja r
        LEFT JOIN pedidos_loja p ON p.rodada_id = r.id
        WHERE r.id = ?
        GROUP BY r.id
    """, (rodada_id,))
    row = cursor.fetchone()
    conn.close()
    return row


def carregar_rodadas_pedidos(status=""):
    conn = conectar()
    cursor = conn.cursor()
    where = "WHERE r.status = ?" if status else ""
    params = [status] if status else []
    cursor.execute(f"""
        SELECT r.*,
               COUNT(DISTINCT CASE WHEN p.status != 'Cancelado' THEN p.id END) AS pedidos_validos,
               COUNT(DISTINCT CASE WHEN p.status != 'Cancelado' THEN p.loja_id END) AS total_lojas,
               COALESCE(SUM(CASE WHEN p.status != 'Cancelado' THEN 1 ELSE 0 END), 0) AS registros_pedidos
        FROM rodadas_pedidos_loja r
        LEFT JOIN pedidos_loja p ON p.rodada_id = r.id
        {where}
        GROUP BY r.id
        ORDER BY r.data_entrega DESC, r.id DESC
    """, params)
    rows = cursor.fetchall()
    conn.close()
    return rows


def contar_rodadas_abertas():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) AS total FROM rodadas_pedidos_loja WHERE status = 'Aberta'")
    total = cursor.fetchone()["total"]
    conn.close()
    return int(total or 0)


def carregar_demanda_rodada(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT d.*, p.codigo, p.nome, p.unidade_padrao, p.forma_abastecimento,
               c.nome AS categoria_nome
        FROM demanda_producao_rodada d
        INNER JOIN produtos_estoque p ON p.id = d.produto_id
        LEFT JOIN categorias_estoque c ON c.id = p.categoria_id
        WHERE d.rodada_id = ?
        ORDER BY p.nome
    """, (rodada_id,))
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


def carregar_mistos_rodada(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.*, l.nome AS loja_nome,
               s1.nome AS sabor_1_nome, s2.nome AS sabor_2_nome
        FROM mistos_rodada m
        INNER JOIN lojas l ON l.id = m.loja_id
        INNER JOIN produtos_estoque s1 ON s1.id = m.sabor_1_id
        INNER JOIN produtos_estoque s2 ON s2.id = m.sabor_2_id
        WHERE m.rodada_id = ?
        ORDER BY l.nome, m.id
    """, (rodada_id,))
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


def carregar_pedidos_rodada(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT p.*, l.nome AS loja_nome,
               COUNT(i.id) AS total_itens
        FROM pedidos_loja p
        INNER JOIN lojas l ON l.id = p.loja_id
        LEFT JOIN itens_pedido_loja i ON i.pedido_id = p.id
        WHERE p.rodada_id = ?
        GROUP BY p.id
        ORDER BY l.nome, p.id
    """, (rodada_id,))
    rows = cursor.fetchall()
    conn.close()
    return rows


def _produto_produzido_internamente(cursor, produto_id):
    cursor.execute("SELECT forma_abastecimento FROM produtos_estoque WHERE id = ?", (produto_id,))
    row = cursor.fetchone()
    return bool(row and row["forma_abastecimento"] == "Produzido internamente")


def fechar_rodada_pedidos(rodada_id, cenario, dia_semana, observacao=""):
    if cenario not in ["Chuva", "Normal", "Verão", "Baixa"]:
        raise ValueError("Cenário inválido.")
    if dia_semana not in DIAS_SEMANA:
        raise ValueError("Dia da semana inválido.")
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN IMMEDIATE")
        cursor.execute("SELECT * FROM rodadas_pedidos_loja WHERE id = ?", (rodada_id,))
        rodada = cursor.fetchone()
        if not rodada:
            raise ValueError("Rodada de pedidos não encontrada.")
        if rodada["status"] != "Aberta":
            raise ValueError("Apenas rodadas abertas podem ser fechadas.")

        cursor.execute("""
            SELECT p.id
            FROM pedidos_loja p
            WHERE p.rodada_id = ? AND p.status != 'Cancelado'
        """, (rodada_id,))
        pedidos = [row["id"] for row in cursor.fetchall()]
        if not pedidos:
            raise ValueError("Não há pedidos válidos para fechar nesta data.")

        cursor.execute("DELETE FROM demanda_producao_rodada WHERE rodada_id = ?", (rodada_id,))
        cursor.execute("DELETE FROM mistos_rodada WHERE rodada_id = ?", (rodada_id,))

        demanda = {}
        placeholders = ",".join("?" for _ in pedidos)
        cursor.execute(f"""
            SELECT i.*, p.loja_id
            FROM itens_pedido_loja i
            INNER JOIN pedidos_loja p ON p.id = i.pedido_id
            WHERE i.pedido_id IN ({placeholders})
            ORDER BY i.id
        """, pedidos)
        itens = cursor.fetchall()

        for item in itens:
            if item["tipo_item"] == "Misto":
                q1 = float(item["quantidade_sabor_1_base"] or 0)
                q2 = float(item["quantidade_sabor_2_base"] or 0)
                for produto_id, qtd in ((item["sabor_1_id"], q1), (item["sabor_2_id"], q2)):
                    if not produto_id or not _produto_produzido_internamente(cursor, produto_id):
                        continue
                    reg = demanda.setdefault(int(produto_id), {"normal": 0.0, "misto": 0.0})
                    reg["misto"] += qtd
                cursor.execute("""
                    INSERT INTO mistos_rodada
                    (rodada_id, pedido_id, item_pedido_id, loja_id, sabor_1_id, sabor_2_id,
                     quantidade_tabuleiros, quantidade_sabor_1, quantidade_sabor_2)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    rodada_id, item["pedido_id"], item["id"], item["loja_id"],
                    item["sabor_1_id"], item["sabor_2_id"], float(item["quantidade_comercial"] or 0), q1, q2
                ))
            elif item["produto_id"] and _produto_produzido_internamente(cursor, item["produto_id"]):
                reg = demanda.setdefault(int(item["produto_id"]), {"normal": 0.0, "misto": 0.0})
                reg["normal"] += float(item["quantidade_base"] or 0)

        for produto_id, valores in demanda.items():
            normal = valores["normal"]
            misto = valores["misto"]
            cursor.execute("SELECT nome, unidade_padrao FROM produtos_estoque WHERE id = ?", (produto_id,))
            produto = cursor.fetchone()
            obs = None
            tabs = normal / EMPADAS_POR_TABULEIRO if produto and produto["unidade_padrao"] == "un" and str(produto["nome"]).startswith("Empada de ") else 0
            if produto and str(produto["nome"]).startswith("Empada de ") and abs(tabs - round(tabs)) > 1e-9:
                obs = "Demanda normal não fecha tabuleiro completo; revisar embalagem do pedido."
            cursor.execute("""
                INSERT INTO demanda_producao_rodada
                (rodada_id, produto_id, quantidade_normal_unidades, quantidade_misto_unidades,
                 quantidade_total_unidades, tabuleiros_normais, observacao)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (rodada_id, produto_id, normal, misto, normal + misto, tabs, obs))

        cursor.execute("""
            UPDATE rodadas_pedidos_loja
            SET status = 'Fechada', data_fechamento = ?, fechado_por = ?,
                cenario = ?, dia_semana = ?, observacao = ?
            WHERE id = ?
        """, (_agora_texto(), _usuario_atual(), cenario, dia_semana, observacao.strip(), rodada_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def reabrir_rodada_pedidos(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN IMMEDIATE")
        cursor.execute("SELECT * FROM rodadas_pedidos_loja WHERE id = ?", (rodada_id,))
        rodada = cursor.fetchone()
        if not rodada:
            raise ValueError("Rodada não encontrada.")
        if rodada["status"] != "Fechada":
            raise ValueError("Apenas rodadas fechadas e ainda não planejadas podem ser reabertas.")
        if rodada["planejamento_id"]:
            raise ValueError("A rodada já possui planejamento e não pode ser reaberta.")
        cursor.execute("DELETE FROM demanda_producao_rodada WHERE rodada_id = ?", (rodada_id,))
        cursor.execute("DELETE FROM mistos_rodada WHERE rodada_id = ?", (rodada_id,))
        cursor.execute("""
            UPDATE rodadas_pedidos_loja
            SET status = 'Aberta', data_reabertura = ?, reaberto_por = ?,
                data_fechamento = NULL, fechado_por = NULL
            WHERE id = ?
        """, (_agora_texto(), _usuario_atual(), rodada_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def demanda_sabores_rodada(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT d.*, p.nome
        FROM demanda_producao_rodada d
        INNER JOIN produtos_estoque p ON p.id = d.produto_id
        WHERE d.rodada_id = ? AND p.nome LIKE 'Empada de %'
    """, (rodada_id,))
    mapa = {}
    for row in cursor.fetchall():
        sabor = row["nome"][len("Empada de "):]
        mapa[sabor] = {
            "normal_unidades": float(row["quantidade_normal_unidades"] or 0),
            "misto_unidades": float(row["quantidade_misto_unidades"] or 0),
            "total_unidades": float(row["quantidade_total_unidades"] or 0),
            "tabuleiros_normais": float(row["tabuleiros_normais"] or 0),
        }
    conn.close()
    return mapa


def total_mistos_rodada(rodada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COALESCE(SUM(quantidade_tabuleiros), 0) AS total FROM mistos_rodada WHERE rodada_id = ?", (rodada_id,))
    total = float(cursor.fetchone()["total"] or 0)
    conn.close()
    return total


def gerar_romaneio_rodada(rodada_id):
    rodada = buscar_rodada_pedidos(rodada_id)
    if not rodada:
        return None, []
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT p.id AS pedido_id, l.nome AS loja_nome, l.codigo AS loja_codigo,
               i.*, pr.nome AS produto_nome, s1.nome AS sabor_1_nome, s2.nome AS sabor_2_nome,
               oe.nome AS origem_nome
        FROM pedidos_loja p
        INNER JOIN lojas l ON l.id = p.loja_id
        INNER JOIN itens_pedido_loja i ON i.pedido_id = p.id
        LEFT JOIN produtos_estoque pr ON pr.id = i.produto_id
        LEFT JOIN produtos_estoque s1 ON s1.id = i.sabor_1_id
        LEFT JOIN produtos_estoque s2 ON s2.id = i.sabor_2_id
        LEFT JOIN origens_expedicao oe ON oe.id = i.origem_expedicao_id
        WHERE p.rodada_id = ? AND p.status != 'Cancelado'
        ORDER BY l.nome, COALESCE(oe.nome, ''), i.id
    """, (rodada_id,))
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    grupos = {}
    for row in rows:
        grupos.setdefault(row["loja_nome"], []).append(row)
    return rodada, grupos


# ============================================================
# MOTORISTAS E VEÍCULOS - V3.0.1
# ============================================================

def carregar_motoristas(apenas_ativos=True):
    conn = conectar()
    cursor = conn.cursor()
    sql = "SELECT * FROM motoristas"
    if apenas_ativos:
        sql += " WHERE ativo = 1"
    sql += " ORDER BY nome"
    cursor.execute(sql)
    rows = cursor.fetchall()
    conn.close()
    return rows


def buscar_motorista(motorista_id):
    if not motorista_id:
        return None
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM motoristas WHERE id = ?", (motorista_id,))
    row = cursor.fetchone()
    conn.close()
    return row


def carregar_veiculos(apenas_ativos=True):
    conn = conectar()
    cursor = conn.cursor()
    sql = "SELECT * FROM veiculos"
    if apenas_ativos:
        sql += " WHERE ativo = 1"
    sql += " ORDER BY descricao, placa"
    cursor.execute(sql)
    rows = cursor.fetchall()
    conn.close()
    return rows


def buscar_veiculo(veiculo_id):
    if not veiculo_id:
        return None
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM veiculos WHERE id = ?", (veiculo_id,))
    row = cursor.fetchone()
    conn.close()
    return row


# ============================================================
# ROMANEIO / CONFERÊNCIA / FEFO - V3.0.1
# ============================================================

def _codigo_romaneio(cursor, pedido):
    data_base = _data_para_iso(pedido["data_entrega_desejada"] or datetime.now().strftime("%Y-%m-%d"), padrao_hoje=True)
    data_codigo = datetime.strptime(data_base, "%Y-%m-%d").strftime("%Y%m%d")
    cursor.execute("SELECT codigo FROM lojas WHERE id = ?", (pedido["loja_id"],))
    loja = cursor.fetchone()
    codigo_loja = str(loja["codigo"] if loja else pedido["loja_id"]).upper().replace(" ", "-")
    return f"ROM-{data_codigo}-{codigo_loja}-{int(pedido['id']):04d}"


def buscar_romaneio(romaneio_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT r.*, l.nome AS loja_nome, l.codigo AS loja_codigo,
               p.data_entrega_desejada, p.observacao AS pedido_observacao,
               rd.data_entrega AS rodada_data_entrega
        FROM romaneios_loja r
        INNER JOIN lojas l ON l.id = r.loja_id
        INNER JOIN pedidos_loja p ON p.id = r.pedido_id
        LEFT JOIN rodadas_pedidos_loja rd ON rd.id = r.rodada_id
        WHERE r.id = ?
    """, (romaneio_id,))
    row = cursor.fetchone()
    conn.close()
    return row


def buscar_romaneio_por_pedido(pedido_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM romaneios_loja WHERE pedido_id = ?", (pedido_id,))
    row = cursor.fetchone()
    conn.close()
    return buscar_romaneio(row["id"]) if row else None


def _descricao_item_romaneio(item):
    if item["tipo_item"] == "Misto":
        return f"Tabuleiro Misto — {item['sabor_1_nome']} 18 + {item['sabor_2_nome']} 17"
    return item["produto_nome"] or "Produto não identificado"


def preparar_romaneio_pedido(pedido_id):
    pedido = buscar_pedido_loja(pedido_id)
    if not pedido:
        raise ValueError("Pedido não encontrado.")
    if pedido["status"] in ["Cancelado", "Finalizado"]:
        raise ValueError("Este pedido não pode gerar romaneio.")

    itens = carregar_itens_pedido_loja(pedido_id)
    if not itens:
        raise ValueError("O pedido não possui itens.")
    if any(i["status_expedicao"] not in ["Separado", "Indisponível"] for i in itens):
        raise ValueError("Todos os itens precisam estar separados ou marcados como indisponíveis antes do romaneio.")

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("SELECT * FROM romaneios_loja WHERE pedido_id = ?", (pedido_id,))
        existente = cursor.fetchone()
        if existente and existente["status"] not in ["Em preparação"]:
            conn.commit()
            return existente["id"]

        if existente:
            romaneio_id = existente["id"]
        else:
            codigo = _codigo_romaneio(cursor, pedido)
            cursor.execute("""
                INSERT INTO romaneios_loja
                (codigo, pedido_id, rodada_id, loja_id, status, data_criacao, criado_por, observacao)
                VALUES (?, ?, ?, ?, 'Em preparação', ?, ?, ?)
            """, (
                codigo, pedido_id, pedido["rodada_id"], pedido["loja_id"],
                _agora_texto(), _usuario_atual(), pedido["observacao"] or ""
            ))
            romaneio_id = cursor.lastrowid

        for item in itens:
            descricao = _descricao_item_romaneio(item)
            qtd_conf_com = 0.0 if item["status_expedicao"] == "Indisponível" else float(item["quantidade_separada_comercial"] or 0)
            qtd_conf_base = 0.0 if item["status_expedicao"] == "Indisponível" else float(item["quantidade_separada_base"] or 0)
            status_item = "Indisponível" if item["status_expedicao"] == "Indisponível" else "Aguardando conferência"
            cursor.execute("""
                INSERT INTO romaneio_itens
                (romaneio_id, item_pedido_id, tipo_item, produto_id, sabor_1_id, sabor_2_id,
                 descricao, origem_nome, embalagem_nome, fator_conversao,
                 quantidade_pedida_comercial, quantidade_pedida_base,
                 quantidade_separada_comercial, quantidade_separada_base,
                 quantidade_conferida_comercial, quantidade_conferida_base, status, observacao)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(romaneio_id, item_pedido_id) DO UPDATE SET
                    descricao = excluded.descricao,
                    origem_nome = excluded.origem_nome,
                    embalagem_nome = excluded.embalagem_nome,
                    fator_conversao = excluded.fator_conversao,
                    quantidade_pedida_comercial = excluded.quantidade_pedida_comercial,
                    quantidade_pedida_base = excluded.quantidade_pedida_base,
                    quantidade_separada_comercial = excluded.quantidade_separada_comercial,
                    quantidade_separada_base = excluded.quantidade_separada_base,
                    quantidade_conferida_comercial = CASE WHEN romaneio_itens.status = 'Aguardando conferência' THEN excluded.quantidade_conferida_comercial ELSE romaneio_itens.quantidade_conferida_comercial END,
                    quantidade_conferida_base = CASE WHEN romaneio_itens.status = 'Aguardando conferência' THEN excluded.quantidade_conferida_base ELSE romaneio_itens.quantidade_conferida_base END,
                    status = CASE WHEN romaneio_itens.status = 'Aguardando conferência' THEN excluded.status ELSE romaneio_itens.status END
            """, (
                romaneio_id, item["id"], item["tipo_item"], item["produto_id"], item["sabor_1_id"], item["sabor_2_id"],
                descricao, item["origem_nome"] or "Origem não definida", item["embalagem_nome"] or "un",
                float(item["fator_conversao"] or 1), float(item["quantidade_comercial"] or 0), float(item["quantidade_base"] or 0),
                float(item["quantidade_separada_comercial"] or 0), float(item["quantidade_separada_base"] or 0),
                qtd_conf_com, qtd_conf_base, status_item, item["observacao_expedicao"] or ""
            ))

        cursor.execute("UPDATE pedidos_loja SET status = 'Em expedição' WHERE id = ?", (pedido_id,))
        conn.commit()
        return romaneio_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _lotes_fefo_cursor(cursor, produto_id):
    hoje = datetime.now().strftime("%Y-%m-%d")
    cursor.execute("""
        SELECT lv.*, p.nome AS produto_nome, p.unidade_padrao
        FROM lotes_validade lv
        INNER JOIN produtos_estoque p ON p.id = lv.produto_id
        WHERE lv.produto_id = ?
          AND lv.status != 'Cancelado'
          AND lv.quantidade_atual > 0.0000001
          AND lv.data_validade >= ?
          AND COALESCE(lv.local_tipo, 'Estoque Central') = 'Estoque Central'
        ORDER BY lv.data_validade ASC, lv.data_producao ASC, lv.id ASC
    """, (produto_id, hoje))
    return [dict(row) for row in cursor.fetchall()]


def _produto_controla_validade_cursor(cursor, produto_id):
    cursor.execute("SELECT controla_validade FROM produtos_estoque WHERE id = ?", (produto_id,))
    row = cursor.fetchone()
    return bool(row and int(row["controla_validade"] or 0) == 1)


def _sugerir_fefo_cursor(cursor, produto_id, quantidade):
    quantidade = float(quantidade or 0)
    if quantidade <= 0 or not _produto_controla_validade_cursor(cursor, produto_id):
        return []
    restante = quantidade
    alocacoes = []
    for lote in _lotes_fefo_cursor(cursor, produto_id):
        usar = min(restante, float(lote["quantidade_atual"] or 0))
        if usar > 0:
            alocacoes.append({"lote": lote, "quantidade": usar})
            restante -= usar
        if restante <= 0.0000001:
            break
    if restante > 0.0000001:
        cursor.execute("SELECT nome, unidade_padrao FROM produtos_estoque WHERE id = ?", (produto_id,))
        produto = cursor.fetchone()
        nome = produto["nome"] if produto else f"Produto #{produto_id}"
        unidade = produto["unidade_padrao"] if produto else "un"
        disponivel = quantidade - restante
        raise ValueError(f"Lotes válidos insuficientes para {nome}: {disponivel:.2f} {unidade} disponíveis por FEFO, {quantidade:.2f} necessários.")
    return alocacoes


def _necessidades_item_romaneio(item):
    qtd_com = float(item["quantidade_conferida_comercial"] or 0)
    if item["tipo_item"] == "Misto":
        return [
            (item["sabor_1_id"], qtd_com * 18),
            (item["sabor_2_id"], qtd_com * 17),
        ]
    return [(item["produto_id"], float(item["quantidade_conferida_base"] or 0))]


def carregar_itens_romaneio(romaneio_id, incluir_sugestoes=True):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ri.*, p.codigo AS produto_codigo, p.nome AS produto_nome, p.unidade_padrao,
               s1.nome AS sabor_1_nome, s2.nome AS sabor_2_nome
        FROM romaneio_itens ri
        LEFT JOIN produtos_estoque p ON p.id = ri.produto_id
        LEFT JOIN produtos_estoque s1 ON s1.id = ri.sabor_1_id
        LEFT JOIN produtos_estoque s2 ON s2.id = ri.sabor_2_id
        WHERE ri.romaneio_id = ?
        ORDER BY ri.origem_nome, ri.id
    """, (romaneio_id,))
    itens = [dict(row) for row in cursor.fetchall()]
    for item in itens:
        cursor.execute("""
            SELECT rl.*, lv.codigo_lote, lv.data_validade, p.nome AS produto_lote_nome, p.unidade_padrao AS unidade_lote
            FROM romaneio_lotes rl
            INNER JOIN lotes_validade lv ON lv.id = rl.lote_id
            INNER JOIN produtos_estoque p ON p.id = rl.produto_id
            WHERE rl.romaneio_item_id = ?
            ORDER BY lv.data_validade, lv.id
        """, (item["id"],))
        item["lotes"] = [dict(row) for row in cursor.fetchall()]
        for lote in item["lotes"]:
            lote["data_validade_br"] = _data_iso_para_br(lote.get("data_validade"))
        item["sugestoes_fefo"] = []
        if incluir_sugestoes and not item["lotes"] and item["status"] != "Indisponível":
            try:
                for produto_id, qtd in _necessidades_item_romaneio(item):
                    if produto_id and qtd > 0:
                        for aloc in _sugerir_fefo_cursor(cursor, produto_id, qtd):
                            item["sugestoes_fefo"].append({
                                "produto_id": produto_id,
                                "produto_nome": aloc["lote"]["produto_nome"],
                                "codigo_lote": aloc["lote"]["codigo_lote"],
                                "data_validade": _data_iso_para_br(aloc["lote"]["data_validade"]),
                                "quantidade": aloc["quantidade"],
                                "unidade": aloc["lote"]["unidade_padrao"],
                            })
            except ValueError as erro:
                item["alerta_fefo"] = str(erro)
    conn.close()
    return itens


def salvar_conferencia_romaneio(romaneio_id, form):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio:
        raise ValueError("Romaneio não encontrado.")
    if romaneio["status"] != "Em preparação":
        raise ValueError("A conferência deste romaneio já foi encerrada.")

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("SELECT * FROM romaneio_itens WHERE romaneio_id = ? ORDER BY id", (romaneio_id,))
        itens = cursor.fetchall()
        for item in itens:
            if item["status"] == "Indisponível":
                continue
            bruto = str(form.get(f"qtd_confirmada_{item['id']}", item["quantidade_separada_comercial"]) or "0").replace(",", ".")
            try:
                qtd = float(bruto)
            except ValueError:
                qtd = 0.0
            maximo = float(item["quantidade_separada_comercial"] or 0)
            if qtd < 0 or qtd > maximo + 1e-9:
                raise ValueError(f"Quantidade conferida inválida para {item['descricao']}.")
            fator = float(item["fator_conversao"] or 1)
            qtd_base = qtd * fator
            if qtd <= 0:
                status_item = "Indisponível"
            elif qtd < maximo - 1e-9:
                status_item = "Divergência"
            else:
                status_item = "Conferido"
            cursor.execute("""
                UPDATE romaneio_itens
                SET quantidade_conferida_comercial = ?, quantidade_conferida_base = ?, status = ?, observacao = ?
                WHERE id = ?
            """, (
                qtd, qtd_base, status_item,
                str(form.get(f"observacao_{item['id']}", item["observacao"] or "") or "").strip(),
                item["id"]
            ))
        motorista_id = str(form.get("motorista_id", "") or "").strip()
        veiculo_id = str(form.get("veiculo_id", "") or "").strip()
        motorista = None
        veiculo = None
        if motorista_id:
            cursor.execute("SELECT * FROM motoristas WHERE id = ? AND ativo = 1", (motorista_id,))
            motorista = cursor.fetchone()
            if not motorista:
                raise ValueError("Selecione um motorista ativo e cadastrado.")
        if veiculo_id:
            cursor.execute("SELECT * FROM veiculos WHERE id = ? AND ativo = 1", (veiculo_id,))
            veiculo = cursor.fetchone()
            if not veiculo:
                raise ValueError("Selecione um veículo ativo e cadastrado.")
        cursor.execute("""
            UPDATE romaneios_loja
            SET motorista_id = ?, veiculo_id = ?,
                motorista = ?, veiculo = ?, placa = ?, observacao = ?
            WHERE id = ?
        """, (
            motorista["id"] if motorista else None,
            veiculo["id"] if veiculo else None,
            motorista["nome"] if motorista else None,
            veiculo["descricao"] if veiculo else None,
            veiculo["placa"] if veiculo else None,
            str(form.get("observacao", "") or "").strip(),
            romaneio_id
        ))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def confirmar_conferencia_romaneio(romaneio_id):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio or romaneio["status"] != "Em preparação":
        raise ValueError("Somente romaneios em preparação podem ser conferidos.")
    if not romaneio["motorista_id"] or not romaneio["veiculo_id"]:
        raise ValueError("Selecione e salve o motorista e o veículo antes de confirmar a conferência.")
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("SELECT * FROM romaneio_itens WHERE romaneio_id = ? ORDER BY id", (romaneio_id,))
        itens = cursor.fetchall()
        if not itens:
            raise ValueError("O romaneio não possui itens.")
        cursor.execute("DELETE FROM romaneio_lotes WHERE romaneio_item_id IN (SELECT id FROM romaneio_itens WHERE romaneio_id = ?)", (romaneio_id,))
        for item in itens:
            if item["status"] == "Aguardando conferência":
                raise ValueError("Salve a conferência de todos os itens antes de confirmar.")
            if item["status"] == "Indisponível":
                continue
            for produto_id, quantidade in _necessidades_item_romaneio(item):
                if not produto_id or quantidade <= 0:
                    continue
                saldo = obter_saldo_produto(produto_id)
                if quantidade > saldo + 1e-9:
                    cursor.execute("SELECT nome, unidade_padrao FROM produtos_estoque WHERE id = ?", (produto_id,))
                    produto = cursor.fetchone()
                    raise ValueError(f"Saldo insuficiente de {produto['nome']}: {saldo:.2f} {produto['unidade_padrao']} disponíveis, {quantidade:.2f} necessários.")
                for alocacao in _sugerir_fefo_cursor(cursor, produto_id, quantidade):
                    cursor.execute("""
                        INSERT INTO romaneio_lotes (romaneio_item_id, produto_id, lote_id, quantidade)
                        VALUES (?, ?, ?, ?)
                    """, (item["id"], produto_id, alocacao["lote"]["id"], alocacao["quantidade"]))
        cursor.execute("""
            UPDATE romaneios_loja
            SET status = 'Conferido', data_conferencia = ?, conferido_por = ?
            WHERE id = ?
        """, (_agora_texto(), _usuario_atual(), romaneio_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _movimentar_saida_romaneio_cursor(cursor, romaneio, item, produto_id, quantidade, observacao):
    cursor.execute("SELECT nome, unidade_padrao, controla_validade FROM produtos_estoque WHERE id = ?", (produto_id,))
    produto = cursor.fetchone()
    if not produto:
        raise ValueError("Produto do romaneio não encontrado.")
    saldo = obter_saldo_produto(produto_id)
    if quantidade > saldo + 1e-9:
        raise ValueError(f"Saldo insuficiente de {produto['nome']} no momento da saída.")

    cursor.execute("""
        INSERT INTO movimentacoes_estoque
        (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_nome, fator_conversao,
         data_movimentacao, origem, origem_id, usuario, observacao)
        VALUES (?, 'Saída', ?, ?, ?, 1, ?, 'Romaneio Loja', ?, ?, ?)
    """, (
        produto_id, quantidade, quantidade, produto["unidade_padrao"],
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"), romaneio["id"], _usuario_atual(), observacao
    ))
    movimento_id = cursor.lastrowid

    if int(produto["controla_validade"] or 0) == 1:
        cursor.execute("""
            SELECT rl.*, lv.quantidade_atual, lv.codigo_lote, lv.data_validade, lv.status
            FROM romaneio_lotes rl
            INNER JOIN lotes_validade lv ON lv.id = rl.lote_id
            WHERE rl.romaneio_item_id = ? AND rl.produto_id = ?
            ORDER BY lv.data_validade, lv.id
        """, (item["id"], produto_id))
        lotes = cursor.fetchall()
        total_alocado = sum(float(x["quantidade"] or 0) for x in lotes)
        if abs(total_alocado - quantidade) > 0.0001:
            raise ValueError(f"A alocação FEFO de {produto['nome']} precisa ser refeita antes da saída.")
        hoje = datetime.now().strftime("%Y-%m-%d")
        for aloc in lotes:
            if aloc["status"] == "Cancelado" or str(aloc["data_validade"]) < hoje or float(aloc["quantidade_atual"] or 0) + 1e-9 < float(aloc["quantidade"] or 0):
                raise ValueError(f"O lote {aloc['codigo_lote']} não está mais disponível. Refaça a conferência FEFO.")
            novo = float(aloc["quantidade_atual"] or 0) - float(aloc["quantidade"] or 0)
            novo_status = "Encerrado" if novo <= 0.0000001 else "Ativo"
            cursor.execute("""
                UPDATE lotes_validade
                SET quantidade_atual = ?, status = ?, atualizado_em = ?,
                    encerrado_em = CASE WHEN ? = 'Encerrado' THEN ? ELSE encerrado_em END
                WHERE id = ?
            """, (novo, novo_status, _agora_texto(), novo_status, _agora_texto(), aloc["lote_id"]))
            cursor.execute("""
                INSERT INTO movimentacoes_validade
                (lote_id, tipo, quantidade, data_movimentacao, usuario, estoque_movimentacao_id, observacao)
                VALUES (?, 'Baixa', ?, ?, ?, ?, ?)
            """, (
                aloc["lote_id"], aloc["quantidade"], _agora_texto(), _usuario_atual(), movimento_id,
                f"Saída pelo romaneio {romaneio['codigo']} para {romaneio['loja_nome']}."
            ))
    return movimento_id


def registrar_saida_romaneio(romaneio_id):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio or romaneio["status"] != "Conferido":
        raise ValueError("O romaneio precisa estar conferido antes da saída.")
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("SELECT * FROM romaneio_itens WHERE romaneio_id = ? ORDER BY id", (romaneio_id,))
        itens = cursor.fetchall()
        cursor.execute("DELETE FROM romaneio_custos_itens WHERE romaneio_id = ?", (romaneio_id,))
        cache_custos = {}
        custo_total_romaneio = 0.0
        for item in itens:
            if item["status"] == "Indisponível":
                continue
            if item["tipo_item"] == "Misto":
                qtd_com = float(item["quantidade_conferida_comercial"] or 0)
                qtd_1 = qtd_com * 18
                qtd_2 = qtd_com * 17
                mov1 = _movimentar_saida_romaneio_cursor(cursor, romaneio, item, item["sabor_1_id"], qtd_1, f"{item['descricao']} — sabor 18.")
                mov2 = _movimentar_saida_romaneio_cursor(cursor, romaneio, item, item["sabor_2_id"], qtd_2, f"{item['descricao']} — sabor 17.")
                custo_total_romaneio += _registrar_snapshot_custo_romaneio_cursor(cursor, romaneio_id, item["id"], item["sabor_1_id"], qtd_1, cache_custos)
                custo_total_romaneio += _registrar_snapshot_custo_romaneio_cursor(cursor, romaneio_id, item["id"], item["sabor_2_id"], qtd_2, cache_custos)
                cursor.execute("UPDATE itens_pedido_loja SET movimento_estoque_id = ?, movimento_estoque_2_id = ? WHERE id = ?", (mov1, mov2, item["item_pedido_id"]))
            else:
                qtd = float(item["quantidade_conferida_base"] or 0)
                if qtd <= 0:
                    continue
                mov = _movimentar_saida_romaneio_cursor(cursor, romaneio, item, item["produto_id"], qtd, item["descricao"])
                custo_total_romaneio += _registrar_snapshot_custo_romaneio_cursor(cursor, romaneio_id, item["id"], item["produto_id"], qtd, cache_custos)
                cursor.execute("UPDATE itens_pedido_loja SET movimento_estoque_id = ? WHERE id = ?", (mov, item["item_pedido_id"]))
        cursor.execute("""
            UPDATE romaneios_loja
            SET status = 'Em rota', data_saida = ?, saida_por = ?, custo_total_snapshot = ?
            WHERE id = ?
        """, (_agora_texto(), _usuario_atual(), custo_total_romaneio, romaneio_id))
        cursor.execute("UPDATE pedidos_loja SET status = 'Em rota' WHERE id = ?", (romaneio["pedido_id"],))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def registrar_retorno_romaneio(romaneio_id, recebido_loja_por, observacao_retorno=""):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio or romaneio["status"] != "Em rota":
        raise ValueError("Somente romaneios em rota podem registrar o retorno da via.")
    recebido_loja_por = str(recebido_loja_por or "").strip()
    if not recebido_loja_por:
        raise ValueError("Informe quem recebeu o pedido na loja, conforme a via retornada.")
    agora = _agora_texto()
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("""
            UPDATE romaneios_loja
            SET status = 'Finalizado', data_retorno = ?, retorno_por = ?,
                recebido_loja_por = ?, observacao_retorno = ?, data_finalizacao = ?
            WHERE id = ?
        """, (agora, _usuario_atual(), recebido_loja_por, str(observacao_retorno or "").strip(), agora, romaneio_id))
        cursor.execute("""
            UPDATE pedidos_loja
            SET status = 'Finalizado', data_finalizacao = ?, finalizado_por = ?
            WHERE id = ?
        """, (agora, _usuario_atual(), romaneio["pedido_id"]))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def cancelar_preparacao_romaneio(romaneio_id):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio:
        raise ValueError("Romaneio não encontrado.")
    if romaneio["status"] != "Em preparação":
        raise ValueError("Somente romaneios ainda em preparação podem voltar para a separação.")
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("DELETE FROM romaneio_lotes WHERE romaneio_item_id IN (SELECT id FROM romaneio_itens WHERE romaneio_id = ?)", (romaneio_id,))
        cursor.execute("DELETE FROM romaneio_itens WHERE romaneio_id = ?", (romaneio_id,))
        cursor.execute("DELETE FROM romaneios_loja WHERE id = ?", (romaneio_id,))
        atualizar_status_pedido_loja(cursor, romaneio["pedido_id"])
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def contar_romaneios_status():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT status, COUNT(*) AS total FROM romaneios_loja GROUP BY status")
    resumo = {row["status"]: int(row["total"] or 0) for row in cursor.fetchall()}
    conn.close()
    return resumo

# ============================================================
# CUSTOS GERENCIAIS - V3.1
# ============================================================

def _calcular_custo_produto_cursor(cursor, produto_id, cache=None, pilha=None):
    """Calcula custo por unidade base, respeitando o método do produto.

    O cálculo nunca movimenta estoque. Quando usa ficha técnica, percorre fichas
    encadeadas e usa o custo vigente dos componentes. O resultado é congelado
    no romaneio no momento da saída.
    """
    cache = cache if cache is not None else {}
    pilha = set(pilha or set())
    produto_id = int(produto_id)

    if produto_id in cache:
        return dict(cache[produto_id])

    cursor.execute("""
        SELECT p.*, c.nome AS categoria_nome
        FROM produtos_estoque p
        LEFT JOIN categorias_estoque c ON c.id = p.categoria_id
        WHERE p.id = ?
    """, (produto_id,))
    produto = cursor.fetchone()
    if not produto:
        return {
            "produto_id": produto_id,
            "custo_unitario": 0.0,
            "origem": "Produto não encontrado",
            "ficha_id": None,
            "ficha_versao": None,
            "alertas": ["Produto não encontrado no cadastro oficial."],
            "total_receita": 0.0,
        }

    custo_padrao = float(produto["custo_padrao"] or 0)
    custo_medio = float(produto["custo_medio"] or 0) if "custo_medio" in produto.keys() else 0.0
    ultimo_custo = float(produto["ultimo_custo"] or 0) if "ultimo_custo" in produto.keys() else 0.0
    metodo = str(produto["metodo_custo"] or "Automático")
    if metodo not in METODOS_CUSTO:
        metodo = "Automático"

    if produto_id in pilha:
        resultado = {
            "produto_id": produto_id,
            "custo_unitario": custo_padrao,
            "origem": "Custo padrão (ciclo detectado)",
            "ficha_id": None,
            "ficha_versao": None,
            "alertas": [f"Ciclo de ficha técnica detectado em {produto['nome']}."],
            "total_receita": 0.0,
        }
        cache[produto_id] = resultado
        return dict(resultado)

    if metodo == "Custo padrão":
        resultado = {
            "produto_id": produto_id,
            "custo_unitario": custo_padrao,
            "origem": "Custo padrão",
            "ficha_id": None,
            "ficha_versao": None,
            "alertas": [] if custo_padrao > 0 else ["Custo padrão zerado."],
            "total_receita": 0.0,
        }
        cache[produto_id] = resultado
        return dict(resultado)

    cursor.execute("""
        SELECT * FROM fichas_tecnicas
        WHERE produto_final_id = ? AND ativo = 1
        ORDER BY id DESC LIMIT 1
    """, (produto_id,))
    ficha = cursor.fetchone()

    if not ficha:
        alertas = ["Produto sem ficha técnica ativa."]
        if metodo == "Automático" and custo_medio > 0:
            custo_base = custo_medio
            origem_base = "Custo médio"
            alertas = []
        elif custo_padrao > 0:
            custo_base = custo_padrao
            origem_base = "Custo padrão"
        elif ultimo_custo > 0:
            custo_base = ultimo_custo
            origem_base = "Último custo"
        else:
            custo_base = 0.0
            origem_base = "Sem custo"
            alertas.append("Produto com custo zerado.")
        resultado = {
            "produto_id": produto_id,
            "custo_unitario": custo_base,
            "origem": origem_base,
            "ficha_id": None,
            "ficha_versao": None,
            "alertas": alertas,
            "total_receita": 0.0,
        }
        cache[produto_id] = resultado
        return dict(resultado)

    cursor.execute("""
        SELECT i.*, p.nome AS componente_nome, p.unidade_padrao AS componente_unidade
        FROM itens_ficha_tecnica i
        INNER JOIN produtos_estoque p ON p.id = i.insumo_produto_id
        WHERE i.ficha_id = ? AND i.ativo = 1
        ORDER BY i.id
    """, (ficha["id"],))
    itens = cursor.fetchall()
    alertas = []
    total_receita = 0.0
    nova_pilha = set(pilha)
    nova_pilha.add(produto_id)

    if not itens:
        alertas.append("Ficha técnica sem componentes ativos.")

    for item in itens:
        custo_componente = _calcular_custo_produto_cursor(
            cursor,
            item["insumo_produto_id"],
            cache=cache,
            pilha=nova_pilha,
        )
        try:
            quantidade_base = converter_quantidade_planejamento(
                float(item["quantidade"] or 0),
                item["unidade"],
                item["componente_unidade"],
            )
        except Exception as exc:
            alertas.append(
                f"{item['componente_nome']}: conversão de {item['unidade']} para "
                f"{item['componente_unidade']} não disponível ({exc})."
            )
            continue

        custo_unitario_componente = float(custo_componente["custo_unitario"] or 0)
        total_receita += quantidade_base * custo_unitario_componente
        if custo_unitario_componente <= 0:
            alertas.append(f"{item['componente_nome']} está com custo zerado.")
        for alerta in custo_componente.get("alertas", []):
            if alerta and alerta not in alertas:
                alertas.append(f"{item['componente_nome']}: {alerta}")

    try:
        rendimento_base = converter_quantidade_planejamento(
            float(ficha["rendimento_quantidade"] or 0),
            ficha["rendimento_unidade"],
            produto["unidade_padrao"],
        )
    except Exception as exc:
        rendimento_base = 0.0
        alertas.append(f"Rendimento incompatível com a unidade base do produto ({exc}).")

    custo_ficha = total_receita / rendimento_base if rendimento_base > 0 else 0.0

    if metodo == "Automático" and custo_ficha <= 0 and custo_padrao > 0:
        resultado = {
            "produto_id": produto_id,
            "custo_unitario": custo_padrao,
            "origem": "Custo padrão (fallback)",
            "ficha_id": ficha["id"],
            "ficha_versao": int(ficha["versao"] or 1),
            "alertas": alertas + ["A ficha não gerou custo positivo; usado o custo padrão."],
            "total_receita": total_receita,
        }
    else:
        resultado = {
            "produto_id": produto_id,
            "custo_unitario": custo_ficha,
            "origem": f"Ficha Técnica v{int(ficha['versao'] or 1)}",
            "ficha_id": ficha["id"],
            "ficha_versao": int(ficha["versao"] or 1),
            "alertas": alertas,
            "total_receita": total_receita,
        }

    cache[produto_id] = resultado
    return dict(resultado)


def calcular_custo_produto(produto_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        return _calcular_custo_produto_cursor(cursor, produto_id, cache={}, pilha=set())
    finally:
        conn.close()


def carregar_custos_produtos(busca="", categoria_id="", somente_alertas=False):
    busca = str(busca or "").strip().lower()
    categoria_id = str(categoria_id or "").strip()
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT p.*, c.nome AS categoria_nome
        FROM produtos_estoque p
        LEFT JOIN categorias_estoque c ON c.id = p.categoria_id
        ORDER BY CAST(p.codigo AS INTEGER), p.nome
    """)
    produtos = [dict(row) for row in cursor.fetchall()]
    cache = {}
    resultado = []
    for produto in produtos:
        if busca and busca not in str(produto.get("codigo") or "").lower() and busca not in str(produto.get("nome") or "").lower():
            continue
        if categoria_id and str(produto.get("categoria_id") or "") != categoria_id:
            continue
        custo = _calcular_custo_produto_cursor(cursor, produto["id"], cache=cache, pilha=set())
        produto.update({
            "custo_calculado": float(custo["custo_unitario"] or 0),
            "origem_custo": custo["origem"],
            "alertas_custo": custo.get("alertas", []),
            "ficha_custo_id": custo.get("ficha_id"),
            "ficha_custo_versao": custo.get("ficha_versao"),
        })
        if somente_alertas and not produto["alertas_custo"]:
            continue
        resultado.append(produto)
    conn.close()
    return resultado


def _registrar_snapshot_custo_romaneio_cursor(cursor, romaneio_id, romaneio_item_id, produto_id, quantidade_base, cache):
    quantidade_base = float(quantidade_base or 0)
    if quantidade_base <= 0:
        return 0.0
    cursor.execute("SELECT unidade_padrao, nome FROM produtos_estoque WHERE id = ?", (produto_id,))
    produto = cursor.fetchone()
    if not produto:
        return 0.0
    custo = _calcular_custo_produto_cursor(cursor, produto_id, cache=cache, pilha=set())
    custo_unitario = float(custo["custo_unitario"] or 0)
    custo_total = quantidade_base * custo_unitario
    observacao = " | ".join(custo.get("alertas", [])[:6])
    cursor.execute("""
        INSERT INTO romaneio_custos_itens
        (romaneio_id, romaneio_item_id, produto_id, quantidade_base, unidade,
         custo_unitario_snapshot, custo_total_snapshot, origem_custo,
         ficha_id, ficha_versao, observacao, data_snapshot)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(romaneio_item_id, produto_id) DO UPDATE SET
            quantidade_base = excluded.quantidade_base,
            unidade = excluded.unidade,
            custo_unitario_snapshot = excluded.custo_unitario_snapshot,
            custo_total_snapshot = excluded.custo_total_snapshot,
            origem_custo = excluded.origem_custo,
            ficha_id = excluded.ficha_id,
            ficha_versao = excluded.ficha_versao,
            observacao = excluded.observacao,
            data_snapshot = excluded.data_snapshot
    """, (
        romaneio_id, romaneio_item_id, produto_id, quantidade_base, produto["unidade_padrao"],
        custo_unitario, custo_total, custo["origem"], custo.get("ficha_id"),
        custo.get("ficha_versao"), observacao, _agora_texto(),
    ))
    return custo_total


def carregar_custos_romaneio(romaneio_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT rc.*, p.codigo AS produto_codigo, p.nome AS produto_nome,
               c.nome AS categoria_nome, ri.descricao AS item_descricao
        FROM romaneio_custos_itens rc
        INNER JOIN produtos_estoque p ON p.id = rc.produto_id
        LEFT JOIN categorias_estoque c ON c.id = p.categoria_id
        INNER JOIN romaneio_itens ri ON ri.id = rc.romaneio_item_id
        WHERE rc.romaneio_id = ?
        ORDER BY ri.id, p.nome
    """, (romaneio_id,))
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


def _filtrar_data_custo(data_texto, inicio, fim):
    data = _data_br_para_objeto(data_texto)
    if inicio and (not data or data < inicio):
        return False
    if fim and (not data or data > fim):
        return False
    return True


def carregar_relatorio_custos_lojas(data_inicio="", data_fim="", loja_id="", categoria_id="", produto_id=""):
    inicio = _data_br_para_objeto(_data_html_para_br(data_inicio)) if data_inicio else None
    fim = _data_br_para_objeto(_data_html_para_br(data_fim)) if data_fim else None
    if fim:
        fim = fim.replace(hour=23, minute=59, second=59)

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT rc.*, r.codigo AS romaneio_codigo, r.data_saida, r.status AS romaneio_status,
               r.loja_id, l.nome AS loja_nome, l.codigo AS loja_codigo,
               p.codigo AS produto_codigo, p.nome AS produto_nome, p.categoria_id,
               c.nome AS categoria_nome, ri.descricao AS item_descricao
        FROM romaneio_custos_itens rc
        INNER JOIN romaneios_loja r ON r.id = rc.romaneio_id
        INNER JOIN lojas l ON l.id = r.loja_id
        INNER JOIN produtos_estoque p ON p.id = rc.produto_id
        LEFT JOIN categorias_estoque c ON c.id = p.categoria_id
        INNER JOIN romaneio_itens ri ON ri.id = rc.romaneio_item_id
        WHERE r.status IN ('Em rota', 'Finalizado')
        ORDER BY r.id DESC, ri.id, p.nome
    """)
    detalhes = []
    por_loja = {}
    por_produto = {}
    por_categoria = {}
    romaneios = set()
    total_geral = 0.0

    for row in cursor.fetchall():
        item = dict(row)
        if loja_id and str(item["loja_id"]) != str(loja_id):
            continue
        if categoria_id and str(item.get("categoria_id") or "") != str(categoria_id):
            continue
        if produto_id and str(item["produto_id"]) != str(produto_id):
            continue
        if not _filtrar_data_custo(item["data_saida"], inicio, fim):
            continue
        detalhes.append(item)
        valor = float(item["custo_total_snapshot"] or 0)
        qtd = float(item["quantidade_base"] or 0)
        total_geral += valor
        romaneios.add(item["romaneio_id"])

        loja = por_loja.setdefault(item["loja_nome"], {"nome": item["loja_nome"], "total": 0.0, "romaneios": set(), "itens": 0})
        loja["total"] += valor
        loja["romaneios"].add(item["romaneio_id"])
        loja["itens"] += 1

        prod = por_produto.setdefault(item["produto_id"], {
            "produto_id": item["produto_id"], "codigo": item["produto_codigo"], "nome": item["produto_nome"],
            "unidade": item["unidade"], "quantidade": 0.0, "total": 0.0,
        })
        prod["quantidade"] += qtd
        prod["total"] += valor

        categoria_nome = item["categoria_nome"] or "Sem categoria"
        cat = por_categoria.setdefault(categoria_nome, {"nome": categoria_nome, "total": 0.0, "itens": 0})
        cat["total"] += valor
        cat["itens"] += 1

    conn.close()
    resumo_lojas = []
    for dados in por_loja.values():
        resumo_lojas.append({
            "nome": dados["nome"], "total": dados["total"], "itens": dados["itens"],
            "romaneios": len(dados["romaneios"]),
        })
    resumo_lojas.sort(key=lambda x: x["total"], reverse=True)
    resumo_produtos = sorted(por_produto.values(), key=lambda x: x["total"], reverse=True)
    resumo_categorias = sorted(por_categoria.values(), key=lambda x: x["total"], reverse=True)
    return {
        "detalhes": detalhes,
        "lojas": resumo_lojas,
        "produtos": resumo_produtos,
        "categorias": resumo_categorias,
        "total_geral": total_geral,
        "total_romaneios": len(romaneios),
    }


def resumo_custos_gerais(data_inicio="", data_fim=""):
    lojas = carregar_relatorio_custos_lojas(data_inicio=data_inicio, data_fim=data_fim)
    data_inicio_br = _data_html_para_br(data_inicio) if data_inicio else ""
    data_fim_br = _data_html_para_br(data_fim) if data_fim else ""
    centros, detalhes_centros, total_centros = carregar_relatorio_centros_custo(
        data_inicio=data_inicio_br,
        data_fim=data_fim_br,
    )
    return {
        "lojas": lojas["total_geral"],
        "centros": total_centros,
        "geral": lojas["total_geral"] + total_centros,
        "romaneios": lojas["total_romaneios"],
        "saidas_centros": len(detalhes_centros),
    }


# ============================================================
# NOTAS DE ENTRADA / EVENTOS DE COMPRA - V3.2
# ============================================================

def _decimal_positivo_ou_zero(valor, campo="Valor"):
    try:
        numero = float(str(valor or 0).replace(",", "."))
    except Exception:
        raise ValueError(f"{campo} inválido.")
    if numero < 0:
        raise ValueError(f"{campo} não pode ser negativo.")
    return numero


def _data_html_para_br_ou_hoje(valor):
    valor = str(valor or "").strip()
    if not valor:
        return datetime.now().strftime("%d/%m/%Y")
    try:
        return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return valor


def carregar_pedidos_compra_para_nota():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT pc.id, pc.solicitante, pc.data_solicitacao, pc.status,
               COUNT(ic.id) AS total_itens
        FROM pedidos_compra pc
        LEFT JOIN itens_compra ic ON ic.pedido_id = pc.id
        WHERE pc.status IN ('Pendente', 'Comprado')
        GROUP BY pc.id
        ORDER BY pc.id DESC
    """)
    dados = cursor.fetchall()
    conn.close()
    return dados


def carregar_itens_compra_disponiveis_nota(pedido_compra_id=None):
    conn = conectar()
    cursor = conn.cursor()
    filtros = ["ic.status != 'Cancelado'"]
    params = []
    if pedido_compra_id:
        filtros.append("ic.pedido_id = ?")
        params.append(int(pedido_compra_id))
    cursor.execute(f"""
        SELECT ic.*, pc.solicitante, f.nome AS fornecedor_nome
        FROM itens_compra ic
        INNER JOIN pedidos_compra pc ON pc.id = ic.pedido_id
        LEFT JOIN fornecedores f ON f.id = ic.fornecedor_id
        WHERE {' AND '.join(filtros)}
        ORDER BY ic.pedido_id DESC, ic.id
    """, params)
    dados = cursor.fetchall()
    conn.close()
    return dados


def carregar_todas_embalagens_ativas():
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT e.*, p.nome AS produto_nome, p.unidade_padrao
        FROM produto_embalagens e
        INNER JOIN produtos_estoque p ON p.id = e.produto_id
        WHERE e.ativo = 1
        ORDER BY p.nome, e.fator_conversao, e.nome
    """)
    dados = cursor.fetchall(); conn.close(); return dados


def recalcular_totais_nota_entrada(nota_id, conn=None):
    proprio = conn is None
    if proprio:
        conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COALESCE(SUM(valor_total_item),0) AS total FROM itens_nota_entrada WHERE nota_id = ?", (nota_id,))
    valor_produtos = float(cursor.fetchone()["total"] or 0)
    cursor.execute("SELECT desconto_geral, frete, outras_despesas FROM notas_entrada WHERE id = ?", (nota_id,))
    nota = cursor.fetchone()
    if nota:
        valor_total = max(0.0, valor_produtos - float(nota["desconto_geral"] or 0) + float(nota["frete"] or 0) + float(nota["outras_despesas"] or 0))
        cursor.execute("UPDATE notas_entrada SET valor_produtos = ?, valor_total = ? WHERE id = ?", (valor_produtos, valor_total, nota_id))
    if proprio:
        conn.commit(); conn.close()


def buscar_nota_entrada(nota_id):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT ne.*, f.nome AS fornecedor_nome, pc.solicitante AS pedido_solicitante
        FROM notas_entrada ne
        INNER JOIN fornecedores f ON f.id = ne.fornecedor_id
        LEFT JOIN pedidos_compra pc ON pc.id = ne.pedido_compra_id
        WHERE ne.id = ?
    """, (nota_id,))
    dado = cursor.fetchone(); conn.close(); return dado


def carregar_itens_nota_entrada(nota_id):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT ine.*, p.codigo AS produto_codigo, p.nome AS produto_nome,
               p.unidade_padrao AS produto_unidade, cc.nome AS centro_custo_nome,
               ic.descricao AS compra_descricao, ic.pedido_id AS compra_pedido_id
        FROM itens_nota_entrada ine
        LEFT JOIN produtos_estoque p ON p.id = ine.produto_id
        LEFT JOIN centros_custo cc ON cc.id = ine.centro_custo_id
        LEFT JOIN itens_compra ic ON ic.id = ine.item_compra_id
        WHERE ine.nota_id = ?
        ORDER BY ine.id
    """, (nota_id,))
    dados = cursor.fetchall(); conn.close(); return dados


def carregar_notas_entrada(status="", busca="", fornecedor_id=""):
    filtros = []; params = []
    if status:
        filtros.append("ne.status = ?"); params.append(status)
    if fornecedor_id:
        filtros.append("ne.fornecedor_id = ?"); params.append(int(fornecedor_id))
    if busca:
        filtros.append("(ne.numero LIKE ? OR ne.chave_acesso LIKE ? OR f.nome LIKE ?)")
        termo = f"%{busca}%"; params.extend([termo, termo, termo])
    where = "WHERE " + " AND ".join(filtros) if filtros else ""
    conn = conectar(); cursor = conn.cursor()
    cursor.execute(f"""
        SELECT ne.*, f.nome AS fornecedor_nome, COUNT(ine.id) AS total_itens
        FROM notas_entrada ne
        INNER JOIN fornecedores f ON f.id = ne.fornecedor_id
        LEFT JOIN itens_nota_entrada ine ON ine.nota_id = ne.id
        {where}
        GROUP BY ne.id
        ORDER BY ne.id DESC
    """, params)
    dados = cursor.fetchall(); conn.close(); return dados


def criar_nota_entrada(fornecedor_id, numero, serie, chave_acesso, data_emissao, data_entrada,
                       pedido_compra_id, desconto_geral, frete, outras_despesas, observacao):
    numero = str(numero or "").strip()
    if not numero:
        raise ValueError("Informe o número do documento.")
    serie = str(serie or "").strip()
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("""
        SELECT id FROM notas_entrada
        WHERE fornecedor_id = ? AND numero = ? AND COALESCE(serie,'') = ? AND status != 'Cancelada'
    """, (int(fornecedor_id), numero, serie))
    if cursor.fetchone():
        conn.close(); raise ValueError("Já existe um documento com esse fornecedor, número e série.")
    cursor.execute("""
        INSERT INTO notas_entrada
        (fornecedor_id, pedido_compra_id, numero, serie, chave_acesso, data_emissao, data_entrada,
         status, desconto_geral, frete, outras_despesas, observacao, criado_por, data_criacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, 'Rascunho', ?, ?, ?, ?, ?, ?)
    """, (
        int(fornecedor_id), int(pedido_compra_id) if pedido_compra_id else None, numero, serie,
        str(chave_acesso or "").strip(), _data_html_para_br_ou_hoje(data_emissao) if data_emissao else None,
        _data_html_para_br_ou_hoje(data_entrada), _decimal_positivo_ou_zero(desconto_geral, "Desconto"),
        _decimal_positivo_ou_zero(frete, "Frete"), _decimal_positivo_ou_zero(outras_despesas, "Outras despesas"),
        str(observacao or "").strip(), _usuario_atual(), _agora_texto()
    ))
    nota_id = cursor.lastrowid
    conn.commit(); conn.close(); return nota_id


def adicionar_item_nota_entrada(nota_id, descricao, natureza, produto_id, centro_custo_id,
                                item_compra_id, quantidade_faturada, quantidade_bonificada,
                                unidade, embalagem_id, valor_unitario, desconto_item, observacao):
    nota = buscar_nota_entrada(nota_id)
    if not nota or nota["status"] != "Rascunho":
        raise ValueError("Somente documentos em rascunho podem receber itens.")
    if natureza not in NATUREZAS_ITEM_NOTA:
        raise ValueError("Natureza do item inválida.")
    q_fat = _decimal_positivo_ou_zero(quantidade_faturada, "Quantidade faturada")
    q_bon = _decimal_positivo_ou_zero(quantidade_bonificada, "Quantidade bonificada")
    valor_unit = _decimal_positivo_ou_zero(valor_unitario, "Valor unitário")
    desconto = _decimal_positivo_ou_zero(desconto_item, "Desconto do item")
    movimenta = 0 if natureza == "Sem movimentação de estoque" else 1
    embalagem_nome = str(unidade or "un").strip()
    fator_conversao = 1.0
    embalagem_id_valor = None
    if movimenta and not produto_id:
        raise ValueError("Vincule um produto oficial para itens que movimentam estoque.")
    if movimenta and produto_id and embalagem_id:
        embalagem = buscar_embalagem_produto(int(embalagem_id), int(produto_id))
        if not embalagem:
            raise ValueError("A embalagem selecionada não pertence ao produto.")
        embalagem_id_valor = int(embalagem_id)
        embalagem_nome = embalagem["nome"]
        fator_conversao = float(embalagem["fator_conversao"] or 1)
    if movimenta and (q_fat + q_bon) <= 0:
        raise ValueError("Informe a quantidade recebida.")
    if natureza == "Compra normal" and q_fat <= 0:
        raise ValueError("Compra normal precisa de quantidade faturada.")
    if natureza == "Bonificação" and q_bon <= 0:
        raise ValueError("Bonificação precisa de quantidade bonificada.")
    if not movimenta and not centro_custo_id:
        raise ValueError("Informe o centro de custo para item sem movimentação de estoque.")
    quantidade_total = (q_fat + q_bon) * fator_conversao
    valor_total_item = max(0.0, q_fat * valor_unit - desconto)
    descricao = str(descricao or "").strip()
    if not descricao:
        if produto_id:
            p = buscar_produto_estoque(int(produto_id)); descricao = p["nome"] if p else "Item"
        else:
            raise ValueError("Informe a descrição do item.")
    conn = conectar(); cursor = conn.cursor()
    if item_compra_id:
        cursor.execute("SELECT nota_entrada_item_id FROM itens_compra WHERE id = ?", (int(item_compra_id),))
        vinculo = cursor.fetchone()
        if vinculo and vinculo["nota_entrada_item_id"]:
            conn.close()
            raise ValueError("Este item do pedido já está vinculado a outro documento de entrada.")
    cursor.execute("""
        INSERT INTO itens_nota_entrada
        (nota_id, item_compra_id, produto_id, embalagem_id, embalagem_nome, fator_conversao,
         centro_custo_id, descricao, natureza, quantidade_faturada, quantidade_bonificada,
         quantidade_total, unidade, valor_unitario, desconto_item, valor_total_item,
         movimenta_estoque, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        nota_id, int(item_compra_id) if item_compra_id else None,
        int(produto_id) if produto_id else None, embalagem_id_valor, embalagem_nome, fator_conversao,
        int(centro_custo_id) if centro_custo_id else None, descricao, natureza, q_fat, q_bon,
        quantidade_total, str(unidade or embalagem_nome or "un").strip(), valor_unit, desconto,
        valor_total_item, movimenta, str(observacao or "").strip()
    ))
    item_id = cursor.lastrowid
    if item_compra_id:
        cursor.execute("UPDATE itens_compra SET nota_entrada_item_id = ? WHERE id = ?", (item_id, int(item_compra_id)))
    recalcular_totais_nota_entrada(nota_id, conn)
    conn.commit(); conn.close(); return item_id


def excluir_item_nota_entrada(item_id):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("SELECT * FROM itens_nota_entrada WHERE id = ?", (item_id,))
    item = cursor.fetchone()
    if not item:
        conn.close(); return None
    cursor.execute("SELECT status FROM notas_entrada WHERE id = ?", (item["nota_id"],))
    nota = cursor.fetchone()
    if not nota or nota["status"] != "Rascunho":
        conn.close(); raise ValueError("Itens de documento lançado não podem ser excluídos.")
    if item["item_compra_id"]:
        cursor.execute("UPDATE itens_compra SET nota_entrada_item_id = NULL WHERE id = ? AND nota_entrada_item_id = ?", (item["item_compra_id"], item_id))
    cursor.execute("DELETE FROM itens_nota_entrada WHERE id = ?", (item_id,))
    recalcular_totais_nota_entrada(item["nota_id"], conn)
    conn.commit(); conn.close(); return item["nota_id"]


def lancar_nota_entrada(nota_id):
    conn = conectar(); cursor = conn.cursor()
    try:
        cursor.execute("SELECT * FROM notas_entrada WHERE id = ?", (nota_id,))
        nota = cursor.fetchone()
        if not nota or nota["status"] != "Rascunho":
            raise ValueError("Documento não encontrado ou já lançado.")
        cursor.execute("SELECT * FROM itens_nota_entrada WHERE nota_id = ? ORDER BY id", (nota_id,))
        itens = cursor.fetchall()
        if not itens:
            raise ValueError("Adicione pelo menos um item antes de lançar.")

        # Despesas gerais são rateadas apenas entre itens de estoque com valor faturado.
        base_rateio = sum(float(i["valor_total_item"] or 0) for i in itens if i["movimenta_estoque"] and float(i["valor_total_item"] or 0) > 0)
        extras_liquidos = float(nota["frete"] or 0) + float(nota["outras_despesas"] or 0) - float(nota["desconto_geral"] or 0)
        usuario = _usuario_atual(); agora = _agora_texto(); data_mov = nota["data_entrada"] or datetime.now().strftime("%d/%m/%Y")

        for item in itens:
            if not item["movimenta_estoque"]:
                cursor.execute("""
                    UPDATE itens_nota_entrada
                    SET custo_total_entrada = valor_total_item,
                        custo_efetivo_unitario = CASE WHEN quantidade_faturada > 0 THEN valor_total_item / quantidade_faturada ELSE 0 END
                    WHERE id = ?
                """, (item["id"],))
                continue

            produto_id = int(item["produto_id"])
            quantidade_total = float(item["quantidade_total"] or 0)
            if quantidade_total <= 0:
                raise ValueError(f"Item {item['descricao']} sem quantidade recebida.")
            cursor.execute("SELECT * FROM produtos_estoque WHERE id = ?", (produto_id,))
            produto = cursor.fetchone()
            if not produto:
                raise ValueError(f"Produto oficial não encontrado para {item['descricao']}.")

            rateio = 0.0
            if base_rateio > 0 and float(item["valor_total_item"] or 0) > 0:
                rateio = extras_liquidos * (float(item["valor_total_item"] or 0) / base_rateio)
            custo_total_entrada = max(0.0, float(item["valor_total_item"] or 0) + rateio)
            custo_efetivo = custo_total_entrada / quantidade_total if quantidade_total > 0 else 0.0

            # Verifica se o item já teve entrada antecipada pelo pedido de compra.
            movimento_existente = None
            if item["item_compra_id"]:
                cursor.execute("SELECT estoque_movimentacao_id FROM itens_compra WHERE id = ?", (item["item_compra_id"],))
                compra_item = cursor.fetchone()
                if compra_item and compra_item["estoque_movimentacao_id"]:
                    cursor.execute("SELECT * FROM movimentacoes_estoque WHERE id = ?", (compra_item["estoque_movimentacao_id"],))
                    movimento_existente = cursor.fetchone()
                    if movimento_existente and abs(float(movimento_existente["quantidade"] or 0) - quantidade_total) > 0.0001:
                        raise ValueError(
                            f"{item['descricao']}: a entrada antecipada possui quantidade diferente do documento. "
                            "Corrija antes de lançar para evitar duplicidade."
                        )

            saldo_anterior = 0.0
            cursor.execute("""
                SELECT COALESCE(SUM(CASE WHEN tipo='Entrada' THEN quantidade WHEN tipo='Saída' THEN -quantidade ELSE 0 END),0) AS saldo
                FROM movimentacoes_estoque WHERE produto_id = ?
            """, (produto_id,))
            saldo_anterior = float(cursor.fetchone()["saldo"] or 0)
            custo_medio_anterior = float(produto["custo_medio"] or 0)
            ultimo_anterior = float(produto["ultimo_custo"] or 0)
            base_saldo = max(0.0, saldo_anterior)
            denominador = base_saldo + quantidade_total
            custo_medio_novo = ((base_saldo * custo_medio_anterior) + custo_total_entrada) / denominador if denominador > 0 else custo_efetivo
            ultimo_novo = custo_efetivo if custo_efetivo > 0 else ultimo_anterior

            if movimento_existente:
                mov_id = movimento_existente["id"]
                reconciliado = 1
                cursor.execute("UPDATE movimentacoes_estoque SET nota_entrada_item_id = ? WHERE id = ?", (item["id"], mov_id))
            else:
                cursor.execute("""
                    INSERT INTO movimentacoes_estoque
                    (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_nome, fator_conversao,
                     data_movimentacao, origem, origem_id, item_compra_id, nota_entrada_item_id,
                     fornecedor_id, usuario, observacao)
                    VALUES (?, 'Entrada', ?, ?, ?, 1, ?, 'Nota de Entrada', ?, ?, ?, ?, ?, ?)
                """, (
                    produto_id, quantidade_total, quantidade_total, produto["unidade_padrao"], data_mov,
                    nota_id, item["item_compra_id"], item["id"], nota["fornecedor_id"], usuario,
                    f"Documento {nota['numero']}/{nota['serie'] or '-'} - {item['natureza']}. {item['observacao'] or ''}".strip()
                ))
                mov_id = cursor.lastrowid; reconciliado = 0

            cursor.execute("""
                UPDATE produtos_estoque
                SET custo_medio = ?, ultimo_custo = ?, data_ultimo_custo = ?
                WHERE id = ?
            """, (custo_medio_novo, ultimo_novo, data_mov, produto_id))
            cursor.execute("""
                UPDATE itens_nota_entrada
                SET custo_total_entrada = ?, custo_efetivo_unitario = ?,
                    custo_medio_anterior = ?, custo_medio_novo = ?,
                    ultimo_custo_anterior = ?, ultimo_custo_novo = ?,
                    movimentacao_estoque_id = ?, reconciliado_movimento_existente = ?
                WHERE id = ?
            """, (
                custo_total_entrada, custo_efetivo, custo_medio_anterior, custo_medio_novo,
                ultimo_anterior, ultimo_novo, mov_id, reconciliado, item["id"]
            ))
            if item["item_compra_id"]:
                cursor.execute("""
                    UPDATE itens_compra
                    SET produto_estoque_id = COALESCE(produto_estoque_id, ?),
                        estoque_movimentacao_id = COALESCE(estoque_movimentacao_id, ?),
                        nota_entrada_item_id = ?
                    WHERE id = ?
                """, (produto_id, mov_id, item["id"], item["item_compra_id"]))

        cursor.execute("""
            UPDATE notas_entrada SET status='Lançada', lancado_por=?, data_lancamento=? WHERE id=?
        """, (usuario, agora, nota_id))
        if nota["pedido_compra_id"]:
            cursor.execute("UPDATE pedidos_compra SET status='Comprado' WHERE id = ? AND status != 'Cancelado'", (nota["pedido_compra_id"],))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()


# ============================================================
# DASHBOARDS E INDICADORES GERENCIAIS - V3.3
# ============================================================

def _dashboard_parse_data(valor):
    """Converte datas usadas pelos módulos para datetime, sem quebrar dados antigos."""
    texto = str(valor or "").strip()
    if not texto:
        return None
    for formato in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            continue
    return None


def _dashboard_periodo(data_inicio="", data_fim=""):
    hoje = datetime.now().date()
    inicio_padrao = hoje.replace(day=1)
    try:
        inicio = datetime.strptime(str(data_inicio or ""), "%Y-%m-%d").date()
    except ValueError:
        inicio = inicio_padrao
    try:
        fim = datetime.strptime(str(data_fim or ""), "%Y-%m-%d").date()
    except ValueError:
        fim = hoje
    if fim < inicio:
        inicio, fim = fim, inicio
    return inicio, fim


def _dashboard_no_periodo(valor, inicio, fim):
    data = _dashboard_parse_data(valor)
    return bool(data and inicio <= data.date() <= fim)


def _dashboard_custo_unitario(produto):
    custo_medio = float(produto.get("custo_medio", 0) or 0)
    custo_padrao = float(produto.get("custo_padrao", 0) or 0)
    ultimo_custo = float(produto.get("ultimo_custo", 0) or 0)
    if custo_medio > 0:
        return custo_medio
    if custo_padrao > 0:
        return custo_padrao
    return ultimo_custo


def _dashboard_barras(registros, campo="valor", limite=6):
    registros = sorted(registros, key=lambda x: float(x.get(campo, 0) or 0), reverse=True)[:limite]
    maior = max([float(item.get(campo, 0) or 0) for item in registros] or [0])
    for item in registros:
        valor = float(item.get(campo, 0) or 0)
        item["percentual_barra"] = round((valor / maior) * 100, 2) if maior > 0 else 0
    return registros


def carregar_indicadores_gerenciais(data_inicio="", data_fim=""):
    """Reúne indicadores sem movimentar dados ou alterar o histórico operacional."""
    inicio, fim = _dashboard_periodo(data_inicio, data_fim)
    conn = conectar()
    cursor = conn.cursor()

    # Saúde do estoque e valor estimado atual.
    saldos = carregar_saldo_estoque()
    estoque_status = {"OK": 0, "Baixo": 0, "Zerado": 0, "Negativo": 0}
    valor_estoque = 0.0
    produtos_alerta = []
    for item in saldos:
        status = item.get("status_saldo", "OK")
        estoque_status[status] = estoque_status.get(status, 0) + 1
        saldo = float(item.get("saldo_atual", 0) or 0)
        custo = _dashboard_custo_unitario(item)
        if saldo > 0 and custo > 0:
            valor_estoque += saldo * custo
        if status != "OK":
            produtos_alerta.append({
                "id": item["id"],
                "nome": item["nome"],
                "codigo": item.get("codigo") or "-",
                "saldo": saldo,
                "unidade": item.get("unidade_padrao") or "",
                "status": status,
            })

    # Validades calculadas pela mesma regra da tela operacional.
    validades = contar_lotes_validade_resumo()

    # Compras e notas.
    cursor.execute("SELECT id, status, valor_total, data_entrada FROM notas_entrada")
    notas = [dict(row) for row in cursor.fetchall()]
    notas_lancadas_periodo = [n for n in notas if n["status"] == "Lançada" and _dashboard_no_periodo(n["data_entrada"], inicio, fim)]
    valor_notas_periodo = sum(float(n.get("valor_total", 0) or 0) for n in notas_lancadas_periodo)
    notas_rascunho = sum(1 for n in notas if n["status"] == "Rascunho")

    # Produção confirmada e ranking de produtos produzidos.
    cursor.execute("""
        SELECT pr.id, pr.data_confirmacao, pr.total_unidades, pr.total_tabuleiros,
               ipr.produto_estoque_id, ipr.sabor, ipr.quantidade_unidades, ipr.status
        FROM producoes_realizadas pr
        LEFT JOIN itens_producao_realizada ipr ON ipr.producao_realizada_id = pr.id
        WHERE pr.status = 'Confirmada'
    """)
    producao_rows = [dict(row) for row in cursor.fetchall()]
    producoes_ids = set()
    producao_unidades = 0.0
    producao_tabuleiros = 0.0
    ranking_producao = {}
    for row in producao_rows:
        if not _dashboard_no_periodo(row.get("data_confirmacao"), inicio, fim):
            continue
        if row["id"] not in producoes_ids:
            producoes_ids.add(row["id"])
            producao_unidades += float(row.get("total_unidades", 0) or 0)
            producao_tabuleiros += float(row.get("total_tabuleiros", 0) or 0)
        if row.get("status") != "Cancelado":
            nome = row.get("sabor") or "Produto sem nome"
            ranking_producao[nome] = ranking_producao.get(nome, 0) + float(row.get("quantidade_unidades", 0) or 0)

    top_producao = _dashboard_barras([
        {"nome": nome, "valor": valor}
        for nome, valor in ranking_producao.items()
    ], limite=7)

    # Consumos e perdas confirmados no período.
    cursor.execute("""
        SELECT cr.data_confirmacao, i.quantidade_utilizada, i.quantidade_perda,
               i.produto_id, p.nome AS produto_nome, p.custo_medio, p.custo_padrao, p.ultimo_custo
        FROM consumos_reais_producao cr
        JOIN itens_consumo_real_producao i ON i.consumo_id = cr.id
        JOIN produtos_estoque p ON p.id = i.produto_id
        WHERE cr.status = 'Confirmado'
    """)
    perdas_registros = 0
    custo_perdas_estimado = 0.0
    ranking_perdas = {}
    for row in cursor.fetchall():
        item = dict(row)
        if not _dashboard_no_periodo(item.get("data_confirmacao"), inicio, fim):
            continue
        perda = float(item.get("quantidade_perda", 0) or 0)
        if perda <= 0:
            continue
        perdas_registros += 1
        custo = _dashboard_custo_unitario(item)
        custo_perdas_estimado += perda * custo
        nome = item.get("produto_nome") or "Produto"
        ranking_perdas[nome] = ranking_perdas.get(nome, 0) + perda * custo
    top_perdas = _dashboard_barras([
        {"nome": nome, "valor": valor}
        for nome, valor in ranking_perdas.items()
    ], limite=5)

    # Expedição e custo congelado por loja.
    cursor.execute("""
        SELECT r.id, r.status, r.data_saida, r.custo_total_snapshot, l.nome AS loja_nome
        FROM romaneios_loja r
        JOIN lojas l ON l.id = r.loja_id
    """)
    romaneios = [dict(row) for row in cursor.fetchall()]
    custo_lojas = {}
    romaneios_periodo = 0
    custo_expedido_periodo = 0.0
    for row in romaneios:
        if row.get("status") not in ["Em rota", "Finalizado"]:
            continue
        if not _dashboard_no_periodo(row.get("data_saida"), inicio, fim):
            continue
        valor = float(row.get("custo_total_snapshot", 0) or 0)
        romaneios_periodo += 1
        custo_expedido_periodo += valor
        loja = row.get("loja_nome") or "Loja"
        custo_lojas[loja] = custo_lojas.get(loja, 0) + valor
    top_lojas = _dashboard_barras([
        {"nome": nome, "valor": valor}
        for nome, valor in custo_lojas.items()
    ], limite=7)

    # Custos confirmados por centros de custo.
    cursor.execute("""
        SELECT s.data_confirmacao, s.custo_total, c.nome AS centro_nome
        FROM solicitacoes_internas s
        LEFT JOIN centros_custo c ON c.id = s.centro_custo_id
        WHERE s.status IN ('Atendida', 'Atendida parcialmente') AND s.destino_tipo = 'Centro de Custo'
    """)
    custo_centros_periodo = 0.0
    centros_ranking = {}
    for row in cursor.fetchall():
        item = dict(row)
        if not _dashboard_no_periodo(item.get("data_confirmacao"), inicio, fim):
            continue
        valor = float(item.get("custo_total", 0) or 0)
        custo_centros_periodo += valor
        nome = item.get("centro_nome") or "Sem centro"
        centros_ranking[nome] = centros_ranking.get(nome, 0) + valor
    top_centros = _dashboard_barras([
        {"nome": nome, "valor": valor}
        for nome, valor in centros_ranking.items()
    ], limite=7)

    # Status de pedidos e romaneios.
    cursor.execute("SELECT status, COUNT(*) AS total FROM pedidos_loja GROUP BY status")
    pedidos_status = {row["status"]: int(row["total"] or 0) for row in cursor.fetchall()}
    cursor.execute("SELECT status, COUNT(*) AS total FROM romaneios_loja GROUP BY status")
    romaneios_status = {row["status"]: int(row["total"] or 0) for row in cursor.fetchall()}

    # Alertas operacionais gerais.
    cursor.execute("SELECT COUNT(*) AS total FROM separacoes_estoque WHERE status NOT IN ('Separado', 'Cancelado', 'Encerrado')")
    separacoes_pendentes = int(cursor.fetchone()["total"] or 0)
    cursor.execute("SELECT COUNT(*) AS total FROM transferencias_internas WHERE status = 'Rascunho'")
    transferencias_rascunho = int(cursor.fetchone()["total"] or 0)
    cursor.execute("SELECT COUNT(*) AS total FROM rodadas_pedidos_loja WHERE status = 'Aberta'")
    rodadas_abertas = int(cursor.fetchone()["total"] or 0)

    conn.close()

    total_produtos = len(saldos)
    total_alertas_estoque = estoque_status.get("Baixo", 0) + estoque_status.get("Zerado", 0) + estoque_status.get("Negativo", 0)
    percentual_estoque_ok = round((estoque_status.get("OK", 0) / total_produtos) * 100, 1) if total_produtos else 0

    alertas_operacionais = [
        {"nivel": "critico", "titulo": "Lotes vencidos", "valor": validades.get("Vencido", 0), "url": "/validade?status=Vencido"},
        {"nivel": "critico", "titulo": "Estoques negativos", "valor": estoque_status.get("Negativo", 0), "url": "/estoque/saldo?status=Negativo"},
        {"nivel": "atencao", "titulo": "Produtos baixos ou zerados", "valor": estoque_status.get("Baixo", 0) + estoque_status.get("Zerado", 0), "url": "/estoque/saldo"},
        {"nivel": "atencao", "titulo": "Validades próximas", "valor": validades.get("Próximo", 0) + validades.get("Crítico", 0) + validades.get("Vence hoje", 0), "url": "/validade"},
        {"nivel": "info", "titulo": "Separações pendentes", "valor": separacoes_pendentes, "url": "/estoque/separacoes"},
        {"nivel": "info", "titulo": "Notas em rascunho", "valor": notas_rascunho, "url": "/compras/notas"},
    ]
    alertas_operacionais = [item for item in alertas_operacionais if item["valor"] > 0]

    return {
        "periodo": {
            "inicio": inicio.strftime("%Y-%m-%d"),
            "fim": fim.strftime("%Y-%m-%d"),
            "inicio_br": inicio.strftime("%d/%m/%Y"),
            "fim_br": fim.strftime("%d/%m/%Y"),
        },
        "kpis": {
            "pedidos_em_andamento": contar_pedidos_loja_pendentes(),
            "compras_pendentes": contar_pedidos_pendentes(),
            "solicitacoes_pendentes": contar_solicitacoes_internas_pendentes(),
            "pre_producoes": contar_pre_producoes_pendentes(),
            "planejamentos_abertos": contar_planejamentos_abertos(),
            "separacoes_pendentes": separacoes_pendentes,
            "transferencias_rascunho": transferencias_rascunho,
            "rodadas_abertas": rodadas_abertas,
            "notas_rascunho": notas_rascunho,
            "valor_notas": valor_notas_periodo,
            "notas_lancadas": len(notas_lancadas_periodo),
            "producoes_confirmadas": len(producoes_ids),
            "producao_unidades": producao_unidades,
            "producao_tabuleiros": producao_tabuleiros,
            "romaneios_periodo": romaneios_periodo,
            "custo_expedido": custo_expedido_periodo,
            "custo_centros": custo_centros_periodo,
            "perdas_registros": perdas_registros,
            "custo_perdas_estimado": custo_perdas_estimado,
            "valor_estoque_estimado": valor_estoque,
            "alertas_estoque": total_alertas_estoque,
            "alertas_validade": validades.get("alertas", 0),
        },
        "estoque": {
            "status": estoque_status,
            "total_produtos": total_produtos,
            "percentual_ok": percentual_estoque_ok,
            "produtos_alerta": produtos_alerta[:8],
        },
        "validades": validades,
        "pedidos_status": pedidos_status,
        "romaneios_status": romaneios_status,
        "top_lojas": top_lojas,
        "top_centros": top_centros,
        "top_producao": top_producao,
        "top_perdas": top_perdas,
        "alertas_operacionais": alertas_operacionais,
    }


# ============================================================
# DASHBOARD / NAVEGAÇÃO
# ============================================================

@app.route("/dashboard", methods=["GET"])
@login_obrigatorio
def dashboard():
    perfil = session.get("perfil")
    if perfil == "loja":
        return redirect(url_for("loja_painel"))

    permissoes = {
        "producao": perfil in ["admin", "producao", "estoque"],
        "compras": perfil in ["admin", "compras", "estoque"],
        "solicitar_compras": perfil in ["admin", "estoque"],
        "estoque": perfil in ["admin", "estoque", "compras"],
        "administracao": perfil == "admin",
        "relatorios_compras": perfil in ["admin", "compras"],
        "confirmar_compras": perfil in ["admin", "compras"],
        "expedicao": perfil in ["admin", "estoque"],
        "custos": perfil == "admin",
        "indicadores": perfil == "admin",
    }

    usuario_solicitacoes = None if perfil in ["admin", "estoque"] else session.get("usuario")
    indicadores = carregar_indicadores_gerenciais()

    return render_template(
        "dashboard.html",
        permissoes=permissoes,
        indicadores=indicadores,
        total_pendentes=contar_pedidos_pendentes(),
        total_pre_producoes=contar_pre_producoes_pendentes(),
        total_solicitacoes_internas=contar_solicitacoes_internas_pendentes(usuario_solicitacoes),
        total_pedidos_loja=contar_pedidos_loja_pendentes()
    )


@app.route("/indicadores", methods=["GET"])
@admin_obrigatorio
def indicadores_gerenciais():
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()
    indicadores = carregar_indicadores_gerenciais(data_inicio, data_fim)
    return render_template("indicadores.html", indicadores=indicadores)



# ============================================================
# ROTAS DE SOLICITAÇÕES INTERNAS - V2.7
# ============================================================

@app.route("/solicitacoes-internas", methods=["GET"])
@login_obrigatorio
def solicitacoes_internas():
    perfil = session.get("perfil")
    somente_usuario = None if perfil in ["admin", "estoque"] else session.get("usuario")

    status = str(request.args.get("status", "") or "").strip()
    destino_tipo = str(request.args.get("destino_tipo", "") or "").strip()
    centro_custo_id = str(request.args.get("centro_custo_id", "") or "").strip()
    celula_id = str(request.args.get("celula_id", "") or "").strip()
    busca = str(request.args.get("busca", "") or "").strip()

    return render_template(
        "solicitacoes_internas.html",
        solicitacoes=carregar_solicitacoes_internas(
            status=status,
            destino_tipo=destino_tipo,
            centro_custo_id=centro_custo_id,
            celula_id=celula_id,
            busca=busca,
            somente_usuario=somente_usuario,
        ),
        centros_custo=carregar_centros_custo(),
        celulas=carregar_celulas_producao(),
        status_filtro=STATUS_SOLICITACAO_INTERNA,
        status_selecionado=status,
        destino_tipo_selecionado=destino_tipo,
        centro_custo_selecionado=centro_custo_id,
        celula_selecionada=celula_id,
        busca=busca,
        pode_operar=perfil in ["admin", "estoque"],
    )


@app.route("/solicitacoes-internas/nova", methods=["GET", "POST"])
@login_obrigatorio
def solicitacao_interna_nova():
    mensagem_erro = None

    if request.method == "POST":
        try:
            solicitacao_id = criar_solicitacao_interna(request.form)
            return redirect(url_for("solicitacao_interna_detalhes", solicitacao_id=solicitacao_id))
        except (ValueError, sqlite3.IntegrityError) as erro:
            mensagem_erro = str(erro)

    return render_template(
        "solicitacao_interna_nova.html",
        produtos=carregar_produtos_estoque(),
        embalagens_por_produto=carregar_embalagens_todos_produtos(),
        centros_custo=carregar_centros_custo(),
        celulas=carregar_celulas_producao(),
        prioridades=PRIORIDADES_SOLICITACAO,
        mensagem_erro=mensagem_erro,
        dados_form=request.form if request.method == "POST" else {},
        data_hoje=datetime.now().strftime("%Y-%m-%d"),
    )


@app.route("/solicitacoes-internas/<int:solicitacao_id>", methods=["GET"])
@login_obrigatorio
def solicitacao_interna_detalhes(solicitacao_id):
    solicitacao = buscar_solicitacao_interna(solicitacao_id)
    if not solicitacao:
        return render_template(
            "acesso_negado.html",
            mensagem="Solicitação interna não encontrada."
        ), 404

    perfil = session.get("perfil")
    usuario = session.get("usuario")
    pode_operar = perfil in ["admin", "estoque"]
    if not pode_operar and solicitacao["solicitante"] != usuario:
        return render_template(
            "acesso_negado.html",
            mensagem="Você não tem permissão para visualizar esta solicitação."
        )

    return render_template(
        "solicitacao_interna_detalhes.html",
        solicitacao=solicitacao,
        itens=carregar_itens_solicitacao_interna(solicitacao_id),
        pode_operar=pode_operar,
        pode_cancelar=(
            solicitacao["status"] not in ["Atendida", "Atendida parcialmente", "Cancelada"]
            and (pode_operar or solicitacao["solicitante"] == usuario)
        ),
        mensagem_sucesso=request.args.get("sucesso"),
        mensagem_erro=request.args.get("erro"),
    )


@app.route("/solicitacoes-internas/<int:solicitacao_id>/salvar-separacao", methods=["POST"])
@operacao_estoque_obrigatorio
def solicitacao_interna_salvar_separacao(solicitacao_id):
    try:
        salvar_separacao_solicitacao_interna(solicitacao_id, request.form)
        return redirect(url_for(
            "solicitacao_interna_detalhes",
            solicitacao_id=solicitacao_id,
            sucesso="Separação atualizada com sucesso."
        ))
    except ValueError as erro:
        return redirect(url_for(
            "solicitacao_interna_detalhes",
            solicitacao_id=solicitacao_id,
            erro=str(erro)
        ))


@app.route("/solicitacoes-internas/<int:solicitacao_id>/confirmar", methods=["POST"])
@operacao_estoque_obrigatorio
def solicitacao_interna_confirmar(solicitacao_id):
    try:
        confirmar_solicitacao_interna(solicitacao_id)
        return redirect(url_for(
            "solicitacao_interna_detalhes",
            solicitacao_id=solicitacao_id,
            sucesso="Atendimento confirmado e estoque movimentado com sucesso."
        ))
    except ValueError as erro:
        return redirect(url_for(
            "solicitacao_interna_detalhes",
            solicitacao_id=solicitacao_id,
            erro=str(erro)
        ))


@app.route("/solicitacoes-internas/<int:solicitacao_id>/cancelar", methods=["POST"])
@login_obrigatorio
def solicitacao_interna_cancelar(solicitacao_id):
    solicitacao = buscar_solicitacao_interna(solicitacao_id)
    if not solicitacao:
        return redirect(url_for("solicitacoes_internas"))

    perfil = session.get("perfil")
    usuario = session.get("usuario")
    if perfil not in ["admin", "estoque"] and solicitacao["solicitante"] != usuario:
        return render_template(
            "acesso_negado.html",
            mensagem="Você não tem permissão para cancelar esta solicitação."
        )

    try:
        cancelar_solicitacao_interna(
            solicitacao_id,
            request.form.get("motivo_cancelamento", "")
        )
        return redirect(url_for(
            "solicitacao_interna_detalhes",
            solicitacao_id=solicitacao_id,
            sucesso="Solicitação cancelada."
        ))
    except ValueError as erro:
        return redirect(url_for(
            "solicitacao_interna_detalhes",
            solicitacao_id=solicitacao_id,
            erro=str(erro)
        ))


@app.route("/solicitacoes-internas/centros-custo", methods=["GET", "POST"])
@operacao_estoque_obrigatorio
def centros_custo():
    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = str(request.form.get("nome", "") or "").strip()
        descricao = str(request.form.get("descricao", "") or "").strip()

        if not nome:
            mensagem_erro = "Informe o nome do centro de custo."
        else:
            try:
                conn = conectar()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO centros_custo
                    (nome, descricao, ativo, data_cadastro, criado_por)
                    VALUES (?, ?, 1, ?, ?)
                """, (nome, descricao, _agora_texto(), _usuario_atual()))
                conn.commit()
                conn.close()
                mensagem_sucesso = "Centro de custo criado com sucesso."
            except sqlite3.IntegrityError:
                mensagem_erro = "Já existe um centro de custo com esse nome."

    return render_template(
        "centros_custo.html",
        centros=carregar_centros_custo(False),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro,
    )


@app.route("/solicitacoes-internas/centros-custo/<int:centro_custo_id>/alternar", methods=["POST"])
@operacao_estoque_obrigatorio
def centro_custo_alternar(centro_custo_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE centros_custo
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (centro_custo_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("centros_custo"))


@app.route("/solicitacoes-internas/relatorio", methods=["GET"])
@operacao_estoque_obrigatorio
def solicitacoes_internas_relatorio():
    data_inicio_html = str(request.args.get("data_inicio", "") or "").strip()
    data_fim_html = str(request.args.get("data_fim", "") or "").strip()
    centro_custo_id = str(request.args.get("centro_custo_id", "") or "").strip()

    data_inicio = _data_html_para_br(data_inicio_html)
    data_fim = _data_html_para_br(data_fim_html)
    resumo, detalhes, total_geral = carregar_relatorio_centros_custo(
        data_inicio=data_inicio,
        data_fim=data_fim,
        centro_custo_id=centro_custo_id,
    )

    return render_template(
        "solicitacoes_internas_relatorio.html",
        resumo=resumo,
        detalhes=detalhes,
        total_geral=total_geral,
        centros_custo=carregar_centros_custo(False),
        data_inicio=data_inicio_html,
        data_fim=data_fim_html,
        centro_custo_selecionado=centro_custo_id,
    )



# ============================================================
# CONSUMO REAL DAS CÉLULAS - V2.4
# ============================================================

def _numero_nao_negativo(valor, nome_campo="Quantidade"):
    try:
        numero = float(str(valor or 0).replace(",", "."))
    except Exception:
        raise ValueError(f"{nome_campo} inválida.")

    if numero < 0:
        raise ValueError(f"{nome_campo} não pode ser negativa.")
    return numero


def contar_consumos_reais_rascunho():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) AS total FROM consumos_reais_producao WHERE status = 'Rascunho'")
    total = int(cursor.fetchone()["total"] or 0)
    conn.close()
    return total


def carregar_consumos_reais(status="", celula_id="", producao_realizada_id=""):
    filtros = []
    parametros = []

    if str(status or "").strip():
        filtros.append("crp.status = ?")
        parametros.append(str(status).strip())

    if str(celula_id or "").strip():
        filtros.append("crp.celula_id = ?")
        parametros.append(int(celula_id))

    if str(producao_realizada_id or "").strip():
        filtros.append("crp.producao_realizada_id = ?")
        parametros.append(int(producao_realizada_id))

    where = f"WHERE {' AND '.join(filtros)}" if filtros else ""

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT
            crp.*,
            cp.nome AS celula_nome,
            cp.centro_custo,
            pr.status AS producao_status,
            pr.data_criacao AS producao_data,
            COUNT(icrp.id) AS total_itens,
            COALESCE(SUM(icrp.quantidade_utilizada), 0) AS total_utilizado,
            COALESCE(SUM(icrp.quantidade_perda), 0) AS total_perda,
            COALESCE(SUM(icrp.quantidade_devolvida), 0) AS total_devolvido
        FROM consumos_reais_producao crp
        INNER JOIN celulas_producao cp ON cp.id = crp.celula_id
        INNER JOIN producoes_realizadas pr ON pr.id = crp.producao_realizada_id
        LEFT JOIN itens_consumo_real_producao icrp ON icrp.consumo_id = crp.id
        {where}
        GROUP BY crp.id
        ORDER BY crp.id DESC
    """, parametros)
    registros = cursor.fetchall()
    conn.close()
    return registros


def carregar_consumos_reais_producao(producao_realizada_id):
    return carregar_consumos_reais(producao_realizada_id=producao_realizada_id)


def buscar_consumo_real_producao(consumo_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            crp.*,
            cp.nome AS celula_nome,
            cp.centro_custo,
            pr.status AS producao_status,
            pr.data_criacao AS producao_data,
            pr.cenario,
            pr.dia_semana,
            pr.total_unidades,
            pr.total_tabuleiros
        FROM consumos_reais_producao crp
        INNER JOIN celulas_producao cp ON cp.id = crp.celula_id
        INNER JOIN producoes_realizadas pr ON pr.id = crp.producao_realizada_id
        WHERE crp.id = ?
    """, (consumo_id,))
    consumo = cursor.fetchone()
    conn.close()
    return consumo


def _previstos_planejamento_por_produto(planejamento_id):
    if not planejamento_id:
        return {}

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            cpp.produto_id,
            cpp.quantidade,
            cpp.unidade,
            pe.unidade_padrao
        FROM consumos_previstos_planejamento cpp
        INNER JOIN produtos_estoque pe ON pe.id = cpp.produto_id
        WHERE cpp.planejamento_id = ?
    """, (planejamento_id,))
    linhas = cursor.fetchall()
    conn.close()

    totais = {}
    for linha in linhas:
        produto_id = int(linha["produto_id"])
        try:
            quantidade = converter_quantidade_planejamento(
                linha["quantidade"],
                linha["unidade"],
                linha["unidade_padrao"]
            )
        except Exception:
            if normalizar_unidade_planejamento(linha["unidade"]) != normalizar_unidade_planejamento(linha["unidade_padrao"]):
                continue
            quantidade = float(linha["quantidade"] or 0)

        totais[produto_id] = totais.get(produto_id, 0.0) + float(quantidade or 0)
    return totais


def _recebidos_planejamento_celula(planejamento_id, celula_id):
    if not planejamento_id:
        return {}

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            iti.produto_id,
            COALESCE(SUM(iti.quantidade), 0) AS quantidade
        FROM transferencias_internas ti
        INNER JOIN itens_transferencia_interna iti ON iti.transferencia_id = ti.id
        INNER JOIN separacoes_estoque se ON se.id = ti.separacao_id
        WHERE se.planejamento_id = ?
          AND ti.celula_destino_id = ?
          AND ti.tipo = 'Estoque para Célula'
          AND ti.status = 'Confirmada'
        GROUP BY iti.produto_id
    """, (planejamento_id, celula_id))
    recebidos = {int(row["produto_id"]): float(row["quantidade"] or 0) for row in cursor.fetchall()}
    conn.close()
    return recebidos


def criar_consumo_real_producao(producao_realizada_id, celula_id, observacao=""):
    producao = buscar_producao_realizada(producao_realizada_id)
    if not producao:
        raise ValueError("Produção realizada não encontrada.")
    if producao["status"] == "Cancelado":
        raise ValueError("Não é possível registrar consumo para uma produção cancelada.")

    celula = buscar_celula_producao(celula_id)
    if not celula or not int(celula["ativo"] or 0):
        raise ValueError("Célula de produção inválida ou inativa.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id FROM consumos_reais_producao
        WHERE producao_realizada_id = ? AND celula_id = ?
    """, (producao_realizada_id, celula_id))
    existente = cursor.fetchone()
    conn.close()
    if existente:
        return int(existente["id"])

    planejamento_id = producao["planejamento_id"]
    previstos = _previstos_planejamento_por_produto(planejamento_id)
    recebidos = _recebidos_planejamento_celula(planejamento_id, celula_id)

    # Prioriza os produtos efetivamente enviados para a célula. Se ainda não houve
    # transferência, usa apenas produtos previstos que já tenham saldo nessa célula.
    candidatos = set(recebidos.keys())
    if not candidatos:
        for produto_id in previstos:
            if obter_saldo_celula_produto(celula_id, produto_id) > 0:
                candidatos.add(produto_id)

    if not candidatos:
        raise ValueError(
            "Esta célula ainda não recebeu itens deste planejamento e não possui saldo dos componentes previstos. "
            "Confirme primeiro a transferência para a célula ou adicione o item manualmente após criar um saldo."
        )

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN IMMEDIATE")
        cursor.execute("""
            INSERT INTO consumos_reais_producao
            (producao_realizada_id, planejamento_id, celula_id, status, data_criacao, criado_por, observacao)
            VALUES (?, ?, ?, 'Rascunho', ?, ?, ?)
        """, (
            producao_realizada_id,
            planejamento_id,
            celula_id,
            _agora_texto(),
            _usuario_atual(),
            str(observacao or "").strip()
        ))
        consumo_id = cursor.lastrowid

        for produto_id in sorted(candidatos):
            cursor.execute("SELECT unidade_padrao FROM produtos_estoque WHERE id = ?", (produto_id,))
            produto = cursor.fetchone()
            if not produto:
                continue
            saldo = _saldo_celula_conn(cursor, celula_id, produto_id)
            cursor.execute("""
                INSERT INTO itens_consumo_real_producao
                (consumo_id, produto_id, origem, quantidade_prevista, quantidade_recebida,
                 saldo_celula_inicial, unidade, observacao)
                VALUES (?, ?, 'Planejamento', ?, ?, ?, ?, ?)
            """, (
                consumo_id,
                produto_id,
                float(previstos.get(produto_id, 0)),
                float(recebidos.get(produto_id, 0)),
                saldo,
                produto["unidade_padrao"],
                "Item criado a partir do planejamento e das transferências confirmadas para a célula."
            ))

        conn.commit()
        return int(consumo_id)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def carregar_itens_consumo_real(consumo_id):
    consumo = buscar_consumo_real_producao(consumo_id)
    if not consumo:
        return []

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            icrp.*,
            pe.codigo AS produto_codigo,
            pe.nome AS produto_nome,
            ce.nome AS categoria_nome,
            mp.nome AS motivo_perda_nome
        FROM itens_consumo_real_producao icrp
        INNER JOIN produtos_estoque pe ON pe.id = icrp.produto_id
        LEFT JOIN categorias_estoque ce ON ce.id = pe.categoria_id
        LEFT JOIN motivos_perda mp ON mp.id = icrp.motivo_perda_id
        WHERE icrp.consumo_id = ?
        ORDER BY ce.nome, pe.nome
    """, (consumo_id,))
    linhas = cursor.fetchall()

    itens = []
    for linha in linhas:
        item = dict(linha)
        item["saldo_atual"] = _saldo_celula_conn(cursor, consumo["celula_id"], linha["produto_id"])
        comprometido = (
            float(linha["quantidade_utilizada"] or 0)
            + float(linha["quantidade_perda"] or 0)
            + float(linha["quantidade_devolvida"] or 0)
        )
        item["saldo_estimado"] = item["saldo_atual"] - comprometido
        itens.append(item)

    conn.close()
    return itens


def salvar_itens_consumo_real(consumo_id, form):
    consumo = buscar_consumo_real_producao(consumo_id)
    if not consumo:
        raise ValueError("Registro de consumo não encontrado.")
    if consumo["status"] != "Rascunho":
        raise ValueError("Somente consumos em rascunho podem ser alterados.")

    itens = carregar_itens_consumo_real(consumo_id)
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN IMMEDIATE")
        for item in itens:
            item_id = int(item["id"])
            utilizada = _numero_nao_negativo(form.get(f"utilizada_{item_id}"), "Quantidade utilizada")
            perda = _numero_nao_negativo(form.get(f"perda_{item_id}"), "Quantidade de perda")
            devolvida = _numero_nao_negativo(form.get(f"devolvida_{item_id}"), "Quantidade devolvida")
            observacao = str(form.get(f"observacao_{item_id}", "") or "").strip()
            motivo_perda_id = form.get(f"motivo_perda_id_{item_id}")
            observacao_perda = str(form.get(f"observacao_perda_{item_id}", "") or "").strip()
            motivo_perda = None
            if perda > 0:
                motivo_perda = buscar_motivo_perda_cursor(cursor, motivo_perda_id, "Produção")
                if not motivo_perda:
                    raise ValueError(f"Selecione o motivo da perda de {item['produto_nome']}.")
                if int(motivo_perda.get("exige_observacao") or 0) == 1 and not observacao_perda:
                    raise ValueError(f"Detalhe a perda de {item['produto_nome']}.")
            else:
                motivo_perda_id = None
                observacao_perda = ""

            cursor.execute("""
                UPDATE itens_consumo_real_producao
                SET quantidade_utilizada = ?, quantidade_perda = ?, quantidade_devolvida = ?,
                    observacao = ?, motivo_perda_id = ?, observacao_perda = ?
                WHERE id = ? AND consumo_id = ?
            """, (utilizada, perda, devolvida, observacao,
                  int(motivo_perda_id) if str(motivo_perda_id or "").isdigit() else None,
                  observacao_perda, item_id, consumo_id))

        observacao_geral = str(form.get("observacao_geral", consumo["observacao"] or "") or "").strip()
        cursor.execute("UPDATE consumos_reais_producao SET observacao = ? WHERE id = ?", (observacao_geral, consumo_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def adicionar_item_consumo_real(consumo_id, produto_id, observacao=""):
    consumo = buscar_consumo_real_producao(consumo_id)
    if not consumo:
        raise ValueError("Registro de consumo não encontrado.")
    if consumo["status"] != "Rascunho":
        raise ValueError("Somente consumos em rascunho podem receber itens.")

    produto = buscar_produto_estoque(produto_id)
    if not produto:
        raise ValueError("Produto oficial não encontrado.")

    saldo = obter_saldo_celula_produto(consumo["celula_id"], produto_id)
    if saldo <= 0:
        raise ValueError("O produto não possui saldo positivo nesta célula.")

    previstos = _previstos_planejamento_por_produto(consumo["planejamento_id"])
    recebidos = _recebidos_planejamento_celula(consumo["planejamento_id"], consumo["celula_id"])

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO itens_consumo_real_producao
            (consumo_id, produto_id, origem, quantidade_prevista, quantidade_recebida,
             saldo_celula_inicial, unidade, observacao)
            VALUES (?, ?, 'Manual', ?, ?, ?, ?, ?)
        """, (
            consumo_id,
            produto_id,
            float(previstos.get(int(produto_id), 0)),
            float(recebidos.get(int(produto_id), 0)),
            saldo,
            produto["unidade_padrao"],
            str(observacao or "").strip()
        ))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.rollback()
        raise ValueError("Este produto já está incluído no consumo.")
    finally:
        conn.close()


def remover_item_manual_consumo(consumo_id, item_id):
    consumo = buscar_consumo_real_producao(consumo_id)
    if not consumo or consumo["status"] != "Rascunho":
        raise ValueError("Somente consumos em rascunho podem ser alterados.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        DELETE FROM itens_consumo_real_producao
        WHERE id = ? AND consumo_id = ? AND origem = 'Manual'
    """, (item_id, consumo_id))
    if cursor.rowcount == 0:
        conn.close()
        raise ValueError("Somente itens adicionados manualmente podem ser removidos.")
    conn.commit()
    conn.close()


def confirmar_consumo_real_producao(consumo_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN IMMEDIATE")
        cursor.execute("""
            SELECT crp.*, pr.status AS producao_status
            FROM consumos_reais_producao crp
            INNER JOIN producoes_realizadas pr ON pr.id = crp.producao_realizada_id
            WHERE crp.id = ?
        """, (consumo_id,))
        consumo = cursor.fetchone()
        if not consumo:
            raise ValueError("Registro de consumo não encontrado.")
        if consumo["status"] != "Rascunho":
            raise ValueError("Este consumo já foi finalizado ou cancelado.")
        if consumo["producao_status"] != "Confirmado":
            raise ValueError("Confirme primeiro a produção realizada antes de baixar o consumo da célula.")

        cursor.execute("""
            SELECT icrp.*, pe.nome AS produto_nome, pe.unidade_padrao
            FROM itens_consumo_real_producao icrp
            INNER JOIN produtos_estoque pe ON pe.id = icrp.produto_id
            WHERE icrp.consumo_id = ?
            ORDER BY icrp.id
        """, (consumo_id,))
        itens = cursor.fetchall()
        if not itens:
            raise ValueError("O consumo não possui itens.")

        total_movimentado = 0.0
        for item in itens:
            utilizada = float(item["quantidade_utilizada"] or 0)
            perda = float(item["quantidade_perda"] or 0)
            devolvida = float(item["quantidade_devolvida"] or 0)
            total = utilizada + perda + devolvida
            if total <= 0:
                continue
            saldo = _saldo_celula_conn(cursor, consumo["celula_id"], item["produto_id"])
            if total > saldo + 0.0000001:
                raise ValueError(
                    f"Saldo insuficiente na célula para {item['produto_nome']}. "
                    f"Disponível: {saldo:.4f} {item['unidade_padrao']}; "
                    f"informado: {total:.4f}."
                )
            total_movimentado += total

        if total_movimentado <= 0:
            raise ValueError("Informe ao menos uma quantidade utilizada, perdida ou devolvida.")

        usuario = _usuario_atual()
        data_mov = datetime.now().strftime("%d/%m/%Y")
        agora = _agora_texto()

        itens_com_retorno = [item for item in itens if float(item["quantidade_devolvida"] or 0) > 0]
        transferencia_retorno_id = None
        if itens_com_retorno:
            cursor.execute("""
                INSERT INTO transferencias_internas
                (tipo, celula_origem_id, status, data_criacao, criado_por,
                 data_confirmacao, confirmado_por, observacao)
                VALUES ('Célula para Estoque', ?, 'Confirmada', ?, ?, ?, ?, ?)
            """, (
                consumo["celula_id"], agora, usuario, agora, usuario,
                f"Retorno automático registrado no consumo real #{consumo_id}."
            ))
            transferencia_retorno_id = cursor.lastrowid

        for item in itens:
            utilizada = float(item["quantidade_utilizada"] or 0)
            perda = float(item["quantidade_perda"] or 0)
            devolvida = float(item["quantidade_devolvida"] or 0)
            movimento_uso_id = None
            movimento_perda_id = None

            if utilizada > 0:
                cursor.execute("""
                    INSERT INTO movimentacoes_celula
                    (celula_id, produto_id, tipo, quantidade, data_movimentacao,
                     origem, origem_id, usuario, observacao)
                    VALUES (?, ?, 'Saída', ?, ?, 'Consumo de Produção', ?, ?, ?)
                """, (
                    consumo["celula_id"], item["produto_id"], utilizada, data_mov,
                    consumo_id, usuario,
                    item["observacao"] or f"Consumo real da produção #{consumo['producao_realizada_id']}."
                ))
                movimento_uso_id = cursor.lastrowid

            if perda > 0:
                cursor.execute("""
                    INSERT INTO movimentacoes_celula
                    (celula_id, produto_id, tipo, quantidade, data_movimentacao,
                     origem, origem_id, usuario, observacao)
                    VALUES (?, ?, 'Saída', ?, ?, 'Perda de Produção', ?, ?, ?)
                """, (
                    consumo["celula_id"], item["produto_id"], perda, data_mov,
                    consumo_id, usuario,
                    item["observacao"] or f"Perda registrada na produção #{consumo['producao_realizada_id']}."
                ))
                movimento_perda_id = cursor.lastrowid
                _registrar_perda_cursor(
                    cursor, "Produção", item["id"], item["produto_id"], perda, item["unidade_padrao"],
                    motivo_id=item["motivo_perda_id"],
                    observacao=item["observacao_perda"] or item["observacao"] or "",
                    consumo_id=consumo_id, item_consumo_id=item["id"],
                    celula_id=consumo["celula_id"], data_registro=agora, usuario=usuario
                )

            if devolvida > 0:
                cursor.execute("""
                    INSERT INTO itens_transferencia_interna
                    (transferencia_id, produto_id, quantidade, unidade, observacao)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    transferencia_retorno_id, item["produto_id"], devolvida,
                    item["unidade_padrao"],
                    item["observacao"] or f"Retorno do consumo real #{consumo_id}."
                ))
                item_transferencia_id = cursor.lastrowid

                cursor.execute("""
                    INSERT INTO movimentacoes_celula
                    (celula_id, produto_id, tipo, quantidade, data_movimentacao,
                     origem, origem_id, usuario, observacao)
                    VALUES (?, ?, 'Saída', ?, ?, 'Retorno ao Estoque', ?, ?, ?)
                """, (
                    consumo["celula_id"], item["produto_id"], devolvida, data_mov,
                    transferencia_retorno_id, usuario,
                    item["observacao"] or f"Retorno ao estoque pelo consumo real #{consumo_id}."
                ))
                mov_celula_retorno = cursor.lastrowid

                cursor.execute("""
                    INSERT INTO movimentacoes_estoque
                    (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_nome,
                     fator_conversao, data_movimentacao, origem, origem_id, usuario, observacao)
                    VALUES (?, 'Entrada', ?, ?, ?, 1, ?, 'Retorno de Célula', ?, ?, ?)
                """, (
                    item["produto_id"], devolvida, devolvida, item["unidade_padrao"],
                    data_mov, transferencia_retorno_id, usuario,
                    item["observacao"] or f"Retorno do consumo real #{consumo_id}."
                ))
                mov_estoque_retorno = cursor.lastrowid

                cursor.execute("""
                    UPDATE itens_transferencia_interna
                    SET estoque_movimentacao_id = ?, celula_movimentacao_id = ?
                    WHERE id = ?
                """, (mov_estoque_retorno, mov_celula_retorno, item_transferencia_id))

            saldo_final = _saldo_celula_conn(cursor, consumo["celula_id"], item["produto_id"])
            cursor.execute("""
                UPDATE itens_consumo_real_producao
                SET movimento_uso_id = ?, movimento_perda_id = ?,
                    transferencia_retorno_id = ?, saldo_celula_final = ?
                WHERE id = ?
            """, (
                movimento_uso_id,
                movimento_perda_id,
                transferencia_retorno_id if devolvida > 0 else None,
                saldo_final,
                item["id"]
            ))

        cursor.execute("""
            UPDATE consumos_reais_producao
            SET status = 'Confirmado', data_confirmacao = ?, confirmado_por = ?
            WHERE id = ?
        """, (agora, usuario, consumo_id))

        cursor.execute("""
            UPDATE producoes_realizadas
            SET observacao = CASE
                WHEN observacao IS NULL OR observacao = '' THEN ?
                ELSE observacao || ' | ' || ?
            END
            WHERE id = ?
        """, (
            f"Consumo real #{consumo_id} confirmado na célula em {agora}.",
            f"Consumo real #{consumo_id} confirmado na célula em {agora}.",
            consumo["producao_realizada_id"]
        ))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def cancelar_consumo_real_producao(consumo_id):
    consumo = buscar_consumo_real_producao(consumo_id)
    if not consumo:
        raise ValueError("Registro de consumo não encontrado.")
    if consumo["status"] != "Rascunho":
        raise ValueError("Somente consumos em rascunho podem ser cancelados.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE consumos_reais_producao
        SET status = 'Cancelado', data_cancelamento = ?, cancelado_por = ?
        WHERE id = ?
    """, (_agora_texto(), _usuario_atual(), consumo_id))
    conn.commit()
    conn.close()


def resumo_consumo_real(itens):
    # Não soma kg, unidades e litros no mesmo total. O resumo é separado
    # por unidade para preservar o significado operacional dos números.
    grupos = {}
    for item in itens:
        unidade = str(item.get("unidade") or "-")
        grupo = grupos.setdefault(unidade, {
            "unidade": unidade,
            "previsto": 0.0,
            "recebido": 0.0,
            "utilizado": 0.0,
            "perda": 0.0,
            "devolvido": 0.0,
        })
        grupo["previsto"] += float(item.get("quantidade_prevista", 0) or 0)
        grupo["recebido"] += float(item.get("quantidade_recebida", 0) or 0)
        grupo["utilizado"] += float(item.get("quantidade_utilizada", 0) or 0)
        grupo["perda"] += float(item.get("quantidade_perda", 0) or 0)
        grupo["devolvido"] += float(item.get("quantidade_devolvida", 0) or 0)
    return list(grupos.values())



# ============================================================
# ANÁLISE PREVISTO X REALIZADO - V2.5
# ============================================================

def _chave_item_analise(produto_id, nome):
    if produto_id:
        return f"produto:{int(produto_id)}"
    return f"nome:{str(nome or '').strip().lower()}"


def _classificar_variacao(previsto, realizado, tolerancia_absoluta=0.0001):
    previsto = float(previsto or 0)
    realizado = float(realizado or 0)
    diferenca = realizado - previsto

    if previsto <= tolerancia_absoluta and realizado <= tolerancia_absoluta:
        return "Sem movimento"
    if previsto <= tolerancia_absoluta and realizado > tolerancia_absoluta:
        return "Não previsto"
    if previsto > tolerancia_absoluta and realizado <= tolerancia_absoluta:
        return "Não realizado"

    tolerancia = max(tolerancia_absoluta, abs(previsto) * 0.001)
    if abs(diferenca) <= tolerancia:
        return "Conforme"
    if diferenca > 0:
        return "Acima"
    return "Abaixo"


def _percentual_variacao(previsto, realizado):
    previsto = float(previsto or 0)
    realizado = float(realizado or 0)
    if abs(previsto) <= 0.0000001:
        return None
    return ((realizado - previsto) / previsto) * 100.0


def _resumo_por_unidade_analise(linhas, campos):
    grupos = {}
    for linha in linhas:
        unidade = str(linha.get("unidade") or "-")
        grupo = grupos.setdefault(unidade, {"unidade": unidade})
        for campo in campos:
            grupo[campo] = float(grupo.get(campo, 0) or 0) + float(linha.get(campo, 0) or 0)
    return list(grupos.values())


def contar_analises_disponiveis():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM producoes_realizadas
        WHERE planejamento_id IS NOT NULL
          AND status = 'Confirmado'
    """)
    total = int(cursor.fetchone()["total"] or 0)
    conn.close()
    return total


def carregar_producoes_para_analise(status=""):
    conn = conectar()
    cursor = conn.cursor()
    parametros = []
    filtro = "WHERE pr.planejamento_id IS NOT NULL"

    if status:
        filtro += " AND pr.status = ?"
        parametros.append(status)

    cursor.execute(f"""
        SELECT
            pr.*,
            pp.status AS planejamento_status,
            pp.total_unidades AS previsto_total_unidades,
            pp.total_tabuleiros AS previsto_total_tabuleiros,
            SUM(CASE WHEN crp.status = 'Confirmado' THEN 1 ELSE 0 END) AS consumos_confirmados,
            SUM(CASE WHEN crp.status = 'Rascunho' THEN 1 ELSE 0 END) AS consumos_rascunho,
            SUM(CASE WHEN crp.status = 'Cancelado' THEN 1 ELSE 0 END) AS consumos_cancelados
        FROM producoes_realizadas pr
        INNER JOIN planejamentos_producao pp ON pp.id = pr.planejamento_id
        LEFT JOIN consumos_reais_producao crp ON crp.producao_realizada_id = pr.id
        {filtro}
        GROUP BY pr.id
        ORDER BY pr.id DESC
    """, parametros)

    registros = []
    for row in cursor.fetchall():
        item = dict(row)
        confirmados = int(item.get("consumos_confirmados") or 0)
        rascunhos = int(item.get("consumos_rascunho") or 0)
        if item["status"] == "Cancelado":
            situacao = "Cancelada"
        elif item["status"] != "Confirmado":
            situacao = "Produção pendente"
        elif confirmados == 0 and rascunhos == 0:
            situacao = "Aguardando consumo"
        elif rascunhos > 0:
            situacao = "Consumo parcial"
        else:
            situacao = "Com consumo confirmado"
        item["situacao_analise"] = situacao
        registros.append(item)

    conn.close()
    return registros


def construir_analise_producao(producao_realizada_id):
    producao_row = buscar_producao_realizada(producao_realizada_id)
    if not producao_row:
        return None

    producao = dict(producao_row)
    planejamento_id = producao.get("planejamento_id")
    if not planejamento_id:
        return {
            "producao": producao,
            "planejamento": None,
            "produtos": [],
            "consumos": [],
            "resumo_produtos": [],
            "resumo_consumos": [],
            "consumos_confirmados": 0,
            "consumos_rascunho": 0,
            "situacao": "Sem planejamento",
            "indicadores": {}
        }

    planejamento_row = buscar_planejamento_producao(planejamento_id)
    planejamento = dict(planejamento_row) if planejamento_row else None

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            ipp.produto_final_id AS produto_id,
            COALESCE(pe.codigo, '') AS produto_codigo,
            COALESCE(pe.nome, ipp.sabor) AS produto_nome,
            COALESCE(pe.unidade_padrao, 'un') AS unidade,
            SUM(ipp.quantidade_unidades) AS quantidade_prevista
        FROM itens_planejamento_producao ipp
        LEFT JOIN produtos_estoque pe ON pe.id = ipp.produto_final_id
        WHERE ipp.planejamento_id = ?
          AND ipp.status != 'Cancelado'
        GROUP BY ipp.produto_final_id, COALESCE(pe.nome, ipp.sabor), COALESCE(pe.unidade_padrao, 'un')
        ORDER BY COALESCE(pe.nome, ipp.sabor)
    """, (planejamento_id,))
    produtos_previstos = cursor.fetchall()

    cursor.execute("""
        SELECT
            ipr.produto_estoque_id AS produto_id,
            COALESCE(pe.codigo, '') AS produto_codigo,
            COALESCE(pe.nome, ipr.sabor) AS produto_nome,
            COALESCE(pe.unidade_padrao, 'un') AS unidade,
            SUM(ipr.quantidade_unidades) AS quantidade_realizada
        FROM itens_producao_realizada ipr
        LEFT JOIN produtos_estoque pe ON pe.id = ipr.produto_estoque_id
        WHERE ipr.producao_realizada_id = ?
          AND ipr.status != 'Cancelado'
        GROUP BY ipr.produto_estoque_id, COALESCE(pe.nome, ipr.sabor), COALESCE(pe.unidade_padrao, 'un')
        ORDER BY COALESCE(pe.nome, ipr.sabor)
    """, (producao_realizada_id,))
    produtos_realizados = cursor.fetchall()

    mapa_produtos = {}
    for row in produtos_previstos:
        chave = _chave_item_analise(row["produto_id"], row["produto_nome"])
        mapa_produtos[chave] = {
            "produto_id": row["produto_id"],
            "produto_codigo": row["produto_codigo"],
            "produto_nome": row["produto_nome"],
            "unidade": row["unidade"],
            "previsto": float(row["quantidade_prevista"] or 0),
            "realizado": 0.0,
        }

    for row in produtos_realizados:
        chave = _chave_item_analise(row["produto_id"], row["produto_nome"])
        item = mapa_produtos.setdefault(chave, {
            "produto_id": row["produto_id"],
            "produto_codigo": row["produto_codigo"],
            "produto_nome": row["produto_nome"],
            "unidade": row["unidade"],
            "previsto": 0.0,
            "realizado": 0.0,
        })
        item["realizado"] += float(row["quantidade_realizada"] or 0)

    produtos = []
    for item in mapa_produtos.values():
        item["diferenca"] = item["realizado"] - item["previsto"]
        item["percentual"] = _percentual_variacao(item["previsto"], item["realizado"])
        item["status_variacao"] = _classificar_variacao(item["previsto"], item["realizado"])
        produtos.append(item)
    produtos.sort(key=lambda x: x["produto_nome"].lower())

    cursor.execute("""
        SELECT
            cpp.produto_id,
            pe.codigo AS produto_codigo,
            pe.nome AS produto_nome,
            ce.nome AS categoria_nome,
            cpp.unidade,
            GROUP_CONCAT(DISTINCT cpp.tipo) AS tipos,
            SUM(cpp.quantidade) AS quantidade_prevista
        FROM consumos_previstos_planejamento cpp
        INNER JOIN produtos_estoque pe ON pe.id = cpp.produto_id
        LEFT JOIN categorias_estoque ce ON ce.id = pe.categoria_id
        WHERE cpp.planejamento_id = ?
        GROUP BY cpp.produto_id, cpp.unidade
        ORDER BY ce.nome, pe.nome
    """, (planejamento_id,))
    consumos_previstos = cursor.fetchall()

    cursor.execute("""
        SELECT
            icrp.produto_id,
            pe.codigo AS produto_codigo,
            pe.nome AS produto_nome,
            ce.nome AS categoria_nome,
            icrp.unidade,
            SUM(icrp.quantidade_recebida) AS quantidade_recebida,
            SUM(icrp.quantidade_utilizada) AS quantidade_utilizada,
            SUM(icrp.quantidade_perda) AS quantidade_perda,
            SUM(icrp.quantidade_devolvida) AS quantidade_devolvida
        FROM itens_consumo_real_producao icrp
        INNER JOIN consumos_reais_producao crp ON crp.id = icrp.consumo_id
        INNER JOIN produtos_estoque pe ON pe.id = icrp.produto_id
        LEFT JOIN categorias_estoque ce ON ce.id = pe.categoria_id
        WHERE crp.producao_realizada_id = ?
          AND crp.status = 'Confirmado'
        GROUP BY icrp.produto_id, icrp.unidade
        ORDER BY ce.nome, pe.nome
    """, (producao_realizada_id,))
    consumos_realizados = cursor.fetchall()

    mapa_consumos = {}
    for row in consumos_previstos:
        chave = f"{row['produto_id']}:{row['unidade']}"
        mapa_consumos[chave] = {
            "produto_id": row["produto_id"],
            "produto_codigo": row["produto_codigo"],
            "produto_nome": row["produto_nome"],
            "categoria_nome": row["categoria_nome"] or "Sem categoria",
            "unidade": row["unidade"],
            "tipos": row["tipos"] or "-",
            "previsto": float(row["quantidade_prevista"] or 0),
            "recebido": 0.0,
            "utilizado": 0.0,
            "perda": 0.0,
            "devolvido": 0.0,
        }

    for row in consumos_realizados:
        chave = f"{row['produto_id']}:{row['unidade']}"
        item = mapa_consumos.setdefault(chave, {
            "produto_id": row["produto_id"],
            "produto_codigo": row["produto_codigo"],
            "produto_nome": row["produto_nome"],
            "categoria_nome": row["categoria_nome"] or "Sem categoria",
            "unidade": row["unidade"],
            "tipos": "Manual/real",
            "previsto": 0.0,
            "recebido": 0.0,
            "utilizado": 0.0,
            "perda": 0.0,
            "devolvido": 0.0,
        })
        item["recebido"] += float(row["quantidade_recebida"] or 0)
        item["utilizado"] += float(row["quantidade_utilizada"] or 0)
        item["perda"] += float(row["quantidade_perda"] or 0)
        item["devolvido"] += float(row["quantidade_devolvida"] or 0)

    consumos = []
    for item in mapa_consumos.values():
        item["saida_total"] = item["utilizado"] + item["perda"]
        item["diferenca_utilizada"] = item["utilizado"] - item["previsto"]
        item["diferenca_operacional"] = item["saida_total"] - item["previsto"]
        item["percentual"] = _percentual_variacao(item["previsto"], item["utilizado"])
        item["status_variacao"] = _classificar_variacao(item["previsto"], item["utilizado"])
        consumos.append(item)
    consumos.sort(key=lambda x: (x["categoria_nome"].lower(), x["produto_nome"].lower()))

    cursor.execute("""
        SELECT
            SUM(CASE WHEN status = 'Confirmado' THEN 1 ELSE 0 END) AS confirmados,
            SUM(CASE WHEN status = 'Rascunho' THEN 1 ELSE 0 END) AS rascunhos,
            SUM(CASE WHEN status = 'Cancelado' THEN 1 ELSE 0 END) AS cancelados
        FROM consumos_reais_producao
        WHERE producao_realizada_id = ?
    """, (producao_realizada_id,))
    contagens = cursor.fetchone()
    conn.close()

    confirmados = int(contagens["confirmados"] or 0)
    rascunhos = int(contagens["rascunhos"] or 0)

    if producao["status"] == "Cancelado":
        situacao = "Produção cancelada"
    elif producao["status"] != "Confirmado":
        situacao = "Produção ainda não confirmada"
    elif confirmados == 0 and rascunhos == 0:
        situacao = "Aguardando consumo real"
    elif rascunhos > 0:
        situacao = "Consumo real parcial"
    else:
        situacao = "Comparativo disponível"

    indicadores = {
        "produtos_conformes": sum(1 for x in produtos if x["status_variacao"] == "Conforme"),
        "produtos_divergentes": sum(1 for x in produtos if x["status_variacao"] not in ["Conforme", "Sem movimento"]),
        "consumos_conformes": sum(1 for x in consumos if x["status_variacao"] == "Conforme"),
        "consumos_divergentes": sum(1 for x in consumos if x["status_variacao"] not in ["Conforme", "Sem movimento"]),
        "itens_com_perda": sum(1 for x in consumos if x["perda"] > 0.0000001),
    }

    return {
        "producao": producao,
        "planejamento": planejamento,
        "produtos": produtos,
        "consumos": consumos,
        "resumo_produtos": _resumo_por_unidade_analise(produtos, ["previsto", "realizado", "diferenca"]),
        "resumo_consumos": _resumo_por_unidade_analise(
            consumos,
            ["previsto", "recebido", "utilizado", "perda", "devolvido", "diferenca_utilizada"]
        ),
        "consumos_confirmados": confirmados,
        "consumos_rascunho": rascunhos,
        "situacao": situacao,
        "indicadores": indicadores,
    }


# ============================================================
# LOTES E VALIDADES - V2.6
# ============================================================

def _data_para_iso(valor, padrao_hoje=False):
    valor = str(valor or "").strip()
    if not valor:
        return datetime.now().strftime("%Y-%m-%d") if padrao_hoje else ""

    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(valor, formato).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"Data inválida: {valor}.")


def _data_iso_para_br(valor):
    valor = str(valor or "").strip()
    if not valor:
        return "-"
    try:
        return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return valor


def _status_lote_validade(data_validade, quantidade_atual, status_banco="Ativo"):
    status_banco = str(status_banco or "Ativo")
    quantidade_atual = float(quantidade_atual or 0)

    if status_banco == "Cancelado":
        return "Cancelado", None
    if quantidade_atual <= 0:
        return "Encerrado", None

    try:
        validade = datetime.strptime(str(data_validade), "%Y-%m-%d").date()
    except ValueError:
        return "Ativo", None

    hoje = datetime.now().date()
    dias = (validade - hoje).days

    if dias < 0:
        return "Vencido", dias
    if dias == 0:
        return "Vence hoje", dias
    if dias <= 2:
        return "Crítico", dias
    if dias <= 5:
        return "Próximo", dias
    return "Ativo", dias


def _classe_status_validade(status):
    return {
        "Ativo": "status-comprado",
        "Próximo": "status-validade-proximo",
        "Crítico": "status-validade-critico",
        "Vence hoje": "status-validade-hoje",
        "Vencido": "status-cancelado",
        "Encerrado": "status-validade-encerrado",
        "Cancelado": "status-cancelado",
    }.get(status, "status-pendente")


def _proximo_codigo_lote(cursor, produto_id, data_producao):
    cursor.execute("SELECT codigo FROM produtos_estoque WHERE id = ?", (produto_id,))
    produto = cursor.fetchone()
    codigo_produto = (produto["codigo"] if produto else str(produto_id).zfill(4)) or str(produto_id).zfill(4)
    data_codigo = datetime.strptime(data_producao, "%Y-%m-%d").strftime("%Y%m%d")
    prefixo = f"{codigo_produto}-{data_codigo}-"
    cursor.execute("""
        SELECT codigo_lote
        FROM lotes_validade
        WHERE codigo_lote LIKE ?
        ORDER BY codigo_lote DESC
        LIMIT 1
    """, (f"{prefixo}%",))
    ultimo = cursor.fetchone()
    sequencia = 1
    if ultimo:
        try:
            sequencia = int(str(ultimo["codigo_lote"]).split("-")[-1]) + 1
        except (ValueError, IndexError):
            sequencia = 1
    return f"{prefixo}{sequencia:03d}"


def _normalizar_lote(row):
    if not row:
        return None
    lote = dict(row)
    status, dias = _status_lote_validade(
        lote.get("data_validade"),
        lote.get("quantidade_atual"),
        lote.get("status")
    )
    lote["status_calculado"] = status
    lote["classe_status"] = _classe_status_validade(status)
    lote["dias_para_vencer"] = dias
    lote["data_producao_br"] = _data_iso_para_br(lote.get("data_producao"))
    lote["data_validade_br"] = _data_iso_para_br(lote.get("data_validade"))
    return lote


def contar_lotes_validade_resumo():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT lotes_validade.*, produtos_estoque.nome AS produto_nome
        FROM lotes_validade
        INNER JOIN produtos_estoque ON produtos_estoque.id = lotes_validade.produto_id
        WHERE lotes_validade.status != 'Cancelado'
    """)
    lotes = [_normalizar_lote(row) for row in cursor.fetchall()]
    conn.close()
    resumo = {status: 0 for status in STATUS_VALIDADE_FILTROS}
    for lote in lotes:
        resumo[lote["status_calculado"]] = resumo.get(lote["status_calculado"], 0) + 1
    resumo["total"] = len(lotes)
    resumo["alertas"] = sum(resumo.get(x, 0) for x in ["Próximo", "Crítico", "Vence hoje", "Vencido"])
    return resumo


def carregar_lotes_validade(busca="", status="", produto_id="", data_inicio="", data_fim=""):
    conn = conectar()
    cursor = conn.cursor()
    filtros = ["1=1"]
    parametros = []

    busca = str(busca or "").strip()
    status = str(status or "").strip()
    produto_id = str(produto_id or "").strip()
    data_inicio = str(data_inicio or "").strip()
    data_fim = str(data_fim or "").strip()

    if busca:
        filtros.append("(lotes_validade.codigo_lote LIKE ? OR produtos_estoque.nome LIKE ? OR lotes_validade.local_descricao LIKE ?)")
        termo = f"%{busca}%"
        parametros.extend([termo, termo, termo])
    if produto_id:
        filtros.append("lotes_validade.produto_id = ?")
        parametros.append(int(produto_id))
    if data_inicio:
        filtros.append("lotes_validade.data_validade >= ?")
        parametros.append(_data_para_iso(data_inicio))
    if data_fim:
        filtros.append("lotes_validade.data_validade <= ?")
        parametros.append(_data_para_iso(data_fim))

    cursor.execute(f"""
        SELECT
            lotes_validade.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao,
            celulas_producao.nome AS celula_nome
        FROM lotes_validade
        INNER JOIN produtos_estoque ON produtos_estoque.id = lotes_validade.produto_id
        LEFT JOIN celulas_producao ON celulas_producao.id = lotes_validade.celula_id
        WHERE {' AND '.join(filtros)}
        ORDER BY lotes_validade.data_validade, lotes_validade.id DESC
    """, parametros)

    lotes = [_normalizar_lote(row) for row in cursor.fetchall()]
    conn.close()
    if status:
        lotes = [lote for lote in lotes if lote["status_calculado"] == status]
    return lotes


def buscar_lote_validade(lote_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            lotes_validade.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao,
            celulas_producao.nome AS celula_nome
        FROM lotes_validade
        INNER JOIN produtos_estoque ON produtos_estoque.id = lotes_validade.produto_id
        LEFT JOIN celulas_producao ON celulas_producao.id = lotes_validade.celula_id
        WHERE lotes_validade.id = ?
    """, (lote_id,))
    lote = _normalizar_lote(cursor.fetchone())
    conn.close()
    return lote


def carregar_movimentacoes_validade(lote_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT *
        FROM movimentacoes_validade
        WHERE lote_id = ?
        ORDER BY id DESC
    """, (lote_id,))
    linhas = cursor.fetchall()
    conn.close()
    return linhas


def carregar_lotes_por_producao(producao_realizada_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            lotes_validade.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome
        FROM lotes_validade
        INNER JOIN produtos_estoque ON produtos_estoque.id = lotes_validade.produto_id
        WHERE lotes_validade.producao_realizada_id = ?
        ORDER BY lotes_validade.id
    """, (producao_realizada_id,))
    lotes = [_normalizar_lote(row) for row in cursor.fetchall()]
    conn.close()
    return lotes


def criar_lote_validade_cursor(
    cursor,
    produto_id,
    data_producao,
    data_validade,
    quantidade,
    local_tipo="Estoque Central",
    celula_id=None,
    local_descricao="",
    origem="Cadastro Manual",
    producao_realizada_id=None,
    item_producao_realizada_id=None,
    observacao="",
    gerar_entrada_estoque=False
):
    produto = buscar_produto_estoque_cursor(cursor, produto_id)
    if not produto:
        raise ValueError("Produto nao encontrado para o lote.")

    try:
        quantidade = float(str(quantidade).replace(",", "."))
    except ValueError:
        quantidade = 0
    if quantidade <= 0:
        raise ValueError("Informe uma quantidade maior que zero.")

    data_producao = _data_para_iso(data_producao, padrao_hoje=True)
    data_validade = _data_para_iso(data_validade)
    if not data_validade:
        dias = int(produto["dias_validade"] or 0)
        if dias <= 0:
            raise ValueError("Informe a data de validade ou configure os dias de validade do produto.")
        data_validade = (datetime.strptime(data_producao, "%Y-%m-%d") + timedelta(days=dias)).strftime("%Y-%m-%d")
    if data_validade < data_producao:
        raise ValueError("A validade nao pode ser anterior a data de producao.")

    if item_producao_realizada_id:
        cursor.execute(
            "SELECT id FROM lotes_validade WHERE item_producao_realizada_id = ?",
            (item_producao_realizada_id,)
        )
        existente = cursor.fetchone()
        if existente:
            return existente["id"]

    codigo_lote = _proximo_codigo_lote(cursor, produto_id, data_producao)
    cursor.execute("""
        INSERT INTO lotes_validade
        (codigo_lote, produto_id, producao_realizada_id, item_producao_realizada_id,
         data_producao, data_validade, quantidade_inicial, quantidade_atual, unidade,
         local_tipo, celula_id, local_descricao, status, origem, criado_por,
         data_criacao, atualizado_em, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Ativo', ?, ?, ?, ?, ?)
    """, (
        codigo_lote,
        produto_id,
        producao_realizada_id,
        item_producao_realizada_id,
        data_producao,
        data_validade,
        quantidade,
        quantidade,
        produto["unidade_padrao"],
        local_tipo or "Estoque Central",
        int(celula_id) if str(celula_id or "").isdigit() else None,
        str(local_descricao or "").strip(),
        origem,
        _usuario_atual(),
        _agora_texto(),
        _agora_texto(),
        str(observacao or "").strip()
    ))
    lote_id = cursor.lastrowid

    if gerar_entrada_estoque:
        cursor.execute("""
            INSERT INTO movimentacoes_estoque
            (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_id,
             embalagem_nome, fator_conversao, data_movimentacao, origem, origem_id,
             item_compra_id, fornecedor_id, usuario, observacao)
            VALUES (?, 'Entrada', ?, ?, NULL, ?, 1, ?, 'Validade - Ajuste', ?, NULL, NULL, ?, ?)
        """, (
            produto_id,
            quantidade,
            quantidade,
            produto["unidade_padrao"],
            _data_iso_para_br(data_producao),
            lote_id,
            _usuario_atual(),
            f"Entrada criada junto com lote {codigo_lote}. {observacao}".strip()
        ))
        estoque_movimentacao_id = cursor.lastrowid
        cursor.execute("""
            INSERT INTO movimentacoes_validade
            (lote_id, tipo, quantidade, data_movimentacao, usuario, estoque_movimentacao_id, observacao)
            VALUES (?, 'Ajuste positivo', ?, ?, ?, ?, ?)
        """, (
            lote_id,
            quantidade,
            _agora_texto(),
            _usuario_atual(),
            estoque_movimentacao_id,
            "Entrada inicial do lote criada no cadastro manual."
        ))

    return lote_id


def criar_lote_validade(
    produto_id,
    data_producao,
    data_validade,
    quantidade,
    local_tipo="Estoque Central",
    celula_id=None,
    local_descricao="",
    origem="Cadastro Manual",
    producao_realizada_id=None,
    item_producao_realizada_id=None,
    observacao="",
    gerar_entrada_estoque=False
):
    produto = buscar_produto_estoque(produto_id)
    if not produto:
        raise ValueError("Produto não encontrado para o lote.")

    try:
        quantidade = float(str(quantidade).replace(",", "."))
    except ValueError:
        quantidade = 0
    if quantidade <= 0:
        raise ValueError("Informe uma quantidade maior que zero.")

    data_producao = _data_para_iso(data_producao, padrao_hoje=True)
    data_validade = _data_para_iso(data_validade)
    if not data_validade:
        dias = int(produto["dias_validade"] or 0)
        if dias <= 0:
            raise ValueError("Informe a data de validade ou configure os dias de validade do produto.")
        data_validade = (datetime.strptime(data_producao, "%Y-%m-%d") + timedelta(days=dias)).strftime("%Y-%m-%d")
    if data_validade < data_producao:
        raise ValueError("A validade não pode ser anterior à data de produção.")

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")

        if item_producao_realizada_id:
            cursor.execute(
                "SELECT id FROM lotes_validade WHERE item_producao_realizada_id = ?",
                (item_producao_realizada_id,)
            )
            existente = cursor.fetchone()
            if existente:
                conn.rollback()
                return existente["id"]

        codigo_lote = _proximo_codigo_lote(cursor, produto_id, data_producao)
        cursor.execute("""
            INSERT INTO lotes_validade
            (codigo_lote, produto_id, producao_realizada_id, item_producao_realizada_id,
             data_producao, data_validade, quantidade_inicial, quantidade_atual, unidade,
             local_tipo, celula_id, local_descricao, status, origem, criado_por,
             data_criacao, atualizado_em, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Ativo', ?, ?, ?, ?, ?)
        """, (
            codigo_lote,
            produto_id,
            producao_realizada_id,
            item_producao_realizada_id,
            data_producao,
            data_validade,
            quantidade,
            quantidade,
            produto["unidade_padrao"],
            local_tipo or "Estoque Central",
            int(celula_id) if str(celula_id or "").isdigit() else None,
            str(local_descricao or "").strip(),
            origem,
            _usuario_atual(),
            _agora_texto(),
            _agora_texto(),
            str(observacao or "").strip()
        ))
        lote_id = cursor.lastrowid

        if gerar_entrada_estoque:
            cursor.execute("""
                INSERT INTO movimentacoes_estoque
                (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_id,
                 embalagem_nome, fator_conversao, data_movimentacao, origem, origem_id,
                 item_compra_id, fornecedor_id, usuario, observacao)
                VALUES (?, 'Entrada', ?, ?, NULL, ?, 1, ?, 'Validade - Ajuste', ?, NULL, NULL, ?, ?)
            """, (
                produto_id,
                quantidade,
                quantidade,
                produto["unidade_padrao"],
                _data_iso_para_br(data_producao),
                lote_id,
                _usuario_atual(),
                f"Entrada criada junto com lote {codigo_lote}. {observacao}".strip()
            ))
            estoque_movimentacao_id = cursor.lastrowid
            cursor.execute("""
                INSERT INTO movimentacoes_validade
                (lote_id, tipo, quantidade, data_movimentacao, usuario, estoque_movimentacao_id, observacao)
                VALUES (?, 'Ajuste positivo', ?, ?, ?, ?, ?)
            """, (
                lote_id,
                quantidade,
                _agora_texto(),
                _usuario_atual(),
                estoque_movimentacao_id,
                "Entrada inicial do lote criada no cadastro manual."
            ))

        conn.commit()
        return lote_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def criar_lote_validade_item_producao(item, producao, data_producao=None, cursor=None):
    produto = (
        buscar_produto_estoque_cursor(cursor, item["produto_estoque_id"])
        if cursor is not None else
        buscar_produto_estoque(item["produto_estoque_id"])
    )
    if not produto or int(produto["controla_validade"] or 0) != 1:
        return None
    dias = int(produto["dias_validade"] or 0)
    if dias <= 0:
        raise ValueError(
            f"O produto {produto['nome']} controla validade, mas nao possui dias de validade configurados."
        )
    data_producao = _data_para_iso(data_producao or datetime.now().strftime("%Y-%m-%d"), padrao_hoje=True)
    validade = (datetime.strptime(data_producao, "%Y-%m-%d") + timedelta(days=dias)).strftime("%Y-%m-%d")
    if cursor is not None:
        return criar_lote_validade_cursor(
            cursor,
            produto_id=item["produto_estoque_id"],
            data_producao=data_producao,
            data_validade=validade,
            quantidade=item["quantidade_unidades"],
            local_tipo="Estoque Central",
            origem="Producao Confirmada",
            producao_realizada_id=producao["id"],
            item_producao_realizada_id=item["id"],
            observacao=f"Lote automatico da producao realizada #{producao['id']}. Sabor: {item['sabor']}.",
            gerar_entrada_estoque=False
        )
    return criar_lote_validade(
        produto_id=item["produto_estoque_id"],
        data_producao=data_producao,
        data_validade=validade,
        quantidade=item["quantidade_unidades"],
        local_tipo="Estoque Central",
        origem="Produção Confirmada",
        producao_realizada_id=producao["id"],
        item_producao_realizada_id=item["id"],
        observacao=f"Lote automático da produção realizada #{producao['id']}. Sabor: {item['sabor']}.",
        gerar_entrada_estoque=False
    )


def gerar_lotes_producao_confirmada(producao_realizada_id):
    producao = buscar_producao_realizada(producao_realizada_id)
    if not producao:
        raise ValueError("Produção realizada não encontrada.")
    if producao["status"] != "Confirmado":
        raise ValueError("Os lotes só podem ser gerados após confirmar a produção.")
    data_base = _data_para_iso(producao["data_confirmacao"], padrao_hoje=True)
    gerados = 0
    for item in carregar_itens_producao_realizada(producao_realizada_id):
        if item["status"] != "Confirmado" or not item["produto_estoque_id"]:
            continue
        lote_id = criar_lote_validade_item_producao(item, producao, data_base)
        if lote_id:
            gerados += 1
    return gerados


def atualizar_lote_validade(lote_id, data_producao, data_validade, local_tipo, celula_id, local_descricao, observacao):
    lote = buscar_lote_validade(lote_id)
    if not lote:
        raise ValueError("Lote não encontrado.")
    data_producao = _data_para_iso(data_producao)
    data_validade = _data_para_iso(data_validade)
    if data_validade < data_producao:
        raise ValueError("A validade não pode ser anterior à data de produção.")
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE lotes_validade
        SET data_producao = ?, data_validade = ?, local_tipo = ?, celula_id = ?,
            local_descricao = ?, observacao = ?, atualizado_em = ?
        WHERE id = ?
    """, (
        data_producao,
        data_validade,
        local_tipo or "Estoque Central",
        int(celula_id) if str(celula_id or "").isdigit() else None,
        str(local_descricao or "").strip(),
        str(observacao or "").strip(),
        _agora_texto(),
        lote_id
    ))
    conn.commit()
    conn.close()


def movimentar_lote_validade(lote_id, tipo, quantidade, observacao="", motivo_perda_id=None):
    if tipo not in TIPOS_MOVIMENTACAO_VALIDADE:
        raise ValueError("Tipo de movimentação de validade inválido.")

    lote = buscar_lote_validade(lote_id)
    if not lote:
        raise ValueError("Lote não encontrado.")
    if lote["status_calculado"] in ["Cancelado", "Encerrado"]:
        raise ValueError("Esse lote não pode mais ser movimentado.")

    try:
        quantidade = float(str(quantidade).replace(",", "."))
    except ValueError:
        quantidade = 0
    if quantidade <= 0:
        raise ValueError("Informe uma quantidade maior que zero.")

    observacao = str(observacao or "").strip()
    motivo_perda = None
    if tipo == "Descarte":
        conn_motivo = conectar(); cur_motivo = conn_motivo.cursor()
        motivo_perda = buscar_motivo_perda_cursor(cur_motivo, motivo_perda_id, "Validade")
        conn_motivo.close()
        if not motivo_perda:
            raise ValueError("Selecione um motivo de descarte válido.")
        if int(motivo_perda.get("exige_observacao") or 0) == 1 and not observacao:
            raise ValueError("Este motivo exige uma observação detalhada.")
    elif tipo in ["Ajuste positivo", "Ajuste negativo"] and not observacao:
        raise ValueError("Informe o motivo/observação para o ajuste.")

    aumenta = tipo == "Ajuste positivo"
    saldo_lote = float(lote["quantidade_atual"] or 0)
    if not aumenta and quantidade > saldo_lote:
        raise ValueError("A quantidade informada é maior que o saldo atual do lote.")

    saldo_estoque = obter_saldo_produto(lote["produto_id"])
    if not aumenta and quantidade > saldo_estoque:
        raise ValueError(
            f"O estoque geral possui apenas {saldo_estoque:.2f} {lote['unidade']}. "
            "Revise divergências antes de movimentar o lote."
        )

    if tipo == "Descarte":
        origem_estoque = "Validade - Descarte"
    elif tipo in ["Ajuste positivo", "Ajuste negativo"]:
        origem_estoque = "Validade - Ajuste"
    else:
        origem_estoque = "Validade - Baixa"

    novo_saldo = saldo_lote + quantidade if aumenta else saldo_lote - quantidade
    novo_status = "Encerrado" if novo_saldo <= 0.0000001 else "Ativo"
    tipo_estoque = "Entrada" if aumenta else "Saída"
    agora = _agora_texto()

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN")
        cursor.execute("""
            INSERT INTO movimentacoes_estoque
            (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_id,
             embalagem_nome, fator_conversao, data_movimentacao, origem, origem_id,
             item_compra_id, fornecedor_id, usuario, observacao)
            VALUES (?, ?, ?, ?, NULL, ?, 1, ?, ?, ?, NULL, NULL, ?, ?)
        """, (
            lote["produto_id"],
            tipo_estoque,
            quantidade,
            quantidade,
            lote["unidade"],
            datetime.now().strftime("%d/%m/%Y"),
            origem_estoque,
            lote_id,
            _usuario_atual(),
            f"{tipo} do lote {lote['codigo_lote']}. {observacao}".strip()
        ))
        movimento_estoque_id = cursor.lastrowid

        cursor.execute("""
            UPDATE lotes_validade
            SET quantidade_atual = ?, status = ?, atualizado_em = ?,
                encerrado_em = CASE WHEN ? = 'Encerrado' THEN ? ELSE NULL END
            WHERE id = ?
        """, (novo_saldo, novo_status, agora, novo_status, agora, lote_id))

        cursor.execute("""
            INSERT INTO movimentacoes_validade
            (lote_id, tipo, quantidade, data_movimentacao, usuario, estoque_movimentacao_id, observacao, motivo_perda_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            lote_id,
            tipo,
            quantidade,
            agora,
            _usuario_atual(),
            movimento_estoque_id,
            observacao,
            int(motivo_perda_id) if tipo == "Descarte" and str(motivo_perda_id or "").isdigit() else None
        ))
        mov_validade_id = cursor.lastrowid
        if tipo == "Descarte":
            _registrar_perda_cursor(
                cursor, "Validade", mov_validade_id, lote["produto_id"], quantidade, lote["unidade"],
                motivo_id=int(motivo_perda_id), observacao=observacao, lote_id=lote_id,
                data_registro=agora, usuario=_usuario_atual()
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def cancelar_lote_validade(lote_id, observacao=""):
    lote = buscar_lote_validade(lote_id)
    if not lote:
        raise ValueError("Lote não encontrado.")
    if float(lote["quantidade_atual"] or 0) > 0:
        raise ValueError("Dê baixa, descarte ou ajuste o saldo do lote antes de cancelar.")
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE lotes_validade
        SET status = 'Cancelado', atualizado_em = ?, encerrado_em = ?,
            observacao = CASE WHEN observacao IS NULL OR observacao = '' THEN ? ELSE observacao || ' | ' || ? END
        WHERE id = ?
    """, (_agora_texto(), _agora_texto(), observacao or "Lote cancelado.", observacao or "Lote cancelado.", lote_id))
    conn.commit()
    conn.close()


# ============================================================
# ROTAS PRINCIPAIS
# ============================================================

@app.route("/", methods=["GET"])
@producao_obrigatorio
def index():
    return redirect(url_for("producao_home"))


@app.route("/producao", methods=["GET"])
@producao_obrigatorio
def producao_home():
    return render_template(
        "producao_home.html",
        total_pre_producoes=contar_pre_producoes_pendentes(),
        total_fichas_tecnicas=contar_fichas_tecnicas_ativas(),
        total_planejamentos=contar_planejamentos_abertos(),
        total_consumos_reais=contar_consumos_reais_rascunho(),
        total_analises=contar_analises_disponiveis(),
        resumo_validade=contar_lotes_validade_resumo(),
        total_rodadas_abertas=contar_rodadas_abertas()
    )


@app.route("/producao/calcular", methods=["GET", "POST"])
@producao_obrigatorio
def producao_calcular():
    sabores = carregar_sabores()
    resultado = []
    rodadas = carregar_rodadas_pedidos()
    rodada_id = request.values.get("rodada_id", "").strip()
    rodada = buscar_rodada_pedidos(int(rodada_id)) if rodada_id.isdigit() else None
    if not rodada:
        candidatas = [r for r in rodadas if r["status"] in ["Fechada", "Planejada"]]
        rodada = candidatas[0] if candidatas else None
        rodada_id = str(rodada["id"]) if rodada else ""

    cenario = rodada["cenario"] if rodada else "Baixa"
    dia_semana = rodada["dia_semana"] if rodada else "Segunda-feira"
    total_producao = 0
    total_empadas = 0
    valores_digitados = {}
    mistos = carregar_mistos_rodada(int(rodada_id)) if rodada_id else []
    if rodada_id:
        demanda_inicial = demanda_sabores_rodada(int(rodada_id))
        for i, sabor in enumerate(sabores):
            info = demanda_inicial.get(sabor["nome"], {})
            normal_un = float(info.get("normal_unidades", 0) or 0)
            valores_digitados[i] = {
                "inicio": 0,
                "pedido": normal_un / EMPADAS_POR_TABULEIRO,
                "pedido_unidades": normal_un,
                "misto_unidades": float(info.get("misto_unidades", 0) or 0),
            }

    if request.method == "POST":
        resultado, cenario, dia_semana, total_producao, total_empadas, valores_digitados = calcular_producao(request.form)

    return render_template(
        "producao_calculo.html",
        sabores=sabores,
        resultado=resultado,
        cenario=cenario,
        dia_semana=dia_semana,
        valores_digitados=valores_digitados,
        total_producao=total_producao,
        total_empadas=total_empadas,
        rodadas=rodadas,
        rodada=rodada,
        rodada_id=rodada_id,
        mistos=mistos,
    )


@app.route("/salvar-producao", methods=["POST"])
@producao_obrigatorio
def salvar_producao():
    (
        resultado,
        cenario,
        dia_semana,
        total_producao,
        total_empadas,
        valores_digitados
    ) = calcular_producao(request.form)

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    confirmar_duplicado = request.form.get("confirmar_duplicado") == "sim"

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*) AS total
        FROM producoes
        WHERE data_hora LIKE ?
        AND cenario = ?
        AND dia_semana = ?
    """, (
        f"{data_hoje}%",
        cenario,
        dia_semana
    ))

    producoes_iguais_hoje = cursor.fetchone()["total"]

    if producoes_iguais_hoje > 0 and not confirmar_duplicado:
        conn.close()

        return render_template(
            "producao_calculo.html",
            sabores=carregar_sabores(),
            resultado=resultado,
            cenario=cenario,
            dia_semana=dia_semana,
            valores_digitados=valores_digitados,
            total_producao=total_producao,
            total_empadas=total_empadas,
            mensagem_aviso="Já existe uma produção salva hoje para esse cenário e dia. Marque a confirmação abaixo se deseja salvar novamente.",
            mostrar_confirmar_duplicado=True,
            rodadas=carregar_rodadas_pedidos(),
            rodada=buscar_rodada_pedidos(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else None,
            rodada_id=request.form.get("rodada_id", ""),
            mistos=carregar_mistos_rodada(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else [],
        )

    cursor.execute("""
        INSERT INTO producoes
        (data_hora, cenario, dia_semana, total_tabuleiros, total_empadas)
        VALUES (?, ?, ?, ?, ?)
    """, (data_hora, cenario, dia_semana, total_producao, total_empadas))

    producao_id = cursor.lastrowid

    for item in resultado:
        cursor.execute("""
            INSERT INTO itens_producao
            (
                producao_id,
                sabor,
                classe,
                inicio,
                pedido,
                meta,
                origem_meta,
                producao,
                estoque_final,
                empadas
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            producao_id,
            item["nome"],
            item["classe"],
            item["inicio"],
            item["pedido"],
            item["meta"],
            item["origem_meta"],
            item["producao"],
            item["estoque_final"],
            item["empadas"]
        ))

    conn.commit()
    conn.close()

    return render_template(
        "producao_calculo.html",
        sabores=carregar_sabores(),
        resultado=resultado,
        cenario=cenario,
        dia_semana=dia_semana,
        valores_digitados=valores_digitados,
        total_producao=total_producao,
        total_empadas=total_empadas,
        mensagem_sucesso="Produção salva no histórico com sucesso!",
        rodadas=carregar_rodadas_pedidos(),
        rodada=buscar_rodada_pedidos(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else None,
        rodada_id=request.form.get("rodada_id", ""),
        mistos=carregar_mistos_rodada(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else [],
    )



# ============================================================
# ROTAS DE PLANEJAMENTO DE PRODUÇÃO
# ============================================================

@app.route("/planejamento-producao/gerar", methods=["POST"])
@producao_obrigatorio
def planejamento_producao_gerar():
    (
        resultado,
        cenario,
        dia_semana,
        total_producao,
        total_empadas,
        valores_digitados
    ) = calcular_producao(request.form)

    try:
        rodada_id = request.form.get("rodada_id", "").strip()
        rodada = buscar_rodada_pedidos(int(rodada_id)) if rodada_id.isdigit() else None
        if rodada and rodada["planejamento_id"]:
            raise ValueError(f"Esta rodada já possui o planejamento #{rodada['planejamento_id']}.")
        planejamento_id, pre_producao_id = criar_planejamento_producao(
            resultado,
            cenario,
            dia_semana,
            total_producao,
            total_empadas,
            rodada_id=int(rodada_id) if rodada_id.isdigit() else None,
        )
        mensagem_sucesso = (
            f"Planejamento #{planejamento_id} criado com consumo previsto, lista de separação "
            f"e pré-produção #{pre_producao_id}. Nenhum saldo foi movimentado."
        )
        mensagem_erro = None
    except Exception as e:
        planejamento_id = None
        pre_producao_id = None
        mensagem_sucesso = None
        mensagem_erro = f"Não foi possível gerar o planejamento: {str(e)}"

    return render_template(
        "producao_calculo.html",
        sabores=carregar_sabores(),
        resultado=resultado,
        cenario=cenario,
        dia_semana=dia_semana,
        valores_digitados=valores_digitados,
        total_producao=total_producao,
        total_empadas=total_empadas,
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro,
        planejamento_id=planejamento_id,
        pre_producao_id=pre_producao_id,
        rodadas=carregar_rodadas_pedidos(),
        rodada=buscar_rodada_pedidos(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else None,
        rodada_id=request.form.get("rodada_id", ""),
        mistos=carregar_mistos_rodada(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else [],
    )


@app.route("/planejamentos-producao", methods=["GET"])
@producao_obrigatorio
def planejamentos_producao_lista():
    status = request.args.get("status", "").strip()
    return render_template(
        "planejamentos_producao.html",
        planejamentos=carregar_planejamentos_producao(status),
        status=status
    )


@app.route("/planejamentos-producao/<int:planejamento_id>", methods=["GET"])
@producao_obrigatorio
def planejamento_producao_detalhes(planejamento_id):
    planejamento = buscar_planejamento_producao(planejamento_id)
    if not planejamento:
        return render_template("acesso_negado.html", mensagem="Planejamento de produção não encontrado."), 404

    return render_template(
        "planejamento_producao_detalhes.html",
        planejamento=planejamento,
        separacao=buscar_separacao_por_planejamento(planejamento_id),
        itens=carregar_itens_planejamento_producao(planejamento_id),
        intermediarios=carregar_consumos_consolidados_planejamento(planejamento_id, "Intermediário"),
        insumos=carregar_consumos_consolidados_planejamento(planejamento_id, "Insumo"),
        consumos_detalhados=carregar_consumos_detalhados_planejamento(planejamento_id)
    )


@app.route("/planejamentos-producao/<int:planejamento_id>/cancelar", methods=["POST"])
@producao_obrigatorio
def planejamento_producao_cancelar(planejamento_id):
    try:
        cancelar_planejamento_producao(planejamento_id)
        return redirect(url_for("planejamento_producao_detalhes", planejamento_id=planejamento_id, mensagem="Planejamento cancelado."))
    except Exception as e:
        return redirect(url_for("planejamento_producao_detalhes", planejamento_id=planejamento_id, erro=str(e)))



# ============================================================
# ROTAS DE SEPARAÇÃO DO ESTOQUE
# ============================================================

@app.route("/estoque/separacoes", methods=["GET"])
@estoque_obrigatorio
def estoque_separacoes():
    status = request.args.get("status", "").strip()
    busca = request.args.get("busca", "").strip()
    return render_template(
        "estoque_separacoes.html",
        separacoes=carregar_separacoes_estoque(status, busca),
        status=status,
        busca=busca,
        total_pendentes=contar_separacoes_pendentes()
    )


@app.route("/estoque/separacoes/<int:separacao_id>", methods=["GET"])
@estoque_obrigatorio
def estoque_separacao_detalhes(separacao_id):
    separacao = buscar_separacao_estoque(separacao_id)
    if not separacao:
        return render_template("acesso_negado.html", mensagem="Separação de estoque não encontrada."), 404

    return render_template(
        "estoque_separacao_detalhes.html",
        separacao=separacao,
        itens=carregar_itens_separacao_estoque(separacao_id)
    )


@app.route("/estoque/separacoes/<int:separacao_id>/salvar", methods=["POST"])
@estoque_obrigatorio
def estoque_separacao_salvar(separacao_id):
    try:
        novo_status = atualizar_separacao_estoque(separacao_id, request.form)
        mensagem = f"Separação atualizada. Status atual: {novo_status}. Nenhum saldo foi movimentado."
        return redirect(url_for("estoque_separacao_detalhes", separacao_id=separacao_id, mensagem=mensagem))
    except Exception as e:
        return redirect(url_for("estoque_separacao_detalhes", separacao_id=separacao_id, erro=str(e)))


# ============================================================
# ROTAS DE ETIQUETAS
# ============================================================

@app.route("/etiquetas", methods=["POST"])
@producao_obrigatorio
def etiquetas():
    (
        resultado,
        cenario,
        dia_semana,
        total_producao,
        total_empadas,
        valores_digitados
    ) = calcular_producao(request.form)

    try:
        rodada_id = request.form.get("rodada_id", "").strip()
        mistos = carregar_mistos_rodada(int(rodada_id)) if rodada_id.isdigit() else []
        etiquetas_geradas = montar_etiquetas_producao(resultado, mistos)
        imprimir_etiquetas_direto(etiquetas_geradas)
        mensagem_sucesso = (
            f"{len(etiquetas_geradas)} etiquetas enviadas para a Argox com sucesso. "
            "A impressão não cria outra pré-produção; use Gerar Planejamento para registrar previsão e conferência."
        )
        mensagem_erro = None

    except Exception as e:
        mensagem_sucesso = None
        mensagem_erro = f"Erro ao imprimir direto na Argox: {str(e)}"

    return render_template(
        "producao_calculo.html",
        sabores=carregar_sabores(),
        resultado=resultado,
        cenario=cenario,
        dia_semana=dia_semana,
        valores_digitados=valores_digitados,
        total_producao=total_producao,
        total_empadas=total_empadas,
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro,
        rodadas=carregar_rodadas_pedidos(),
        rodada=buscar_rodada_pedidos(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else None,
        rodada_id=request.form.get("rodada_id", ""),
        mistos=carregar_mistos_rodada(int(request.form.get("rodada_id"))) if str(request.form.get("rodada_id", "")).isdigit() else [],
    )


@app.route("/etiquetas-avulsas", methods=["GET"])
@producao_obrigatorio
def etiquetas_avulsas():
    return render_template(
        "etiquetas_avulsas.html",
        sabores=carregar_sabores()
    )


@app.route("/gerar-etiquetas-avulsas", methods=["POST"])
@producao_obrigatorio
def gerar_etiquetas_avulsas():
    try:
        etiquetas_geradas = montar_etiquetas_avulsas(request.form)
        imprimir_etiquetas_direto(etiquetas_geradas)

        mensagem_sucesso = f"{len(etiquetas_geradas)} etiquetas avulsas enviadas para a Argox com sucesso."
        mensagem_erro = None

    except Exception as e:
        mensagem_sucesso = None
        mensagem_erro = f"Erro ao imprimir etiquetas avulsas: {str(e)}"

    return render_template(
        "etiquetas_avulsas.html",
        sabores=carregar_sabores(),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )



# ============================================================
# ROTAS DE PRODUÇÃO REALIZADA / PRÉ-PRODUÇÃO
# ============================================================

@app.route("/producao-realizada", methods=["GET"])
@producao_obrigatorio
def producao_realizada_lista():
    status = request.args.get("status", "").strip()

    return render_template(
        "producao_realizada.html",
        producoes=carregar_producoes_realizadas(status),
        status=status,
        total_pendentes=contar_pre_producoes_pendentes()
    )


@app.route("/producao-realizada/<int:producao_realizada_id>", methods=["GET"])
@producao_obrigatorio
def producao_realizada_detalhes(producao_realizada_id):
    producao = buscar_producao_realizada(producao_realizada_id)

    if not producao:
        return render_template(
            "acesso_negado.html",
            mensagem="Pré-produção não encontrada."
        )

    return render_template(
        "producao_realizada_detalhes.html",
        producao=producao,
        itens=carregar_itens_producao_realizada(producao_realizada_id),
        produtos=carregar_produtos_estoque(apenas_ativos=True),
        embalagens_por_produto=carregar_embalagens_todos_produtos(),
        consumos_reais=carregar_consumos_reais_producao(producao_realizada_id),
        lotes_validade=carregar_lotes_por_producao(producao_realizada_id)
    )


@app.route("/producao-realizada/confirmar/<int:producao_realizada_id>", methods=["POST"])
@producao_obrigatorio
def producao_realizada_confirmar(producao_realizada_id):
    observacao = request.form.get("observacao_confirmacao", "").strip()

    try:
        confirmar_producao_realizada_estoque(producao_realizada_id, observacao)
    except Exception as e:
        return render_template(
            "acesso_negado.html",
            mensagem=f"Não foi possível confirmar a produção: {str(e)}"
        )

    return redirect(url_for("producao_realizada_detalhes", producao_realizada_id=producao_realizada_id))


@app.route("/producao-realizada/item/<int:item_id>/atualizar", methods=["POST"])
@producao_obrigatorio
def producao_realizada_item_atualizar(item_id):
    try:
        producao_realizada_id = atualizar_item_pre_producao(
            item_id=item_id,
            produto_estoque_id=request.form.get("produto_estoque_id"),
            embalagem_id=request.form.get("embalagem_id"),
            quantidade_embalagem=request.form.get("quantidade_embalagem"),
            observacao=request.form.get("observacao", "")
        )
    except Exception as e:
        return render_template(
            "acesso_negado.html",
            mensagem=f"Não foi possível atualizar o item da produção: {str(e)}"
        )

    return redirect(url_for("producao_realizada_detalhes", producao_realizada_id=producao_realizada_id))


@app.route("/producao-realizada/item/<int:item_id>/cancelar", methods=["POST"])
@producao_obrigatorio
def producao_realizada_item_cancelar(item_id):
    try:
        producao_realizada_id = cancelar_item_pre_producao(item_id)
    except Exception as e:
        return render_template(
            "acesso_negado.html",
            mensagem=f"Não foi possível cancelar o item da produção: {str(e)}"
        )

    return redirect(url_for("producao_realizada_detalhes", producao_realizada_id=producao_realizada_id))


@app.route("/producao-realizada/<int:producao_realizada_id>/adicionar-item", methods=["POST"])
@producao_obrigatorio
def producao_realizada_item_adicionar(producao_realizada_id):
    try:
        adicionar_item_pre_producao(
            producao_realizada_id=producao_realizada_id,
            produto_estoque_id=request.form.get("produto_estoque_id"),
            embalagem_id=request.form.get("embalagem_id"),
            quantidade_embalagem=request.form.get("quantidade_embalagem"),
            sabor=request.form.get("sabor", ""),
            observacao=request.form.get("observacao", "")
        )
    except Exception as e:
        return render_template(
            "acesso_negado.html",
            mensagem=f"Não foi possível adicionar o item na produção: {str(e)}"
        )

    return redirect(url_for("producao_realizada_detalhes", producao_realizada_id=producao_realizada_id))


@app.route("/producao-realizada/cancelar/<int:producao_realizada_id>", methods=["POST"])
@producao_obrigatorio
def producao_realizada_cancelar(producao_realizada_id):
    producao = buscar_producao_realizada(producao_realizada_id)

    if not producao:
        return redirect(url_for("producao_realizada_lista"))

    if producao["status"] != "Pendente":
        return render_template(
            "acesso_negado.html",
            mensagem="Somente pré-produções pendentes podem ser canceladas."
        )

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE producoes_realizadas
        SET status = 'Cancelado',
            data_cancelamento = ?,
            cancelado_por = ?
        WHERE id = ?
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        session.get("usuario"),
        producao_realizada_id
    ))

    cursor.execute("""
        UPDATE itens_producao_realizada
        SET status = 'Cancelado'
        WHERE producao_realizada_id = ?
    """, (producao_realizada_id,))

    cursor.execute("""
        UPDATE consumos_reais_producao
        SET status = 'Cancelado', data_cancelamento = ?, cancelado_por = ?
        WHERE producao_realizada_id = ? AND status = 'Rascunho'
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        session.get("usuario"),
        producao_realizada_id
    ))

    conn.commit()
    conn.close()

    return redirect(url_for("producao_realizada_detalhes", producao_realizada_id=producao_realizada_id))




# ============================================================
# ROTAS DE CONSUMO REAL DAS CÉLULAS - V2.4
# ============================================================

@app.route("/consumos-producao", methods=["GET"])
@consumo_producao_obrigatorio
def consumos_producao_lista():
    status = request.args.get("status", "").strip()
    celula_id = request.args.get("celula_id", "").strip()
    return render_template(
        "consumos_producao.html",
        consumos=carregar_consumos_reais(status=status, celula_id=celula_id),
        status=status,
        celula_id=celula_id,
        celulas=carregar_celulas_producao(apenas_ativas=False),
        total_rascunhos=contar_consumos_reais_rascunho()
    )


@app.route("/producao-realizada/<int:producao_realizada_id>/consumos/novo", methods=["GET", "POST"])
@consumo_producao_obrigatorio
def consumo_producao_novo(producao_realizada_id):
    producao = buscar_producao_realizada(producao_realizada_id)
    if not producao:
        return render_template("acesso_negado.html", mensagem="Produção realizada não encontrada.")

    if request.method == "POST":
        try:
            consumo_id = criar_consumo_real_producao(
                producao_realizada_id,
                int(request.form.get("celula_id")),
                request.form.get("observacao", "")
            )
            return redirect(url_for("consumo_producao_detalhes", consumo_id=consumo_id))
        except Exception as e:
            return render_template(
                "acesso_negado.html",
                mensagem=f"Não foi possível criar o consumo real: {str(e)}"
            )

    existentes = {int(c["celula_id"]) for c in carregar_consumos_reais_producao(producao_realizada_id)}
    celulas = [c for c in carregar_celulas_producao() if int(c["id"]) not in existentes]
    return render_template(
        "consumo_producao_novo.html",
        producao=producao,
        celulas=celulas
    )


@app.route("/consumos-producao/<int:consumo_id>", methods=["GET"])
@consumo_producao_obrigatorio
def consumo_producao_detalhes(consumo_id):
    consumo = buscar_consumo_real_producao(consumo_id)
    if not consumo:
        return render_template("acesso_negado.html", mensagem="Registro de consumo não encontrado.")

    itens = carregar_itens_consumo_real(consumo_id)
    produtos_celula = []
    for produto in carregar_produtos_estoque(apenas_ativos=True):
        if obter_saldo_celula_produto(consumo["celula_id"], produto["id"]) > 0:
            produtos_celula.append(produto)

    return render_template(
        "consumo_producao_detalhes.html",
        consumo=consumo,
        itens=itens,
        resumo=resumo_consumo_real(itens),
        produtos_celula=produtos_celula,
        motivos_perda=carregar_motivos_perda("Produção")
    )


@app.route("/consumos-producao/<int:consumo_id>/salvar", methods=["POST"])
@consumo_producao_obrigatorio
def consumo_producao_salvar(consumo_id):
    acao = request.form.get("acao", "salvar")
    try:
        salvar_itens_consumo_real(consumo_id, request.form)
        if acao == "confirmar":
            confirmar_consumo_real_producao(consumo_id)
    except Exception as e:
        return render_template(
            "acesso_negado.html",
            mensagem=f"Não foi possível salvar o consumo real: {str(e)}"
        )
    return redirect(url_for("consumo_producao_detalhes", consumo_id=consumo_id))


@app.route("/consumos-producao/<int:consumo_id>/adicionar-item", methods=["POST"])
@consumo_producao_obrigatorio
def consumo_producao_adicionar_item(consumo_id):
    try:
        adicionar_item_consumo_real(
            consumo_id,
            int(request.form.get("produto_id")),
            request.form.get("observacao", "")
        )
    except Exception as e:
        return render_template("acesso_negado.html", mensagem=f"Não foi possível adicionar o item: {str(e)}")
    return redirect(url_for("consumo_producao_detalhes", consumo_id=consumo_id))


@app.route("/consumos-producao/<int:consumo_id>/item/<int:item_id>/remover", methods=["POST"])
@consumo_producao_obrigatorio
def consumo_producao_remover_item(consumo_id, item_id):
    try:
        remover_item_manual_consumo(consumo_id, item_id)
    except Exception as e:
        return render_template("acesso_negado.html", mensagem=f"Não foi possível remover o item: {str(e)}")
    return redirect(url_for("consumo_producao_detalhes", consumo_id=consumo_id))


@app.route("/consumos-producao/<int:consumo_id>/cancelar", methods=["POST"])
@consumo_producao_obrigatorio
def consumo_producao_cancelar(consumo_id):
    try:
        cancelar_consumo_real_producao(consumo_id)
    except Exception as e:
        return render_template("acesso_negado.html", mensagem=f"Não foi possível cancelar o consumo: {str(e)}")
    return redirect(url_for("consumo_producao_detalhes", consumo_id=consumo_id))


# ============================================================
# ROTAS DE FICHA TÉCNICA V2.0
# ============================================================

@app.route("/fichas-tecnicas", methods=["GET"])
@producao_obrigatorio
def fichas_tecnicas_lista():
    busca = request.args.get("busca", "").strip()
    ativo = request.args.get("ativo", "").strip()

    return render_template(
        "fichas_tecnicas.html",
        fichas=carregar_fichas_tecnicas(busca, ativo),
        busca=busca,
        ativo=ativo,
        total_ativas=contar_fichas_tecnicas_ativas(),
        mensagem_sucesso=request.args.get("sucesso", "").strip(),
        mensagem_erro=request.args.get("erro", "").strip()
    )


@app.route("/fichas-tecnicas/nova", methods=["GET", "POST"])
@producao_obrigatorio
def ficha_tecnica_nova():
    mensagem_erro = None

    if request.method == "POST":
        try:
            ficha_id = criar_ficha_tecnica(
                produto_final_id=request.form.get("produto_final_id"),
                tipo=request.form.get("tipo"),
                rendimento_quantidade=request.form.get("rendimento_quantidade"),
                rendimento_unidade=request.form.get("rendimento_unidade"),
                observacao=request.form.get("observacao", "")
            )
            return redirect(url_for(
                "ficha_tecnica_detalhes",
                ficha_id=ficha_id,
                sucesso="Ficha técnica criada com sucesso."
            ))
        except Exception as e:
            mensagem_erro = str(e)

    return render_template(
        "ficha_tecnica_form.html",
        produtos=carregar_produtos_sem_ficha(),
        mensagem_erro=mensagem_erro
    )


@app.route("/fichas-tecnicas/<int:ficha_id>", methods=["GET", "POST"])
@producao_obrigatorio
def ficha_tecnica_detalhes(ficha_id):
    ficha = buscar_ficha_tecnica(ficha_id)

    if not ficha:
        return render_template(
            "acesso_negado.html",
            mensagem="Ficha técnica não encontrada."
        )

    mensagem_sucesso = request.args.get("sucesso", "").strip() or None
    mensagem_erro = request.args.get("erro", "").strip() or None

    if request.method == "POST":
        try:
            atualizar_ficha_tecnica(
                ficha_id=ficha_id,
                tipo=request.form.get("tipo"),
                rendimento_quantidade=request.form.get("rendimento_quantidade"),
                rendimento_unidade=request.form.get("rendimento_unidade"),
                observacao=request.form.get("observacao", "")
            )
            return redirect(url_for(
                "ficha_tecnica_detalhes",
                ficha_id=ficha_id,
                sucesso="Dados da ficha técnica atualizados com sucesso."
            ))
        except Exception as e:
            mensagem_erro = str(e)
            ficha = buscar_ficha_tecnica(ficha_id)

    quantidade_simulada = request.args.get("quantidade_simulada", "").strip()
    itens = carregar_itens_ficha_tecnica(ficha_id)
    simulacao = None

    if quantidade_simulada:
        simulacao = simular_consumo_ficha(ficha, itens, quantidade_simulada)

    return render_template(
        "ficha_tecnica_detalhes.html",
        ficha=ficha,
        itens=itens,
        produtos=carregar_produtos_estoque(apenas_ativos=True),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro,
        quantidade_simulada=quantidade_simulada,
        simulacao=simulacao,
        total_utilizacoes=contar_utilizacoes_ficha(ficha_id)
    )


@app.route("/fichas-tecnicas/<int:ficha_id>/adicionar-item", methods=["POST"])
@producao_obrigatorio
def ficha_tecnica_adicionar_item(ficha_id):
    try:
        adicionar_item_ficha_tecnica(
            ficha_id=ficha_id,
            insumo_produto_id=request.form.get("insumo_produto_id"),
            quantidade=request.form.get("quantidade"),
            unidade=request.form.get("unidade"),
            observacao=request.form.get("observacao", "")
        )
        return redirect(url_for(
            "ficha_tecnica_detalhes",
            ficha_id=ficha_id,
            sucesso="Componente adicionado à ficha técnica."
        ))
    except Exception as e:
        return redirect(url_for(
            "ficha_tecnica_detalhes",
            ficha_id=ficha_id,
            erro=str(e)
        ))


@app.route("/fichas-tecnicas/item/<int:item_id>/editar", methods=["GET", "POST"])
@producao_obrigatorio
def ficha_tecnica_item_editar(item_id):
    item = buscar_item_ficha_tecnica(item_id)

    if not item:
        return render_template("acesso_negado.html", mensagem="Componente da ficha técnica não encontrado.")

    mensagem_erro = None

    if request.method == "POST":
        try:
            ficha_id = atualizar_item_ficha_tecnica(
                item_id=item_id,
                insumo_produto_id=request.form.get("insumo_produto_id"),
                quantidade=request.form.get("quantidade"),
                unidade=request.form.get("unidade"),
                observacao=request.form.get("observacao", "")
            )
            return redirect(url_for(
                "ficha_tecnica_detalhes",
                ficha_id=ficha_id,
                sucesso="Componente atualizado com sucesso."
            ))
        except Exception as e:
            mensagem_erro = str(e)
            item = buscar_item_ficha_tecnica(item_id)

    return render_template(
        "ficha_tecnica_item_editar.html",
        item=item,
        produtos=carregar_produtos_estoque(apenas_ativos=True),
        mensagem_erro=mensagem_erro
    )


@app.route("/fichas-tecnicas/item/<int:item_id>/alternar", methods=["POST"])
@producao_obrigatorio
def ficha_tecnica_item_alternar(item_id):
    item = buscar_item_ficha_tecnica(item_id)
    ficha_id = item["ficha_id"] if item else None

    try:
        ficha_id = alternar_item_ficha_tecnica(item_id)
        return redirect(url_for(
            "ficha_tecnica_detalhes",
            ficha_id=ficha_id,
            sucesso="Status do componente alterado."
        ))
    except Exception as e:
        if ficha_id:
            return redirect(url_for("ficha_tecnica_detalhes", ficha_id=ficha_id, erro=str(e)))
        return render_template("acesso_negado.html", mensagem=str(e))


@app.route("/fichas-tecnicas/item/<int:item_id>/remover", methods=["POST"])
@producao_obrigatorio
def ficha_tecnica_item_remover(item_id):
    item = buscar_item_ficha_tecnica(item_id)
    ficha_id = item["ficha_id"] if item else None

    try:
        ficha_id = remover_item_ficha_tecnica(item_id)
        return redirect(url_for(
            "ficha_tecnica_detalhes",
            ficha_id=ficha_id,
            sucesso="Componente removido da ficha técnica."
        ))
    except Exception as e:
        if ficha_id:
            return redirect(url_for("ficha_tecnica_detalhes", ficha_id=ficha_id, erro=str(e)))
        return render_template("acesso_negado.html", mensagem=str(e))


@app.route("/fichas-tecnicas/<int:ficha_id>/alternar", methods=["POST"])
@producao_obrigatorio
def ficha_tecnica_alternar(ficha_id):
    try:
        alternar_ficha_tecnica(ficha_id)
        return redirect(url_for(
            "ficha_tecnica_detalhes",
            ficha_id=ficha_id,
            sucesso="Status da ficha técnica alterado."
        ))
    except Exception as e:
        return redirect(url_for("ficha_tecnica_detalhes", ficha_id=ficha_id, erro=str(e)))


@app.route("/fichas-tecnicas/<int:ficha_id>/duplicar", methods=["GET", "POST"])
@producao_obrigatorio
def ficha_tecnica_duplicar(ficha_id):
    ficha = buscar_ficha_tecnica(ficha_id)

    if not ficha:
        return render_template("acesso_negado.html", mensagem="Ficha técnica de origem não encontrada.")

    mensagem_erro = None

    if request.method == "POST":
        try:
            nova_ficha_id = duplicar_ficha_tecnica(
                ficha_id=ficha_id,
                produto_destino_id=request.form.get("produto_destino_id"),
                observacao_extra=request.form.get("observacao_extra", "")
            )
            return redirect(url_for(
                "ficha_tecnica_detalhes",
                ficha_id=nova_ficha_id,
                sucesso="Ficha técnica duplicada com sucesso. Confira os componentes antes de utilizá-la."
            ))
        except Exception as e:
            mensagem_erro = str(e)

    return render_template(
        "ficha_tecnica_duplicar.html",
        ficha=ficha,
        produtos=carregar_produtos_sem_ficha(excluir_produto_id=ficha["produto_final_id"]),
        mensagem_erro=mensagem_erro
    )


@app.route("/fichas-tecnicas/<int:ficha_id>/excluir", methods=["POST"])
@producao_obrigatorio
def ficha_tecnica_excluir(ficha_id):
    try:
        excluir_ficha_tecnica(ficha_id)
        return redirect(url_for(
            "fichas_tecnicas_lista",
            sucesso="Ficha técnica excluída definitivamente."
        ))
    except Exception as e:
        return redirect(url_for("ficha_tecnica_detalhes", ficha_id=ficha_id, erro=str(e)))


@app.route("/fichas-tecnicas/<int:ficha_id>/revisoes", methods=["GET"])
@producao_obrigatorio
def ficha_tecnica_revisoes(ficha_id):
    ficha = buscar_ficha_tecnica(ficha_id)

    if not ficha:
        return render_template("acesso_negado.html", mensagem="Ficha técnica não encontrada.")

    return render_template(
        "ficha_tecnica_revisoes.html",
        ficha=ficha,
        revisoes=carregar_revisoes_ficha(ficha_id)
    )


@app.route("/fichas-tecnicas/<int:ficha_id>/revisoes/<int:revisao_id>", methods=["GET"])
@producao_obrigatorio
def ficha_tecnica_revisao_detalhes(ficha_id, revisao_id):
    ficha = buscar_ficha_tecnica(ficha_id)
    revisao, dados = buscar_revisao_ficha(ficha_id, revisao_id)

    if not ficha or not revisao:
        return render_template("acesso_negado.html", mensagem="Revisão da ficha técnica não encontrada.")

    return render_template(
        "ficha_tecnica_revisao_detalhes.html",
        ficha=ficha,
        revisao=revisao,
        dados=dados
    )


# ============================================================
# ROTAS DE METAS
# ============================================================

@app.route("/metas", methods=["GET", "POST"])
@admin_obrigatorio
def metas():
    sabores = carregar_sabores()

    if request.method == "POST":
        conn = conectar()
        cursor = conn.cursor()

        for sabor in sabores:
            sabor_id = sabor["id"]

            nome = request.form.get(f"nome_{sabor_id}", "").strip()
            classe = request.form.get(f"classe_{sabor_id}", "Bronze")

            chuva = int(request.form.get(f"chuva_{sabor_id}", 0) or 0)
            normal = int(request.form.get(f"normal_{sabor_id}", 0) or 0)
            verao = int(request.form.get(f"verao_{sabor_id}", 0) or 0)
            baixa = int(request.form.get(f"baixa_{sabor_id}", 0) or 0)

            cursor.execute("""
                UPDATE sabores
                SET nome = ?, classe = ?, chuva = ?, normal = ?, verao = ?, baixa = ?
                WHERE id = ?
            """, (nome, classe, chuva, normal, verao, baixa, sabor_id))

            for dia in DIAS_SEMANA:
                campo = f"dia_{dia}_{sabor_id}"
                valor = request.form.get(campo, "").strip()

                cursor.execute("""
                    DELETE FROM metas_dia
                    WHERE dia_semana = ? AND sabor_id = ?
                """, (dia, sabor_id))

                if valor != "":
                    meta_dia = int(valor)

                    cursor.execute("""
                        INSERT INTO metas_dia
                        (dia_semana, sabor_id, meta)
                        VALUES (?, ?, ?)
                    """, (dia, sabor_id, meta_dia))

        conn.commit()
        conn.close()

        return redirect(url_for("metas"))

    return render_template(
        "metas.html",
        sabores=carregar_sabores(),
        dias_semana=DIAS_SEMANA,
        metas_por_dia=carregar_metas_por_dia(),
        dias_exclusivos=DIAS_EXCLUSIVOS
    )


# ============================================================
# ROTAS DE HISTÓRICO
# ============================================================

@app.route("/historico", methods=["GET"])
@admin_obrigatorio
def historico():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            id,
            data_hora,
            cenario,
            dia_semana,
            total_tabuleiros,
            total_empadas
        FROM producoes
        ORDER BY id DESC
    """)

    producoes = cursor.fetchall()
    conn.close()

    return render_template(
        "historico.html",
        producoes=producoes
    )


@app.route("/historico/<int:producao_id>", methods=["GET"])
@admin_obrigatorio
def historico_detalhes(producao_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            id,
            data_hora,
            cenario,
            dia_semana,
            total_tabuleiros,
            total_empadas
        FROM producoes
        WHERE id = ?
    """, (producao_id,))

    producao = cursor.fetchone()

    if producao is None:
        conn.close()

        return render_template(
            "historico_detalhes.html",
            producao=None,
            itens=[]
        )

    cursor.execute("""
        SELECT
            sabor,
            classe,
            inicio,
            pedido,
            meta,
            origem_meta,
            producao,
            estoque_final,
            empadas
        FROM itens_producao
        WHERE producao_id = ?
        ORDER BY id
    """, (producao_id,))

    itens = cursor.fetchall()
    conn.close()

    return render_template(
        "historico_detalhes.html",
        producao=producao,
        itens=itens
    )


@app.route("/historico/excluir/<int:producao_id>", methods=["POST"])
@admin_obrigatorio
def excluir_producao(producao_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        DELETE FROM itens_producao
        WHERE producao_id = ?
    """, (producao_id,))

    cursor.execute("""
        DELETE FROM producoes
        WHERE id = ?
    """, (producao_id,))

    conn.commit()
    conn.close()

    return redirect(url_for("historico"))


# ============================================================
# ROTAS DE BACKUP
# ============================================================

@app.route("/backup", methods=["GET"])
@admin_obrigatorio
def backup():
    return render_template(
        "backup.html",
        backups=listar_backups()
    )


@app.route("/fazer-backup", methods=["POST"])
@admin_obrigatorio
def fazer_backup():
    nome_backup = criar_backup_banco()

    return render_template(
        "backup.html",
        backups=listar_backups(),
        mensagem_sucesso=f"Backup criado com sucesso: {nome_backup}"
    )


@app.route("/baixar-backup/<nome_arquivo>", methods=["GET"])
@admin_obrigatorio
def baixar_backup(nome_arquivo):
    if ".." in nome_arquivo or "/" in nome_arquivo or "\\" in nome_arquivo:
        return "Nome de arquivo inválido", 400

    caminho = os.path.join("backups", nome_arquivo)

    if not os.path.exists(caminho):
        return "Backup não encontrado", 404

    return send_file(
        caminho,
        as_attachment=True,
        download_name=nome_arquivo
    )


# ============================================================
# ROTAS DE USUÁRIOS
# ============================================================

@app.route("/usuarios", methods=["GET", "POST"])
@admin_obrigatorio
def usuarios():
    mensagem_sucesso = None
    mensagem_erro = None

    conn = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        novo_usuario = request.form.get("usuario", "").strip()
        nova_senha = request.form.get("senha", "").strip()
        perfil = request.form.get("perfil", "producao")
        loja_id = request.form.get("loja_id", "").strip()

        if novo_usuario == "" or nova_senha == "":
            mensagem_erro = "Preencha usuário e senha."
        elif perfil not in PERFIS_USUARIO:
            mensagem_erro = "Perfil inválido."
        elif perfil == "loja" and not loja_id:
            mensagem_erro = "Selecione a loja para o usuário de loja."
        else:
            try:
                cursor.execute("""
                    INSERT INTO usuarios (usuario, senha_hash, perfil, loja_id)
                    VALUES (?, ?, ?, ?)
                """, (
                    novo_usuario,
                    generate_password_hash(nova_senha),
                    perfil,
                    int(loja_id) if perfil == "loja" and loja_id else None
                ))

                conn.commit()
                mensagem_sucesso = "Usuário criado com sucesso."

            except sqlite3.IntegrityError:
                mensagem_erro = "Esse usuário já existe."

    cursor.execute("""
        SELECT usuarios.id, usuarios.usuario, usuarios.perfil, usuarios.loja_id, lojas.nome AS loja_nome
        FROM usuarios
        LEFT JOIN lojas ON lojas.id = usuarios.loja_id
        ORDER BY usuarios.perfil, usuarios.usuario
    """)

    lista_usuarios = cursor.fetchall()
    conn.close()

    return render_template(
        "usuarios.html",
        usuarios=lista_usuarios,
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro,
        lojas=carregar_lojas(apenas_ativas=False)
    )


@app.route("/usuarios/trocar-senha/<int:usuario_id>", methods=["POST"])
@admin_obrigatorio
def trocar_senha_usuario(usuario_id):
    nova_senha = request.form.get("nova_senha", "").strip()

    if nova_senha == "":
        return redirect(url_for("usuarios"))

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE usuarios
        SET senha_hash = ?
        WHERE id = ?
    """, (
        generate_password_hash(nova_senha),
        usuario_id
    ))

    conn.commit()
    conn.close()

    return redirect(url_for("usuarios"))


@app.route("/usuarios/excluir/<int:usuario_id>", methods=["POST"])
@admin_obrigatorio
def excluir_usuario(usuario_id):
    usuario_logado = session.get("usuario")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, usuario, perfil
        FROM usuarios
        WHERE id = ?
    """, (usuario_id,))

    usuario = cursor.fetchone()

    if not usuario:
        conn.close()
        return redirect(url_for("usuarios"))

    # Não deixa o admin apagar a si mesmo
    if usuario["usuario"] == usuario_logado:
        conn.close()
        return redirect(url_for("usuarios"))

    # Não deixa apagar o último admin
    if usuario["perfil"] == "admin":
        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM usuarios
            WHERE perfil = 'admin'
        """)

        total_admins = cursor.fetchone()["total"]

        if total_admins <= 1:
            conn.close()
            return redirect(url_for("usuarios"))

    cursor.execute("""
        DELETE FROM usuarios
        WHERE id = ?
    """, (usuario_id,))

    conn.commit()
    conn.close()

    return redirect(url_for("usuarios"))




# ============================================================
# ROTAS DO MÓDULO DE COMPRAS
# ============================================================

@app.route("/compras", methods=["GET"])
@pode_ver_compras_obrigatorio
def compras_inicio():
    return render_template("compras_home.html", total_pendentes=contar_pedidos_pendentes())


@app.route("/compras/novo", methods=["GET", "POST"])
@pode_solicitar_compra_obrigatorio
def compras_novo():
    mensagem_erro = None

    if request.method == "POST":
        descricoes = request.form.getlist("descricao[]")
        quantidades = request.form.getlist("quantidade[]")
        unidades = request.form.getlist("unidade[]")
        categorias = request.form.getlist("categoria_id[]")
        observacoes_itens = request.form.getlist("observacao_item[]")
        observacao_pedido = request.form.get("observacao_pedido", "").strip()

        itens_validos = []

        for i, descricao in enumerate(descricoes):
            descricao = descricao.strip()

            if descricao == "":
                continue

            try:
                quantidade = float(str(quantidades[i]).replace(",", "."))
            except:
                quantidade = 0

            unidade = unidades[i].strip() if i < len(unidades) else "un"
            categoria_id = categorias[i] if i < len(categorias) else ""
            observacao_item = observacoes_itens[i].strip() if i < len(observacoes_itens) else ""

            if quantidade <= 0 or categoria_id == "":
                mensagem_erro = "Preencha quantidade e categoria para todos os itens informados."
                break

            itens_validos.append({
                "descricao": descricao,
                "quantidade": quantidade,
                "unidade": unidade,
                "categoria_id": int(categoria_id),
                "observacao": observacao_item
            })

        if not mensagem_erro and not itens_validos:
            mensagem_erro = "Adicione pelo menos um item ao pedido."

        if not mensagem_erro:
            conn = conectar()
            cursor = conn.cursor()

            data_solicitacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            cursor.execute("""
                INSERT INTO pedidos_compra
                (solicitante, data_solicitacao, status, observacao)
                VALUES (?, ?, 'Pendente', ?)
            """, (session.get("usuario"), data_solicitacao, observacao_pedido))

            pedido_id = cursor.lastrowid

            for item in itens_validos:
                cursor.execute("""
                    INSERT INTO itens_compra
                    (pedido_id, descricao, quantidade, unidade, categoria_id, observacao, status)
                    VALUES (?, ?, ?, ?, ?, ?, 'Pendente')
                """, (
                    pedido_id,
                    item["descricao"],
                    item["quantidade"],
                    item["unidade"],
                    item["categoria_id"],
                    item["observacao"]
                ))

            conn.commit()
            conn.close()

            return redirect(url_for("compras_detalhes", pedido_id=pedido_id))

    return render_template(
        "compras_novo.html",
        categorias=carregar_categorias_compra(),
        unidades=UNIDADES_COMPRA,
        mensagem_erro=mensagem_erro
    )


@app.route("/compras/pedidos", methods=["GET"])
@pode_ver_compras_obrigatorio
def compras_pedidos():
    status = request.args.get("status", "Pendente")

    if status not in STATUS_COMPRA and status != "Todos":
        status = "Pendente"

    conn = conectar()
    cursor = conn.cursor()

    if status == "Todos":
        cursor.execute("""
            SELECT
                pedidos_compra.id,
                pedidos_compra.solicitante,
                pedidos_compra.data_solicitacao,
                pedidos_compra.status,
                pedidos_compra.data_compra,
                fornecedores.nome AS fornecedor_nome,
                COUNT(itens_compra.id) AS total_itens
            FROM pedidos_compra
            LEFT JOIN itens_compra ON itens_compra.pedido_id = pedidos_compra.id
            LEFT JOIN fornecedores ON fornecedores.id = pedidos_compra.fornecedor_id
            GROUP BY pedidos_compra.id
            ORDER BY pedidos_compra.id DESC
        """)
    else:
        cursor.execute("""
            SELECT
                pedidos_compra.id,
                pedidos_compra.solicitante,
                pedidos_compra.data_solicitacao,
                pedidos_compra.status,
                pedidos_compra.data_compra,
                fornecedores.nome AS fornecedor_nome,
                COUNT(itens_compra.id) AS total_itens
            FROM pedidos_compra
            LEFT JOIN itens_compra ON itens_compra.pedido_id = pedidos_compra.id
            LEFT JOIN fornecedores ON fornecedores.id = pedidos_compra.fornecedor_id
            WHERE pedidos_compra.status = ?
            GROUP BY pedidos_compra.id
            ORDER BY pedidos_compra.id DESC
        """, (status,))

    pedidos = cursor.fetchall()
    conn.close()

    return render_template(
        "compras_pedidos.html",
        pedidos=pedidos,
        status=status,
        total_pendentes=contar_pedidos_pendentes()
    )


@app.route("/compras/pedido/<int:pedido_id>", methods=["GET"])
@pode_ver_compras_obrigatorio
def compras_detalhes(pedido_id):
    pedido = buscar_pedido_compra(pedido_id)

    if not pedido:
        return render_template(
            "compras_detalhes.html",
            pedido=None,
            itens=[],
            fornecedores=[],
            categorias=[]
        )

    return render_template(
        "compras_detalhes.html",
        pedido=pedido,
        itens=carregar_itens_pedido_compra(pedido_id),
        fornecedores=carregar_fornecedores(),
        categorias=carregar_categorias_compra(),
        produtos_estoque=carregar_produtos_estoque(),
        categorias_estoque=carregar_categorias_estoque(),
        unidades=UNIDADES_COMPRA
    )


@app.route("/compras/confirmar-item/<int:item_id>", methods=["POST"])
@compras_obrigatorio
def compras_confirmar_item(item_id):
    fornecedor_nome = request.form.get("fornecedor_nome", "").strip()
    data_compra = request.form.get("data_compra", "").strip()
    observacao_compra = request.form.get("observacao_compra", "").strip()

    fornecedor_id = criar_fornecedor_se_nao_existir(fornecedor_nome)

    if data_compra == "":
        data_compra = datetime.now().strftime("%d/%m/%Y")
    else:
        try:
            data_compra = datetime.strptime(data_compra, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            pass

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT pedido_id
        FROM itens_compra
        WHERE id = ?
    """, (item_id,))

    item = cursor.fetchone()

    if not item:
        conn.close()
        return redirect(url_for("compras_pedidos"))

    pedido_id = item["pedido_id"]

    cursor.execute("""
        UPDATE itens_compra
        SET status = 'Comprado',
            fornecedor_id = ?,
            data_compra = ?,
            observacao_compra = ?,
            comprado_por = ?
        WHERE id = ?
    """, (
        fornecedor_id,
        data_compra,
        observacao_compra,
        session.get("usuario"),
        item_id
    ))

    cursor.execute("""
        UPDATE pedidos_compra
        SET fornecedor_id = COALESCE(fornecedor_id, ?),
            data_compra = COALESCE(data_compra, ?),
            comprado_por = COALESCE(comprado_por, ?)
        WHERE id = ?
    """, (
        fornecedor_id,
        data_compra,
        session.get("usuario"),
        pedido_id
    ))

    conn.commit()
    conn.close()

    atualizar_status_pedido_compra(pedido_id)
    return redirect(url_for("compras_detalhes", pedido_id=pedido_id))


@app.route("/compras/confirmar-pedido/<int:pedido_id>", methods=["POST"])
@compras_obrigatorio
def compras_confirmar_pedido(pedido_id):
    fornecedor_nome = request.form.get("fornecedor_nome", "").strip()
    data_compra = request.form.get("data_compra", "").strip()
    observacao_compra = request.form.get("observacao_compra", "").strip()

    fornecedor_id = criar_fornecedor_se_nao_existir(fornecedor_nome)

    if data_compra == "":
        data_compra = datetime.now().strftime("%d/%m/%Y")
    else:
        try:
            data_compra = datetime.strptime(data_compra, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            pass

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE itens_compra
        SET status = 'Comprado',
            fornecedor_id = ?,
            data_compra = ?,
            observacao_compra = ?,
            comprado_por = ?
        WHERE pedido_id = ?
        AND status = 'Pendente'
    """, (
        fornecedor_id,
        data_compra,
        observacao_compra,
        session.get("usuario"),
        pedido_id
    ))

    cursor.execute("""
        UPDATE pedidos_compra
        SET status = 'Comprado',
            fornecedor_id = ?,
            data_compra = ?,
            observacao_compra = ?,
            comprado_por = ?
        WHERE id = ?
    """, (
        fornecedor_id,
        data_compra,
        observacao_compra,
        session.get("usuario"),
        pedido_id
    ))

    conn.commit()
    conn.close()

    return redirect(url_for("compras_detalhes", pedido_id=pedido_id))


@app.route("/compras/cancelar-pedido/<int:pedido_id>", methods=["POST"])
@login_obrigatorio
def compras_cancelar_pedido(pedido_id):
    pedido = buscar_pedido_compra(pedido_id)

    if not pedido:
        return redirect(url_for("compras_pedidos"))

    usuario_logado = session.get("usuario")
    perfil = session.get("perfil")

    if perfil not in ["admin", "compras"] and pedido["solicitante"] != usuario_logado:
        return render_template(
            "acesso_negado.html",
            mensagem="Você não tem permissão para cancelar este pedido."
        )

    observacao_cancelamento = request.form.get("observacao_cancelamento", "").strip()
    data_cancelamento = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE pedidos_compra
        SET status = 'Cancelado',
            data_cancelamento = ?,
            cancelado_por = ?,
            observacao_cancelamento = ?
        WHERE id = ?
    """, (data_cancelamento, usuario_logado, observacao_cancelamento, pedido_id))

    cursor.execute("""
        UPDATE itens_compra
        SET status = 'Cancelado'
        WHERE pedido_id = ?
        AND status = 'Pendente'
    """, (pedido_id,))

    conn.commit()
    conn.close()

    return redirect(url_for("compras_pedidos", status="Todos"))


@app.route("/compras/relatorios", methods=["GET"])
@admin_ou_compras_obrigatorio
def compras_relatorios():
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()
    fornecedor_id = request.args.get("fornecedor_id", "").strip()

    filtros = ["itens_compra.status = 'Comprado'"]
    parametros = []

    if data_inicio:
        filtros.append("substr(itens_compra.data_compra, 7, 4) || '-' || substr(itens_compra.data_compra, 4, 2) || '-' || substr(itens_compra.data_compra, 1, 2) >= ?")
        parametros.append(data_inicio)

    if data_fim:
        filtros.append("substr(itens_compra.data_compra, 7, 4) || '-' || substr(itens_compra.data_compra, 4, 2) || '-' || substr(itens_compra.data_compra, 1, 2) <= ?")
        parametros.append(data_fim)

    if categoria_id:
        filtros.append("itens_compra.categoria_id = ?")
        parametros.append(categoria_id)

    if fornecedor_id:
        filtros.append("itens_compra.fornecedor_id = ?")
        parametros.append(fornecedor_id)

    where_sql = " AND ".join(filtros)

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute(f"""
        SELECT
            pedidos_compra.id AS pedido_id,
            pedidos_compra.solicitante,
            itens_compra.descricao,
            itens_compra.quantidade,
            itens_compra.unidade,
            itens_compra.data_compra,
            itens_compra.observacao_compra,
            categorias_compra.nome AS categoria_nome,
            fornecedores.nome AS fornecedor_nome,
            itens_compra.comprado_por
        FROM itens_compra
        INNER JOIN pedidos_compra ON pedidos_compra.id = itens_compra.pedido_id
        INNER JOIN categorias_compra ON categorias_compra.id = itens_compra.categoria_id
        LEFT JOIN fornecedores ON fornecedores.id = itens_compra.fornecedor_id
        WHERE {where_sql}
        ORDER BY itens_compra.id DESC
    """, parametros)

    itens = cursor.fetchall()

    cursor.execute("""
        SELECT categorias_compra.nome, COUNT(itens_compra.id) AS total
        FROM itens_compra
        INNER JOIN categorias_compra ON categorias_compra.id = itens_compra.categoria_id
        WHERE itens_compra.status = 'Comprado'
        GROUP BY categorias_compra.id
        ORDER BY total DESC
    """)
    resumo_categorias = cursor.fetchall()

    cursor.execute("""
        SELECT fornecedores.nome, COUNT(itens_compra.id) AS total
        FROM itens_compra
        INNER JOIN fornecedores ON fornecedores.id = itens_compra.fornecedor_id
        WHERE itens_compra.status = 'Comprado'
        GROUP BY fornecedores.id
        ORDER BY total DESC
    """)
    resumo_fornecedores = cursor.fetchall()

    conn.close()

    return render_template(
        "compras_relatorios.html",
        itens=itens,
        categorias=carregar_categorias_compra(),
        fornecedores=carregar_fornecedores(),
        resumo_categorias=resumo_categorias,
        resumo_fornecedores=resumo_fornecedores,
        data_inicio=data_inicio,
        data_fim=data_fim,
        categoria_id=categoria_id,
        fornecedor_id=fornecedor_id
    )


@app.route("/compras/categorias", methods=["GET", "POST"])
@admin_ou_compras_obrigatorio
def compras_categorias():
    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()

        if nome == "":
            mensagem_erro = "Digite o nome da categoria."
        else:
            try:
                conn = conectar()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO categorias_compra (nome, ativo)
                    VALUES (?, 1)
                """, (nome,))
                conn.commit()
                conn.close()
                mensagem_sucesso = "Categoria cadastrada com sucesso."
            except sqlite3.IntegrityError:
                mensagem_erro = "Essa categoria já existe."

    return render_template(
        "compras_categorias.html",
        categorias=carregar_categorias_compra(apenas_ativas=False),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )


@app.route("/compras/categorias/alternar/<int:categoria_id>", methods=["POST"])
@admin_ou_compras_obrigatorio
def compras_alternar_categoria(categoria_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE categorias_compra
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (categoria_id,))

    conn.commit()
    conn.close()

    return redirect(url_for("compras_categorias"))


@app.route("/compras/fornecedores", methods=["GET", "POST"])
@admin_ou_compras_obrigatorio
def compras_fornecedores():
    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        telefone = request.form.get("telefone", "").strip()
        email = request.form.get("email", "").strip()
        observacao = request.form.get("observacao", "").strip()

        if nome == "":
            mensagem_erro = "Digite o nome do fornecedor."
        else:
            try:
                conn = conectar()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO fornecedores (nome, telefone, email, observacao, ativo)
                    VALUES (?, ?, ?, ?, 1)
                """, (nome, telefone, email, observacao))
                conn.commit()
                conn.close()
                mensagem_sucesso = "Fornecedor cadastrado com sucesso."
            except sqlite3.IntegrityError:
                mensagem_erro = "Esse fornecedor já existe."

    return render_template(
        "compras_fornecedores.html",
        fornecedores=carregar_fornecedores(apenas_ativos=False),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )


@app.route("/compras/fornecedores/alternar/<int:fornecedor_id>", methods=["POST"])
@admin_ou_compras_obrigatorio
def compras_alternar_fornecedor(fornecedor_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE fornecedores
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (fornecedor_id,))

    conn.commit()
    conn.close()

    return redirect(url_for("compras_fornecedores"))





# ============================================================
# CÉLULAS DE PRODUÇÃO E TRANSFERÊNCIAS INTERNAS - V2.3
# ============================================================

def carregar_celulas_producao(apenas_ativas=True):
    conn = conectar()
    cursor = conn.cursor()

    if apenas_ativas:
        cursor.execute("""
            SELECT *
            FROM celulas_producao
            WHERE ativo = 1
            ORDER BY nome
        """)
    else:
        cursor.execute("""
            SELECT *
            FROM celulas_producao
            ORDER BY ativo DESC, nome
        """)

    celulas = cursor.fetchall()
    conn.close()
    return celulas


def buscar_celula_producao(celula_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM celulas_producao WHERE id = ?", (celula_id,))
    celula = cursor.fetchone()
    conn.close()
    return celula


def contar_celulas_ativas():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) AS total FROM celulas_producao WHERE ativo = 1")
    total = cursor.fetchone()["total"]
    conn.close()
    return int(total or 0)


def obter_saldo_celula_produto(celula_id, produto_id, conn=None):
    propria_conexao = conn is None
    if propria_conexao:
        conn = conectar()

    cursor = conn.cursor()
    cursor.execute("""
        SELECT COALESCE(SUM(
            CASE
                WHEN tipo = 'Entrada' THEN quantidade
                WHEN tipo = 'Saída' THEN -quantidade
                ELSE 0
            END
        ), 0) AS saldo
        FROM movimentacoes_celula
        WHERE celula_id = ? AND produto_id = ?
    """, (celula_id, produto_id))
    saldo = float(cursor.fetchone()["saldo"] or 0)

    if propria_conexao:
        conn.close()
    return saldo


def carregar_saldos_celulas(celula_id="", busca="", categoria_id=""):
    filtros = ["celulas_producao.ativo = 1", "produtos_estoque.ativo = 1"]
    parametros = []

    celula_id = str(celula_id or "").strip()
    busca = str(busca or "").strip()
    categoria_id = str(categoria_id or "").strip()

    if celula_id:
        filtros.append("celulas_producao.id = ?")
        parametros.append(celula_id)

    if busca:
        filtros.append("(produtos_estoque.nome LIKE ? OR produtos_estoque.codigo LIKE ?)")
        termo = f"%{busca}%"
        parametros.extend([termo, termo])

    if categoria_id:
        filtros.append("produtos_estoque.categoria_id = ?")
        parametros.append(categoria_id)

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT
            celulas_producao.id AS celula_id,
            celulas_producao.nome AS celula_nome,
            produtos_estoque.id AS produto_id,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS unidade,
            categorias_estoque.nome AS categoria_nome,
            COALESCE(SUM(
                CASE
                    WHEN movimentacoes_celula.tipo = 'Entrada' THEN movimentacoes_celula.quantidade
                    WHEN movimentacoes_celula.tipo = 'Saída' THEN -movimentacoes_celula.quantidade
                    ELSE 0
                END
            ), 0) AS saldo
        FROM movimentacoes_celula
        INNER JOIN celulas_producao ON celulas_producao.id = movimentacoes_celula.celula_id
        INNER JOIN produtos_estoque ON produtos_estoque.id = movimentacoes_celula.produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE {' AND '.join(filtros)}
        GROUP BY celulas_producao.id, produtos_estoque.id
        HAVING ABS(saldo) > 0.0000001
        ORDER BY celulas_producao.nome, categorias_estoque.nome, produtos_estoque.nome
    """, parametros)
    saldos = cursor.fetchall()
    conn.close()
    return saldos


def quantidade_transferida_separacao_item(separacao_item_id, conn=None):
    propria_conexao = conn is None
    if propria_conexao:
        conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COALESCE(SUM(itens_transferencia_interna.quantidade), 0) AS total
        FROM itens_transferencia_interna
        INNER JOIN transferencias_internas
            ON transferencias_internas.id = itens_transferencia_interna.transferencia_id
        WHERE itens_transferencia_interna.separacao_item_id = ?
          AND transferencias_internas.status = 'Confirmada'
    """, (separacao_item_id,))
    total = float(cursor.fetchone()["total"] or 0)
    if propria_conexao:
        conn.close()
    return total


def carregar_itens_separacao_para_transferencia(separacao_id):
    itens = carregar_itens_separacao_estoque(separacao_id)
    resultado = []
    for item_original in itens:
        item = dict(item_original)
        transferido = quantidade_transferida_separacao_item(item["id"])
        separado = float(item["quantidade_separada"] or 0)
        item["quantidade_transferida"] = transferido
        item["quantidade_a_transferir"] = max(separado - transferido, 0)
        resultado.append(item)
    return resultado


def carregar_transferencias_internas(status="", tipo="", celula_id=""):
    filtros = []
    parametros = []

    if status:
        filtros.append("transferencias_internas.status = ?")
        parametros.append(status)
    if tipo:
        filtros.append("transferencias_internas.tipo = ?")
        parametros.append(tipo)
    if celula_id:
        filtros.append("(transferencias_internas.celula_origem_id = ? OR transferencias_internas.celula_destino_id = ?)")
        parametros.extend([celula_id, celula_id])

    where_sql = "WHERE " + " AND ".join(filtros) if filtros else ""

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT
            transferencias_internas.*,
            origem.nome AS celula_origem_nome,
            destino.nome AS celula_destino_nome,
            COUNT(itens_transferencia_interna.id) AS total_itens,
            COALESCE(SUM(itens_transferencia_interna.quantidade), 0) AS total_quantidade
        FROM transferencias_internas
        LEFT JOIN celulas_producao origem ON origem.id = transferencias_internas.celula_origem_id
        LEFT JOIN celulas_producao destino ON destino.id = transferencias_internas.celula_destino_id
        LEFT JOIN itens_transferencia_interna ON itens_transferencia_interna.transferencia_id = transferencias_internas.id
        {where_sql}
        GROUP BY transferencias_internas.id
        ORDER BY transferencias_internas.id DESC
    """, parametros)
    transferencias = cursor.fetchall()
    conn.close()
    return transferencias


def contar_transferencias_rascunho():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) AS total FROM transferencias_internas WHERE status = 'Rascunho'")
    total = cursor.fetchone()["total"]
    conn.close()
    return int(total or 0)


def buscar_transferencia_interna(transferencia_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            transferencias_internas.*,
            origem.nome AS celula_origem_nome,
            destino.nome AS celula_destino_nome
        FROM transferencias_internas
        LEFT JOIN celulas_producao origem ON origem.id = transferencias_internas.celula_origem_id
        LEFT JOIN celulas_producao destino ON destino.id = transferencias_internas.celula_destino_id
        WHERE transferencias_internas.id = ?
    """, (transferencia_id,))
    transferencia = cursor.fetchone()
    conn.close()
    return transferencia


def carregar_itens_transferencia_interna(transferencia_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            itens_transferencia_interna.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao AS produto_unidade,
            categorias_estoque.nome AS categoria_nome
        FROM itens_transferencia_interna
        INNER JOIN produtos_estoque ON produtos_estoque.id = itens_transferencia_interna.produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        WHERE itens_transferencia_interna.transferencia_id = ?
        ORDER BY produtos_estoque.nome
    """, (transferencia_id,))
    itens = cursor.fetchall()
    conn.close()
    return itens


def criar_transferencia_manual(tipo, celula_id, produto_id, quantidade, observacao=""):
    if tipo not in TIPOS_TRANSFERENCIA_INTERNA:
        raise ValueError("Tipo de transferência inválido.")

    try:
        celula_id = int(celula_id)
        produto_id = int(produto_id)
        quantidade = float(str(quantidade or 0).replace(",", "."))
    except Exception:
        raise ValueError("Informe célula, produto e quantidade válidos.")

    if quantidade <= 0:
        raise ValueError("A quantidade deve ser maior que zero.")

    celula = buscar_celula_producao(celula_id)
    produto = buscar_produto_estoque(produto_id)
    if not celula or int(celula["ativo"] or 0) != 1:
        raise ValueError("Célula inválida ou inativa.")
    if not produto or int(produto["ativo"] or 0) != 1:
        raise ValueError("Produto inválido ou inativo.")

    celula_origem_id = celula_id if tipo == "Célula para Estoque" else None
    celula_destino_id = celula_id if tipo == "Estoque para Célula" else None

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO transferencias_internas
        (tipo, celula_origem_id, celula_destino_id, status, data_criacao, criado_por, observacao)
        VALUES (?, ?, ?, 'Rascunho', ?, ?, ?)
    """, (tipo, celula_origem_id, celula_destino_id, _agora_texto(), _usuario_atual(), observacao.strip()))
    transferencia_id = cursor.lastrowid
    cursor.execute("""
        INSERT INTO itens_transferencia_interna
        (transferencia_id, produto_id, quantidade, unidade, observacao)
        VALUES (?, ?, ?, ?, ?)
    """, (transferencia_id, produto_id, quantidade, produto["unidade_padrao"], observacao.strip()))
    conn.commit()
    conn.close()
    return transferencia_id


def criar_transferencia_da_separacao(separacao_id, celula_destino_id, form):
    separacao = buscar_separacao_estoque(separacao_id)
    if not separacao:
        raise ValueError("Separação não encontrada.")
    if separacao["status"] in ["Cancelado", "Encerrado"]:
        raise ValueError("Esta separação não aceita novas transferências.")

    try:
        celula_destino_id = int(celula_destino_id)
    except Exception:
        raise ValueError("Selecione a célula de destino.")

    celula = buscar_celula_producao(celula_destino_id)
    if not celula or int(celula["ativo"] or 0) != 1:
        raise ValueError("Célula de destino inválida ou inativa.")

    itens = carregar_itens_separacao_para_transferencia(separacao_id)
    itens_escolhidos = []
    for item in itens:
        campo = f"quantidade_transferir_{item['id']}"
        bruto = str(form.get(campo, "0") or "0").strip().replace(",", ".")
        try:
            quantidade = float(bruto)
        except Exception:
            quantidade = 0
        restante = float(item["quantidade_a_transferir"] or 0)
        if quantidade < -0.0000001:
            raise ValueError(f"Quantidade negativa para {item['produto_nome']}.")
        if quantidade > restante + 0.0000001:
            raise ValueError(
                f"A quantidade de {item['produto_nome']} excede o separado ainda não transferido "
                f"({restante:.4f} {item['unidade']})."
            )
        if quantidade > 0:
            itens_escolhidos.append((item, quantidade))

    if not itens_escolhidos:
        raise ValueError("Informe pelo menos uma quantidade para transferir.")

    observacao = str(form.get("observacao", "") or "").strip()
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO transferencias_internas
        (tipo, celula_destino_id, separacao_id, status, data_criacao, criado_por, observacao)
        VALUES ('Estoque para Célula', ?, ?, 'Rascunho', ?, ?, ?)
    """, (celula_destino_id, separacao_id, _agora_texto(), _usuario_atual(), observacao))
    transferencia_id = cursor.lastrowid

    for item, quantidade in itens_escolhidos:
        cursor.execute("""
            INSERT INTO itens_transferencia_interna
            (transferencia_id, produto_id, quantidade, unidade, separacao_item_id, observacao)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            transferencia_id, item["produto_id"], quantidade,
            item["unidade"], item["id"], item["observacao"] or ""
        ))

    conn.commit()
    conn.close()
    return transferencia_id


def _saldo_estoque_conn(cursor, produto_id):
    cursor.execute("""
        SELECT COALESCE(SUM(CASE WHEN tipo = 'Entrada' THEN quantidade WHEN tipo = 'Saída' THEN -quantidade ELSE 0 END), 0) AS saldo
        FROM movimentacoes_estoque
        WHERE produto_id = ?
    """, (produto_id,))
    return float(cursor.fetchone()["saldo"] or 0)


def _saldo_celula_conn(cursor, celula_id, produto_id):
    cursor.execute("""
        SELECT COALESCE(SUM(CASE WHEN tipo = 'Entrada' THEN quantidade WHEN tipo = 'Saída' THEN -quantidade ELSE 0 END), 0) AS saldo
        FROM movimentacoes_celula
        WHERE celula_id = ? AND produto_id = ?
    """, (celula_id, produto_id))
    return float(cursor.fetchone()["saldo"] or 0)


def _atualizar_encerramento_separacao(separacao_id, cursor, usuario):
    if not separacao_id:
        return
    cursor.execute("""
        SELECT id, quantidade_prevista, quantidade_separada
        FROM itens_separacao_estoque
        WHERE separacao_id = ?
    """, (separacao_id,))
    itens = cursor.fetchall()
    if not itens:
        return

    todos_transferidos = True
    for item in itens:
        previsto = float(item["quantidade_prevista"] or 0)
        separado = float(item["quantidade_separada"] or 0)

        # A separação só encerra quando tudo que foi previsto estiver fisicamente separado
        # e todo o separado tiver sido transferido para uma célula.
        if separado + 0.0000001 < previsto:
            todos_transferidos = False
            break

        cursor.execute("""
            SELECT COALESCE(SUM(iti.quantidade), 0) AS total
            FROM itens_transferencia_interna iti
            INNER JOIN transferencias_internas ti ON ti.id = iti.transferencia_id
            WHERE iti.separacao_item_id = ? AND ti.status = 'Confirmada'
        """, (item["id"],))
        transferido = float(cursor.fetchone()["total"] or 0)
        if transferido + 0.0000001 < separado:
            todos_transferidos = False
            break

    if todos_transferidos:
        agora = _agora_texto()
        cursor.execute("""
            UPDATE separacoes_estoque
            SET status = 'Encerrado', data_atualizacao = ?, atualizado_por = ?,
                data_conclusao = ?, concluido_por = ?
            WHERE id = ? AND status != 'Cancelado'
        """, (agora, usuario, agora, usuario, separacao_id))


def confirmar_transferencia_interna(transferencia_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("BEGIN IMMEDIATE")
        cursor.execute("SELECT * FROM transferencias_internas WHERE id = ?", (transferencia_id,))
        transferencia = cursor.fetchone()
        if not transferencia:
            raise ValueError("Transferência não encontrada.")
        if transferencia["status"] != "Rascunho":
            raise ValueError("Somente transferências em rascunho podem ser confirmadas.")

        cursor.execute("""
            SELECT iti.*, pe.nome AS produto_nome, pe.unidade_padrao
            FROM itens_transferencia_interna iti
            INNER JOIN produtos_estoque pe ON pe.id = iti.produto_id
            WHERE iti.transferencia_id = ?
            ORDER BY iti.id
        """, (transferencia_id,))
        itens = cursor.fetchall()
        if not itens:
            raise ValueError("A transferência não possui itens.")

        usuario = _usuario_atual()
        data_mov = datetime.now().strftime("%d/%m/%Y")

        # Valida todos antes de movimentar.
        for item in itens:
            quantidade = float(item["quantidade"] or 0)
            if quantidade <= 0:
                raise ValueError(f"Quantidade inválida para {item['produto_nome']}.")
            if transferencia["tipo"] == "Estoque para Célula":
                saldo = _saldo_estoque_conn(cursor, item["produto_id"])
                if quantidade > saldo + 0.0000001:
                    raise ValueError(
                        f"Saldo central insuficiente para {item['produto_nome']}. "
                        f"Disponível: {saldo:.4f} {item['unidade_padrao']}; solicitado: {quantidade:.4f}."
                    )
            elif transferencia["tipo"] == "Célula para Estoque":
                saldo = _saldo_celula_conn(cursor, transferencia["celula_origem_id"], item["produto_id"])
                if quantidade > saldo + 0.0000001:
                    raise ValueError(
                        f"Saldo insuficiente na célula para {item['produto_nome']}. "
                        f"Disponível: {saldo:.4f} {item['unidade_padrao']}; solicitado: {quantidade:.4f}."
                    )
            else:
                raise ValueError("Tipo de transferência inválido.")

        for item in itens:
            quantidade = float(item["quantidade"] or 0)
            observacao = item["observacao"] or transferencia["observacao"] or ""

            if transferencia["tipo"] == "Estoque para Célula":
                cursor.execute("""
                    INSERT INTO movimentacoes_estoque
                    (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_nome, fator_conversao,
                     data_movimentacao, origem, origem_id, usuario, observacao)
                    VALUES (?, 'Saída', ?, ?, ?, 1, ?, 'Transferência para Célula', ?, ?, ?)
                """, (
                    item["produto_id"], quantidade, quantidade, item["unidade_padrao"],
                    data_mov, transferencia_id, usuario, observacao
                ))
                estoque_mov_id = cursor.lastrowid

                cursor.execute("""
                    INSERT INTO movimentacoes_celula
                    (celula_id, produto_id, tipo, quantidade, data_movimentacao, origem, origem_id, usuario, observacao)
                    VALUES (?, ?, 'Entrada', ?, ?, 'Transferência Interna', ?, ?, ?)
                """, (
                    transferencia["celula_destino_id"], item["produto_id"], quantidade,
                    data_mov, transferencia_id, usuario, observacao
                ))
                celula_mov_id = cursor.lastrowid
            else:
                cursor.execute("""
                    INSERT INTO movimentacoes_celula
                    (celula_id, produto_id, tipo, quantidade, data_movimentacao, origem, origem_id, usuario, observacao)
                    VALUES (?, ?, 'Saída', ?, ?, 'Retorno ao Estoque', ?, ?, ?)
                """, (
                    transferencia["celula_origem_id"], item["produto_id"], quantidade,
                    data_mov, transferencia_id, usuario, observacao
                ))
                celula_mov_id = cursor.lastrowid

                cursor.execute("""
                    INSERT INTO movimentacoes_estoque
                    (produto_id, tipo, quantidade, quantidade_embalagem, embalagem_nome, fator_conversao,
                     data_movimentacao, origem, origem_id, usuario, observacao)
                    VALUES (?, 'Entrada', ?, ?, ?, 1, ?, 'Retorno de Célula', ?, ?, ?)
                """, (
                    item["produto_id"], quantidade, quantidade, item["unidade_padrao"],
                    data_mov, transferencia_id, usuario, observacao
                ))
                estoque_mov_id = cursor.lastrowid

            cursor.execute("""
                UPDATE itens_transferencia_interna
                SET estoque_movimentacao_id = ?, celula_movimentacao_id = ?
                WHERE id = ?
            """, (estoque_mov_id, celula_mov_id, item["id"]))

        cursor.execute("""
            UPDATE transferencias_internas
            SET status = 'Confirmada', data_confirmacao = ?, confirmado_por = ?
            WHERE id = ?
        """, (_agora_texto(), usuario, transferencia_id))

        _atualizar_encerramento_separacao(transferencia["separacao_id"], cursor, usuario)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def cancelar_transferencia_interna(transferencia_id):
    transferencia = buscar_transferencia_interna(transferencia_id)
    if not transferencia:
        raise ValueError("Transferência não encontrada.")
    if transferencia["status"] != "Rascunho":
        raise ValueError("Somente transferências em rascunho podem ser canceladas.")

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE transferencias_internas
        SET status = 'Cancelada', data_cancelamento = ?, cancelado_por = ?
        WHERE id = ?
    """, (_agora_texto(), _usuario_atual(), transferencia_id))
    conn.commit()
    conn.close()


# ============================================================
# ROTAS DO MÓDULO DE ESTOQUE
# ============================================================

@app.route("/estoque", methods=["GET"])
@estoque_obrigatorio
def estoque_inicio():
    return render_template(
        "estoque_home.html",
        total_produtos=len(carregar_produtos_estoque()),
        total_baixo=contar_produtos_baixo_estoque(),
        total_separacoes=contar_separacoes_pendentes(),
        total_celulas=contar_celulas_ativas(),
        total_transferencias=contar_transferencias_rascunho(),
        resumo_validade=contar_lotes_validade_resumo(),
        total_solicitacoes_internas=contar_solicitacoes_internas_pendentes()
    )



@app.route("/estoque/celulas", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_celulas():
    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = str(request.form.get("nome", "") or "").strip()
        descricao = str(request.form.get("descricao", "") or "").strip()
        centro_custo = str(request.form.get("centro_custo", "") or "").strip()

        if not nome:
            mensagem_erro = "Informe o nome da célula de produção."
        else:
            try:
                conn = conectar()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO celulas_producao
                    (nome, descricao, centro_custo, ativo, data_cadastro, criado_por)
                    VALUES (?, ?, ?, 1, ?, ?)
                """, (nome, descricao, centro_custo, _agora_texto(), _usuario_atual()))
                conn.commit()
                conn.close()
                mensagem_sucesso = "Célula cadastrada com sucesso."
            except sqlite3.IntegrityError:
                mensagem_erro = "Já existe uma célula com esse nome."

    return render_template(
        "estoque_celulas.html",
        celulas=carregar_celulas_producao(apenas_ativas=False),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )


@app.route("/estoque/celulas/<int:celula_id>/alternar", methods=["POST"])
@estoque_obrigatorio
def estoque_celula_alternar(celula_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE celulas_producao
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (celula_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("estoque_celulas"))


@app.route("/estoque/celulas/saldos", methods=["GET"])
@estoque_obrigatorio
def estoque_celulas_saldos():
    celula_id = request.args.get("celula_id", "").strip()
    busca = request.args.get("busca", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()
    return render_template(
        "estoque_celulas_saldos.html",
        saldos=carregar_saldos_celulas(celula_id, busca, categoria_id),
        celulas=carregar_celulas_producao(),
        categorias=carregar_categorias_estoque(),
        celula_id=celula_id,
        busca=busca,
        categoria_id=categoria_id
    )


@app.route("/estoque/transferencias", methods=["GET"])
@estoque_obrigatorio
def estoque_transferencias():
    status = request.args.get("status", "").strip()
    tipo = request.args.get("tipo", "").strip()
    celula_id = request.args.get("celula_id", "").strip()
    return render_template(
        "estoque_transferencias.html",
        transferencias=carregar_transferencias_internas(status, tipo, celula_id),
        celulas=carregar_celulas_producao(),
        tipos=TIPOS_TRANSFERENCIA_INTERNA,
        status_opcoes=STATUS_TRANSFERENCIA_INTERNA,
        status=status,
        tipo=tipo,
        celula_id=celula_id
    )


@app.route("/estoque/transferencias/nova", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_transferencia_nova():
    erro = None
    if request.method == "POST":
        try:
            transferencia_id = criar_transferencia_manual(
                request.form.get("tipo", ""),
                request.form.get("celula_id", ""),
                request.form.get("produto_id", ""),
                request.form.get("quantidade", ""),
                request.form.get("observacao", "")
            )
            return redirect(url_for("estoque_transferencia_detalhes", transferencia_id=transferencia_id))
        except Exception as e:
            erro = str(e)

    return render_template(
        "estoque_transferencia_form.html",
        celulas=carregar_celulas_producao(),
        produtos=carregar_produtos_estoque(),
        tipos=TIPOS_TRANSFERENCIA_INTERNA,
        erro=erro
    )


@app.route("/estoque/separacoes/<int:separacao_id>/transferir", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_separacao_transferir(separacao_id):
    separacao = buscar_separacao_estoque(separacao_id)
    if not separacao:
        return redirect(url_for("estoque_separacoes"))

    erro = None
    if request.method == "POST":
        try:
            transferencia_id = criar_transferencia_da_separacao(
                separacao_id,
                request.form.get("celula_destino_id", ""),
                request.form
            )
            return redirect(url_for("estoque_transferencia_detalhes", transferencia_id=transferencia_id))
        except Exception as e:
            erro = str(e)

    return render_template(
        "estoque_transferencia_separacao.html",
        separacao=separacao,
        itens=carregar_itens_separacao_para_transferencia(separacao_id),
        celulas=carregar_celulas_producao(),
        erro=erro
    )


@app.route("/estoque/transferencias/<int:transferencia_id>", methods=["GET"])
@estoque_obrigatorio
def estoque_transferencia_detalhes(transferencia_id):
    transferencia = buscar_transferencia_interna(transferencia_id)
    if not transferencia:
        return redirect(url_for("estoque_transferencias"))
    return render_template(
        "estoque_transferencia_detalhes.html",
        transferencia=transferencia,
        itens=carregar_itens_transferencia_interna(transferencia_id),
        mensagem=request.args.get("mensagem", ""),
        erro=request.args.get("erro", "")
    )


@app.route("/estoque/transferencias/<int:transferencia_id>/confirmar", methods=["POST"])
@estoque_obrigatorio
def estoque_transferencia_confirmar(transferencia_id):
    try:
        confirmar_transferencia_interna(transferencia_id)
        return redirect(url_for(
            "estoque_transferencia_detalhes",
            transferencia_id=transferencia_id,
            mensagem="Transferência confirmada. O saldo saiu do estoque central e entrou na célula correspondente."
        ))
    except Exception as e:
        return redirect(url_for(
            "estoque_transferencia_detalhes",
            transferencia_id=transferencia_id,
            erro=str(e)
        ))


@app.route("/estoque/transferencias/<int:transferencia_id>/cancelar", methods=["POST"])
@estoque_obrigatorio
def estoque_transferencia_cancelar(transferencia_id):
    try:
        cancelar_transferencia_interna(transferencia_id)
        return redirect(url_for(
            "estoque_transferencia_detalhes",
            transferencia_id=transferencia_id,
            mensagem="Transferência cancelada sem movimentar saldos."
        ))
    except Exception as e:
        return redirect(url_for(
            "estoque_transferencia_detalhes",
            transferencia_id=transferencia_id,
            erro=str(e)
        ))


@app.route("/estoque/produtos", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_produtos():
    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        categoria_id = request.form.get("categoria_id", "").strip()
        unidade_padrao = request.form.get("unidade_padrao", "").strip()
        estoque_minimo = request.form.get("estoque_minimo", "0").strip()
        custo_padrao = request.form.get("custo_padrao", "0").strip()
        ativo_venda = 1 if request.form.get("ativo_venda") == "1" else 0
        forma_abastecimento = request.form.get("forma_abastecimento", "Separado diretamente do estoque").strip()
        origem_expedicao_id = request.form.get("origem_expedicao_id", "").strip()
        ordem_loja = request.form.get("ordem_loja", "0").strip()
        controla_validade = 1 if request.form.get("controla_validade") == "1" else 0
        dias_validade = request.form.get("dias_validade", "").strip()
        observacao = request.form.get("observacao", "").strip()

        if nome == "" or unidade_padrao == "":
            mensagem_erro = "Preencha nome e unidade padrão do produto."
        elif controla_validade and (not str(dias_validade).isdigit() or int(dias_validade) <= 0):
            mensagem_erro = "Informe os dias de validade para produtos controlados por lote."
        else:
            try:
                criar_produto_estoque(
                    nome, categoria_id, unidade_padrao, estoque_minimo, observacao,
                    custo_padrao, ativo_venda, controla_validade, dias_validade
                )
                mensagem_sucesso = "Produto cadastrado com sucesso."
            except sqlite3.IntegrityError:
                mensagem_erro = "Esse produto já existe no estoque."
            except Exception as e:
                mensagem_erro = f"Erro ao cadastrar produto: {str(e)}"

    return render_template(
        "estoque_produtos.html",
        produtos=carregar_produtos_estoque(apenas_ativos=False),
        categorias=carregar_categorias_estoque(),
        unidades=UNIDADES_COMPRA,
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )


@app.route("/estoque/produtos/alternar/<int:produto_id>", methods=["POST"])
@estoque_obrigatorio
def estoque_alternar_produto(produto_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE produtos_estoque
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (produto_id,))

    conn.commit()
    conn.close()
    return redirect(url_for("estoque_produtos"))



@app.route("/estoque/produtos/editar/<int:produto_id>", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_editar_produto(produto_id):
    produto = buscar_produto_estoque(produto_id)

    if not produto:
        return redirect(url_for("estoque_produtos"))

    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        categoria_id = request.form.get("categoria_id", "").strip()
        unidade_padrao = request.form.get("unidade_padrao", "").strip()
        estoque_minimo = request.form.get("estoque_minimo", "0").strip()
        custo_padrao = request.form.get("custo_padrao", "0").strip()
        metodo_custo = request.form.get("metodo_custo", "Automático").strip()
        if metodo_custo not in METODOS_CUSTO:
            metodo_custo = "Automático"
        ativo = 1 if request.form.get("ativo") == "1" else 0
        ativo_venda = 1 if request.form.get("ativo_venda") == "1" else 0
        forma_abastecimento = request.form.get("forma_abastecimento", "Separado diretamente do estoque").strip()
        origem_expedicao_id = request.form.get("origem_expedicao_id", "").strip()
        ordem_loja = request.form.get("ordem_loja", "0").strip()
        controla_validade = 1 if request.form.get("controla_validade") == "1" else 0
        dias_validade = request.form.get("dias_validade", "").strip()
        observacao = request.form.get("observacao", "").strip()

        if nome == "" or unidade_padrao == "":
            mensagem_erro = "Preencha nome e unidade padrão do produto."
        elif controla_validade and (not str(dias_validade).isdigit() or int(dias_validade) <= 0):
            mensagem_erro = "Informe os dias de validade para produtos controlados por lote."
        else:
            try:
                estoque_minimo = float(str(estoque_minimo or 0).replace(",", "."))
            except:
                estoque_minimo = 0

            try:
                custo_padrao = float(str(custo_padrao or 0).replace(",", "."))
            except:
                custo_padrao = 0

            conn = conectar()
            cursor = conn.cursor()

            try:
                cursor.execute("""
                    UPDATE produtos_estoque
                    SET nome = ?,
                        categoria_id = ?,
                        unidade_padrao = ?,
                        estoque_minimo = ?,
                        custo_padrao = ?,
                        metodo_custo = ?,
                        ativo = ?,
                        ativo_venda = ?,
                        forma_abastecimento = ?,
                        origem_expedicao_id = ?,
                        ordem_loja = ?,
                        controla_validade = ?,
                        dias_validade = ?,
                        observacao = ?
                    WHERE id = ?
                """, (
                    nome,
                    int(categoria_id) if categoria_id else None,
                    unidade_padrao,
                    estoque_minimo,
                    custo_padrao,
                    metodo_custo,
                    ativo,
                    ativo_venda,
                    forma_abastecimento if forma_abastecimento in FORMAS_ABASTECIMENTO else "Separado diretamente do estoque",
                    int(origem_expedicao_id) if origem_expedicao_id else None,
                    int(ordem_loja) if str(ordem_loja or "0").lstrip("-").isdigit() else 0,
                    controla_validade,
                    int(dias_validade) if str(dias_validade or "").isdigit() else None,
                    observacao,
                    produto_id
                ))

                conn.commit()
                mensagem_sucesso = "Produto atualizado com sucesso."
                produto = buscar_produto_estoque(produto_id)

            except sqlite3.IntegrityError:
                mensagem_erro = "Já existe outro produto com esse nome."
            except Exception as e:
                mensagem_erro = f"Erro ao atualizar produto: {str(e)}"
            finally:
                conn.close()

    return render_template(
        "estoque_produto_editar.html",
        produto=produto,
        categorias=carregar_categorias_estoque(apenas_ativas=False),
        unidades=UNIDADES_COMPRA,
        embalagens=carregar_embalagens_produto(produto_id, apenas_ativas=False),
        resumo_produto=montar_resumo_produto_estoque(produto),
        ultimas_movimentacoes=carregar_ultimas_movimentacoes_produto(produto_id),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro,
        origens_expedicao=carregar_origens_expedicao(apenas_ativas=False),
        formas_abastecimento=FORMAS_ABASTECIMENTO,
        metodos_custo=METODOS_CUSTO,
        custo_calculado_atual=calcular_custo_produto(produto_id)
    )


@app.route("/estoque/produtos/<int:produto_id>/embalagens", methods=["POST"])
@estoque_obrigatorio
def estoque_adicionar_embalagem(produto_id):
    produto = buscar_produto_estoque(produto_id)

    if not produto:
        return redirect(url_for("estoque_produtos"))

    nome = request.form.get("nome_embalagem", "").strip()
    fator_conversao = request.form.get("fator_conversao", "1").strip()
    observacao = request.form.get("observacao_embalagem", "").strip()

    try:
        criar_embalagem_produto(produto_id, nome, fator_conversao, padrao=0, observacao=observacao)
    except Exception:
        pass

    return redirect(url_for("estoque_editar_produto", produto_id=produto_id))


@app.route("/estoque/produtos/embalagens/alternar/<int:embalagem_id>", methods=["POST"])
@estoque_obrigatorio
def estoque_alternar_embalagem(embalagem_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT produto_id
        FROM produto_embalagens
        WHERE id = ?
    """, (embalagem_id,))

    embalagem = cursor.fetchone()

    if not embalagem:
        conn.close()
        return redirect(url_for("estoque_produtos"))

    produto_id = embalagem["produto_id"]

    cursor.execute("""
        UPDATE produto_embalagens
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (embalagem_id,))

    conn.commit()
    conn.close()

    return redirect(url_for("estoque_editar_produto", produto_id=produto_id))


@app.route("/estoque/saldo", methods=["GET"])
@estoque_obrigatorio
def estoque_saldo():
    busca = request.args.get("busca", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()
    status = request.args.get("status", "").strip()

    return render_template(
        "estoque_saldo.html",
        saldos=carregar_saldo_estoque(busca, categoria_id, status),
        categorias=carregar_categorias_estoque(apenas_ativas=False),
        resumo_categorias=carregar_resumo_saldo_por_categoria(),
        filtros={
            "busca": busca,
            "categoria_id": categoria_id,
            "status": status
        }
    )


@app.route("/estoque/movimentacao", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_movimentacao():
    mensagem_sucesso = None
    mensagem_erro = None
    produto_id_selecionado = request.args.get("produto_id", "").strip()

    if request.method == "POST":
        produto_id = request.form.get("produto_id", "").strip()
        tipo = request.form.get("tipo", "Entrada").strip()
        quantidade = request.form.get("quantidade", "0").strip()
        embalagem_id = request.form.get("embalagem_id", "").strip()
        data_movimentacao = request.form.get("data_movimentacao", "").strip()
        origem = request.form.get("origem", "Entrada Manual").strip()
        observacao = request.form.get("observacao", "").strip()
        permitir_saldo_negativo = (
            request.form.get("permitir_saldo_negativo") == "sim"
            and session.get("perfil") == "admin"
        )
        justificativa_saldo_negativo = request.form.get("justificativa_saldo_negativo", "").strip()

        if produto_id == "":
            mensagem_erro = "Selecione um produto cadastrado."
        else:
            try:
                registrar_movimentacao_estoque(
                    int(produto_id),
                    tipo,
                    quantidade,
                    data_movimentacao,
                    origem,
                    embalagem_id=int(embalagem_id) if embalagem_id else None,
                    observacao=observacao,
                    permitir_saldo_negativo=permitir_saldo_negativo,
                    justificativa_saldo_negativo=justificativa_saldo_negativo
                )
                mensagem_sucesso = "Movimentação registrada com sucesso."
            except Exception as e:
                mensagem_erro = str(e)

    return render_template(
        "estoque_movimentacao.html",
        produtos=carregar_produtos_estoque(),
        embalagens_por_produto=carregar_embalagens_todos_produtos(),
        tipos=TIPOS_MOVIMENTACAO_ESTOQUE,
        origens=ORIGENS_ESTOQUE,
        data_hoje=datetime.now().strftime("%Y-%m-%d"),
        produto_id_selecionado=produto_id_selecionado,
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )


@app.route("/estoque/historico", methods=["GET"])
@estoque_obrigatorio
def estoque_historico():
    produto_id = request.args.get("produto_id", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()
    tipo = request.args.get("tipo", "").strip()
    origem = request.args.get("origem", "").strip()
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()

    filtros_sql = []
    parametros = []

    if produto_id:
        filtros_sql.append("movimentacoes_estoque.produto_id = ?")
        parametros.append(produto_id)

    if categoria_id:
        filtros_sql.append("produtos_estoque.categoria_id = ?")
        parametros.append(categoria_id)

    if tipo:
        filtros_sql.append("movimentacoes_estoque.tipo = ?")
        parametros.append(tipo)

    if origem:
        filtros_sql.append("movimentacoes_estoque.origem = ?")
        parametros.append(origem)

    where_sql = ""
    if filtros_sql:
        where_sql = "WHERE " + " AND ".join(filtros_sql)

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute(f"""
        SELECT
            movimentacoes_estoque.*,
            produtos_estoque.codigo AS produto_codigo,
            produtos_estoque.nome AS produto_nome,
            produtos_estoque.unidade_padrao,
            categorias_estoque.nome AS categoria_nome,
            fornecedores.nome AS fornecedor_nome
        FROM movimentacoes_estoque
        INNER JOIN produtos_estoque ON produtos_estoque.id = movimentacoes_estoque.produto_id
        LEFT JOIN categorias_estoque ON categorias_estoque.id = produtos_estoque.categoria_id
        LEFT JOIN fornecedores ON fornecedores.id = movimentacoes_estoque.fornecedor_id
        {where_sql}
        ORDER BY movimentacoes_estoque.id DESC
        LIMIT 500
    """, parametros)

    movimentacoes_base = cursor.fetchall()
    conn.close()

    movimentacoes = []
    data_inicio_dt = None
    data_fim_dt = None

    if data_inicio:
        try:
            data_inicio_dt = datetime.strptime(data_inicio, "%Y-%m-%d")
        except:
            data_inicio_dt = None

    if data_fim:
        try:
            data_fim_dt = datetime.strptime(data_fim, "%Y-%m-%d")
        except:
            data_fim_dt = None

    for mov in movimentacoes_base:
        incluir = True

        try:
            data_mov_dt = datetime.strptime(mov["data_movimentacao"], "%d/%m/%Y")
        except:
            data_mov_dt = None

        if data_inicio_dt and data_mov_dt and data_mov_dt < data_inicio_dt:
            incluir = False

        if data_fim_dt and data_mov_dt and data_mov_dt > data_fim_dt:
            incluir = False

        if incluir:
            movimentacoes.append(mov)

    total_entradas = 0
    total_saidas = 0

    for mov in movimentacoes:
        if mov["tipo"] == "Entrada":
            total_entradas += mov["quantidade"]
        elif mov["tipo"] == "Saída":
            total_saidas += mov["quantidade"]

    return render_template(
        "estoque_historico.html",
        movimentacoes=movimentacoes,
        produtos=carregar_produtos_estoque(),
        categorias=carregar_categorias_estoque(apenas_ativas=False),
        tipos=TIPOS_MOVIMENTACAO_ESTOQUE,
        origens=ORIGENS_ESTOQUE,
        filtros={
            "produto_id": produto_id,
            "categoria_id": categoria_id,
            "tipo": tipo,
            "origem": origem,
            "data_inicio": data_inicio,
            "data_fim": data_fim
        },
        resumo={
            "total_movimentacoes": len(movimentacoes),
            "total_entradas": total_entradas,
            "total_saidas": total_saidas
        }
    )


@app.route("/estoque/categorias", methods=["GET", "POST"])
@estoque_obrigatorio
def estoque_categorias():
    mensagem_sucesso = None
    mensagem_erro = None

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()

        if nome == "":
            mensagem_erro = "Digite o nome da categoria."
        else:
            try:
                conn = conectar()
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO categorias_estoque (nome, ativo)
                    VALUES (?, 1)
                """, (nome,))
                conn.commit()
                conn.close()
                mensagem_sucesso = "Categoria cadastrada com sucesso."
            except sqlite3.IntegrityError:
                mensagem_erro = "Essa categoria já existe."

    return render_template(
        "estoque_categorias.html",
        categorias=carregar_categorias_estoque(apenas_ativas=False),
        mensagem_sucesso=mensagem_sucesso,
        mensagem_erro=mensagem_erro
    )


@app.route("/estoque/categorias/alternar/<int:categoria_id>", methods=["POST"])
@estoque_obrigatorio
def estoque_alternar_categoria(categoria_id):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE categorias_estoque
        SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (categoria_id,))

    conn.commit()
    conn.close()
    return redirect(url_for("estoque_categorias"))


@app.route("/compras/lancar-estoque/<int:item_id>", methods=["POST"])
@estoque_obrigatorio
def compras_lancar_estoque(item_id):
    item = buscar_item_compra_para_estoque(item_id)

    if not item:
        return redirect(url_for("compras_pedidos"))

    if item["status"] != "Comprado":
        return render_template("acesso_negado.html", mensagem="Só é possível lançar no estoque itens já comprados.")

    if item["estoque_movimentacao_id"]:
        return redirect(url_for("compras_detalhes", pedido_id=item["pedido_id"]))

    acao_produto = request.form.get("acao_produto", "existente")
    produto_id = request.form.get("produto_id", "").strip()
    quantidade = request.form.get("quantidade", "").strip() or item["quantidade"]
    data_movimentacao = request.form.get("data_movimentacao", "").strip()
    observacao = request.form.get("observacao", "").strip()

    try:
        if acao_produto == "novo":
            novo_nome = request.form.get("novo_produto_nome", "").strip()
            nova_categoria_id = request.form.get("nova_categoria_id", "").strip()
            nova_unidade = request.form.get("nova_unidade", "").strip() or item["unidade"]
            estoque_minimo = request.form.get("estoque_minimo", "0").strip()

            produto_id = criar_produto_estoque(novo_nome, nova_categoria_id, nova_unidade, estoque_minimo, "Criado a partir de item de compra.")
        else:
            produto_id = int(produto_id)

        if not produto_id:
            raise ValueError("Selecione ou crie um produto do estoque.")

        movimentacao_id = registrar_movimentacao_estoque(
            produto_id=produto_id,
            tipo="Entrada",
            quantidade=quantidade,
            data_movimentacao=data_movimentacao,
            origem="Compra",
            origem_id=item["pedido_id"],
            item_compra_id=item_id,
            fornecedor_id=item["fornecedor_id"],
            observacao=observacao or f"Entrada gerada pelo pedido de compra #{item['pedido_id']}. Descrição solicitada: {item['descricao']}"
        )

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE itens_compra
            SET produto_estoque_id = ?, estoque_movimentacao_id = ?
            WHERE id = ?
        """, (produto_id, movimentacao_id, item_id))
        conn.commit()
        conn.close()

    except Exception as e:
        return render_template("acesso_negado.html", mensagem=f"Não foi possível lançar no estoque: {str(e)}")

    return redirect(url_for("compras_detalhes", pedido_id=item["pedido_id"]))



# ============================================================
# ROTAS DE NOTAS DE ENTRADA / EVENTOS DE COMPRA - V3.2
# ============================================================

@app.route("/compras/notas", methods=["GET"])
@pode_ver_compras_obrigatorio
def notas_entrada_lista():
    status = request.args.get("status", "").strip()
    busca = request.args.get("busca", "").strip()
    fornecedor_id = request.args.get("fornecedor_id", "").strip()
    return render_template(
        "notas_entrada.html",
        notas=carregar_notas_entrada(status, busca, fornecedor_id),
        fornecedores=carregar_fornecedores(False),
        status=status, busca=busca, fornecedor_id=fornecedor_id,
    )


@app.route("/compras/notas/nova", methods=["GET", "POST"])
@admin_ou_compras_obrigatorio
def nota_entrada_nova():
    erro = None
    if request.method == "POST":
        try:
            nota_id = criar_nota_entrada(
                request.form.get("fornecedor_id"), request.form.get("numero"), request.form.get("serie"),
                request.form.get("chave_acesso"), request.form.get("data_emissao"), request.form.get("data_entrada"),
                request.form.get("pedido_compra_id"), request.form.get("desconto_geral"), request.form.get("frete"),
                request.form.get("outras_despesas"), request.form.get("observacao"),
            )
            return redirect(url_for("nota_entrada_detalhes", nota_id=nota_id))
        except Exception as exc:
            erro = str(exc)
    return render_template(
        "nota_entrada_form.html", fornecedores=carregar_fornecedores(),
        pedidos=carregar_pedidos_compra_para_nota(), erro=erro,
        hoje=datetime.now().strftime("%Y-%m-%d"),
    )


@app.route("/compras/notas/<int:nota_id>", methods=["GET"])
@pode_ver_compras_obrigatorio
def nota_entrada_detalhes(nota_id):
    nota = buscar_nota_entrada(nota_id)
    if not nota:
        return render_template("acesso_negado.html", mensagem="Documento de entrada não encontrado.")
    return render_template(
        "nota_entrada_detalhes.html", nota=nota, itens=carregar_itens_nota_entrada(nota_id),
        produtos=carregar_produtos_estoque(), centros=carregar_centros_custo(),
        itens_compra=carregar_itens_compra_disponiveis_nota(nota["pedido_compra_id"]),
        naturezas=NATUREZAS_ITEM_NOTA, unidades=UNIDADES_COMPRA,
        embalagens=carregar_todas_embalagens_ativas(),
    )


@app.route("/compras/notas/<int:nota_id>/item", methods=["POST"])
@admin_ou_compras_obrigatorio
def nota_entrada_adicionar_item(nota_id):
    try:
        adicionar_item_nota_entrada(
            nota_id, request.form.get("descricao"), request.form.get("natureza"),
            request.form.get("produto_id"), request.form.get("centro_custo_id"),
            request.form.get("item_compra_id"), request.form.get("quantidade_faturada"),
            request.form.get("quantidade_bonificada"), request.form.get("unidade"),
            request.form.get("embalagem_id"), request.form.get("valor_unitario"), request.form.get("desconto_item"),
            request.form.get("observacao"),
        )
    except Exception as exc:
        return redirect(url_for("nota_entrada_detalhes", nota_id=nota_id, erro=str(exc)))
    return redirect(url_for("nota_entrada_detalhes", nota_id=nota_id))


@app.route("/compras/notas/item/<int:item_id>/excluir", methods=["POST"])
@admin_ou_compras_obrigatorio
def nota_entrada_excluir_item(item_id):
    try:
        nota_id = excluir_item_nota_entrada(item_id)
    except Exception as exc:
        return render_template("acesso_negado.html", mensagem=str(exc))
    return redirect(url_for("nota_entrada_detalhes", nota_id=nota_id))


@app.route("/compras/notas/<int:nota_id>/lancar", methods=["POST"])
@admin_ou_compras_obrigatorio
def nota_entrada_lancar(nota_id):
    try:
        lancar_nota_entrada(nota_id)
    except Exception as exc:
        return redirect(url_for("nota_entrada_detalhes", nota_id=nota_id, erro=str(exc)))
    return redirect(url_for("nota_entrada_detalhes", nota_id=nota_id, sucesso="Documento lançado com sucesso."))


@app.route("/compras/notas/<int:nota_id>/excluir", methods=["POST"])
@admin_ou_compras_obrigatorio
def nota_entrada_excluir(nota_id):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("SELECT status FROM notas_entrada WHERE id = ?", (nota_id,))
    nota = cursor.fetchone()
    if not nota or nota["status"] != "Rascunho":
        conn.close(); return render_template("acesso_negado.html", mensagem="Somente rascunhos podem ser excluídos.")
    cursor.execute("UPDATE itens_compra SET nota_entrada_item_id = NULL WHERE nota_entrada_item_id IN (SELECT id FROM itens_nota_entrada WHERE nota_id = ?)", (nota_id,))
    cursor.execute("DELETE FROM itens_nota_entrada WHERE nota_id = ?", (nota_id,))
    cursor.execute("DELETE FROM notas_entrada WHERE id = ?", (nota_id,))
    conn.commit(); conn.close()
    return redirect(url_for("notas_entrada_lista"))


# ============================================================
# ROTAS DE LOTES E VALIDADES - V2.6
# ============================================================

@app.route("/validade", methods=["GET"])
@producao_obrigatorio
def validade_inicio():
    busca = request.args.get("busca", "").strip()
    status = request.args.get("status", "").strip()
    produto_id = request.args.get("produto_id", "").strip()
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()
    return render_template(
        "validade_home.html",
        lotes=carregar_lotes_validade(busca, status, produto_id, data_inicio, data_fim),
        produtos=carregar_produtos_estoque(apenas_ativos=False),
        resumo=contar_lotes_validade_resumo(),
        status_opcoes=STATUS_VALIDADE_FILTROS,
        busca=busca,
        status=status,
        produto_id=produto_id,
        data_inicio=data_inicio,
        data_fim=data_fim
    )


@app.route("/validade/novo", methods=["GET", "POST"])
@producao_obrigatorio
def validade_novo():
    mensagem_erro = None
    if request.method == "POST":
        try:
            lote_id = criar_lote_validade(
                produto_id=int(request.form.get("produto_id")),
                data_producao=request.form.get("data_producao", ""),
                data_validade=request.form.get("data_validade", ""),
                quantidade=request.form.get("quantidade", "0"),
                local_tipo=request.form.get("local_tipo", "Estoque Central"),
                celula_id=request.form.get("celula_id"),
                local_descricao=request.form.get("local_descricao", ""),
                origem="Cadastro Manual",
                observacao=request.form.get("observacao", ""),
                gerar_entrada_estoque=request.form.get("gerar_entrada_estoque") == "1"
            )
            return redirect(url_for("validade_detalhes", lote_id=lote_id, mensagem="Lote criado com sucesso."))
        except Exception as e:
            mensagem_erro = str(e)

    return render_template(
        "validade_novo.html",
        produtos=carregar_produtos_estoque(apenas_ativos=True),
        celulas=carregar_celulas_producao(apenas_ativas=True),
        locais=LOCAIS_VALIDADE,
        hoje=datetime.now().strftime("%Y-%m-%d"),
        mensagem_erro=mensagem_erro
    )


@app.route("/validade/<int:lote_id>", methods=["GET", "POST"])
@producao_obrigatorio
def validade_detalhes(lote_id):
    mensagem = request.args.get("mensagem", "").strip()
    erro = request.args.get("erro", "").strip()
    lote = buscar_lote_validade(lote_id)
    if not lote:
        return render_template("acesso_negado.html", mensagem="Lote de validade não encontrado.")

    if request.method == "POST":
        try:
            atualizar_lote_validade(
                lote_id,
                request.form.get("data_producao", ""),
                request.form.get("data_validade", ""),
                request.form.get("local_tipo", "Estoque Central"),
                request.form.get("celula_id"),
                request.form.get("local_descricao", ""),
                request.form.get("observacao", "")
            )
            return redirect(url_for("validade_detalhes", lote_id=lote_id, mensagem="Lote atualizado com sucesso."))
        except Exception as e:
            erro = str(e)
            lote = buscar_lote_validade(lote_id)

    return render_template(
        "validade_detalhes.html",
        lote=lote,
        movimentacoes=carregar_movimentacoes_validade(lote_id),
        celulas=carregar_celulas_producao(apenas_ativas=False),
        locais=LOCAIS_VALIDADE,
        tipos_movimentacao=TIPOS_MOVIMENTACAO_VALIDADE,
        motivos_perda=carregar_motivos_perda("Validade"),
        destinos=carregar_destinos_lote(lote_id),
        mensagem=mensagem,
        erro=erro
    )


@app.route("/validade/<int:lote_id>/movimentar", methods=["POST"])
@producao_obrigatorio
def validade_movimentar(lote_id):
    try:
        movimentar_lote_validade(
            lote_id,
            request.form.get("tipo", ""),
            request.form.get("quantidade", "0"),
            request.form.get("observacao", ""),
            request.form.get("motivo_perda_id")
        )
        return redirect(url_for("validade_detalhes", lote_id=lote_id, mensagem="Movimentação registrada com sucesso."))
    except Exception as e:
        return redirect(url_for("validade_detalhes", lote_id=lote_id, erro=str(e)))


@app.route("/validade/<int:lote_id>/cancelar", methods=["POST"])
@producao_obrigatorio
def validade_cancelar(lote_id):
    try:
        cancelar_lote_validade(lote_id, request.form.get("observacao_cancelamento", ""))
        return redirect(url_for("validade_detalhes", lote_id=lote_id, mensagem="Lote cancelado."))
    except Exception as e:
        return redirect(url_for("validade_detalhes", lote_id=lote_id, erro=str(e)))


@app.route("/producao-realizada/<int:producao_realizada_id>/gerar-lotes", methods=["POST"])
@producao_obrigatorio
def producao_realizada_gerar_lotes(producao_realizada_id):
    try:
        gerados = gerar_lotes_producao_confirmada(producao_realizada_id)
        mensagem = f"Processamento concluído. {gerados} lote(s) com controle de validade disponível(is)."
        return redirect(url_for("producao_realizada_detalhes", producao_realizada_id=producao_realizada_id, mensagem=mensagem))
    except Exception as e:
        return render_template("acesso_negado.html", mensagem=f"Não foi possível gerar os lotes: {str(e)}")





# ============================================================
# ROTAS DE RASTREABILIDADE E PERDAS - V3.4
# ============================================================

@app.route("/rastreabilidade", methods=["GET"])
@producao_obrigatorio
def rastreabilidade_inicio():
    busca = request.args.get("busca", "").strip()
    status = request.args.get("status", "").strip()
    produto_id = request.args.get("produto_id", "").strip()
    loja_id = request.args.get("loja_id", "").strip()
    return render_template(
        "rastreabilidade_home.html",
        lotes=carregar_lotes_rastreabilidade(busca, status, produto_id, loja_id),
        produtos=carregar_produtos_estoque(apenas_ativos=False),
        lojas=carregar_lojas(apenas_ativas=False),
        resumo=contar_lotes_validade_resumo(),
        status_opcoes=STATUS_VALIDADE_FILTROS,
        busca=busca, status=status, produto_id=produto_id, loja_id=loja_id,
    )


@app.route("/rastreabilidade/lote/<int:lote_id>", methods=["GET"])
@producao_obrigatorio
def rastreabilidade_lote(lote_id):
    lote = buscar_lote_validade(lote_id)
    if not lote:
        return render_template("acesso_negado.html", mensagem="Lote não encontrado."), 404
    contexto = carregar_contexto_producao_lote(lote)
    return render_template(
        "rastreabilidade_lote.html",
        lote=lote,
        contexto=contexto,
        destinos=carregar_destinos_lote(lote_id),
        eventos=carregar_linha_tempo_lote(lote_id),
        movimentacoes=carregar_movimentacoes_validade(lote_id),
    )


@app.route("/perdas", methods=["GET"])
@producao_obrigatorio
def perdas_relatorio():
    filtros = {
        "data_inicio": request.args.get("data_inicio", "").strip(),
        "data_fim": request.args.get("data_fim", "").strip(),
        "origem_tipo": request.args.get("origem_tipo", "").strip(),
        "produto_id": request.args.get("produto_id", "").strip(),
        "motivo_id": request.args.get("motivo_id", "").strip(),
        "celula_id": request.args.get("celula_id", "").strip(),
    }
    relatorio = carregar_relatorio_perdas(**filtros)
    return render_template(
        "perdas_relatorio.html", relatorio=relatorio, filtros=filtros,
        produtos=carregar_produtos_estoque(apenas_ativos=False),
        motivos=carregar_motivos_perda(apenas_ativos=False),
        celulas=carregar_celulas_producao(apenas_ativas=False),
        origens=ORIGENS_PERDA_FILTROS,
    )


@app.route("/motivos-perda", methods=["GET", "POST"])
@admin_obrigatorio
def motivos_perda_gerenciar():
    mensagem = request.args.get("mensagem", "")
    erro = request.args.get("erro", "")
    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        aplicacao = request.form.get("aplicacao", "Ambos")
        observacao = request.form.get("observacao", "").strip()
        exige = 1 if request.form.get("exige_observacao") == "1" else 0
        if not nome or aplicacao not in APLICACOES_MOTIVO_PERDA:
            erro = "Informe nome e aplicação válidos."
        else:
            try:
                conn = conectar(); cursor = conn.cursor()
                cursor.execute("INSERT INTO motivos_perda (nome, aplicacao, exige_observacao, ativo, observacao) VALUES (?, ?, ?, 1, ?)", (nome, aplicacao, exige, observacao))
                conn.commit(); conn.close()
                return redirect(url_for("motivos_perda_gerenciar", mensagem="Motivo criado com sucesso."))
            except sqlite3.IntegrityError:
                erro = "Já existe um motivo com esse nome."
    return render_template("motivos_perda.html", motivos=carregar_motivos_perda(apenas_ativos=False), aplicacoes=APLICACOES_MOTIVO_PERDA, mensagem=mensagem, erro=erro)


@app.route("/motivos-perda/<int:motivo_id>/alternar", methods=["POST"])
@admin_obrigatorio
def motivo_perda_alternar(motivo_id):
    conn = conectar(); cursor = conn.cursor()
    cursor.execute("UPDATE motivos_perda SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END WHERE id = ?", (motivo_id,))
    conn.commit(); conn.close()
    return redirect(url_for("motivos_perda_gerenciar", mensagem="Status atualizado."))


# ============================================================
# ROTAS DE FECHAMENTO / DEMANDA DAS LOJAS - V2.9
# ============================================================

@app.route("/pedidos-lojas/fechamentos", methods=["GET"])
@producao_obrigatorio
def rodadas_pedidos_lista():
    status = request.args.get("status", "").strip()
    return render_template(
        "rodadas_pedidos.html",
        rodadas=carregar_rodadas_pedidos(status),
        status=status,
        status_opcoes=STATUS_RODADA_PEDIDOS,
    )


@app.route("/pedidos-lojas/fechamentos/<int:rodada_id>", methods=["GET"])
@producao_obrigatorio
def rodada_pedidos_detalhes(rodada_id):
    rodada = buscar_rodada_pedidos(rodada_id)
    if not rodada:
        return render_template("acesso_negado.html", mensagem="Fechamento de pedidos não encontrado."), 404
    return render_template(
        "rodada_pedidos_detalhes.html",
        rodada=rodada,
        pedidos=carregar_pedidos_rodada(rodada_id),
        demanda=carregar_demanda_rodada(rodada_id),
        mistos=carregar_mistos_rodada(rodada_id),
        mensagem=request.args.get("mensagem", ""),
        erro=request.args.get("erro", ""),
        cenarios=["Chuva", "Normal", "Verão", "Baixa"],
        dias_semana=DIAS_SEMANA,
    )


@app.route("/pedidos-lojas/fechamentos/<int:rodada_id>/fechar", methods=["POST"])
@producao_obrigatorio
def rodada_pedidos_fechar(rodada_id):
    try:
        fechar_rodada_pedidos(
            rodada_id,
            request.form.get("cenario", "Baixa"),
            request.form.get("dia_semana", "Segunda-feira"),
            request.form.get("observacao", ""),
        )
        return redirect(url_for("rodada_pedidos_detalhes", rodada_id=rodada_id, mensagem="Pedidos fechados e demanda congelada para produção."))
    except Exception as e:
        return redirect(url_for("rodada_pedidos_detalhes", rodada_id=rodada_id, erro=str(e)))


@app.route("/pedidos-lojas/fechamentos/<int:rodada_id>/reabrir", methods=["POST"])
@admin_obrigatorio
def rodada_pedidos_reabrir(rodada_id):
    try:
        reabrir_rodada_pedidos(rodada_id)
        return redirect(url_for("rodada_pedidos_detalhes", rodada_id=rodada_id, mensagem="Rodada reaberta. As lojas podem cancelar e reenviar pedidos novamente."))
    except Exception as e:
        return redirect(url_for("rodada_pedidos_detalhes", rodada_id=rodada_id, erro=str(e)))


@app.route("/romaneios/<int:rodada_id>", methods=["GET"])
@login_obrigatorio
def romaneio_rodada(rodada_id):
    rodada, grupos = gerar_romaneio_rodada(rodada_id)
    if not rodada:
        return render_template("acesso_negado.html", mensagem="Rodada não encontrada."), 404
    return render_template("romaneio_rodada.html", rodada=rodada, grupos=grupos)


# ============================================================
# ROTAS DE PEDIDOS DAS LOJAS / EXPEDIÇÃO - V2.8
# ============================================================

@app.route("/loja", methods=["GET"])
@loja_obrigatorio
def loja_painel():
    loja_id = session.get("loja_id")
    if session.get("perfil") == "admin":
        loja_id = request.args.get("loja_id", "") or None
    loja = buscar_loja(loja_id) if loja_id else None
    return render_template(
        "loja_painel.html",
        loja=loja,
        pedidos=carregar_pedidos_loja(loja_id=loja_id, limite=5) if loja_id else [],
        total_abertos=contar_pedidos_loja_pendentes(loja_id),
    )


@app.route("/loja/pedidos/novo", methods=["GET", "POST"])
@loja_obrigatorio
def loja_pedido_novo():
    loja_id = session.get("loja_id")
    if session.get("perfil") == "admin":
        loja_id = request.args.get("loja_id", "") or request.form.get("loja_id", "")
    loja = buscar_loja(loja_id) if loja_id else None
    if not loja:
        return render_template("acesso_negado.html", mensagem="Selecione ou vincule uma loja antes de criar o pedido.")

    erro = None
    if request.method == "POST":
        try:
            pedido_id = criar_pedido_loja(
                int(loja_id),
                request.form.get("data_entrega_desejada", ""),
                request.form.get("observacao", ""),
                request.form.getlist("produto_id"),
                request.form.getlist("embalagem_id"),
                request.form.getlist("quantidade"),
                request.form.get("misto_sabor_1", ""),
                request.form.get("misto_sabor_2", ""),
                request.form.get("misto_quantidade", "0"),
            )
            return redirect(url_for("loja_pedido_detalhes", pedido_id=pedido_id, mensagem="Pedido enviado com sucesso."))
        except Exception as e:
            erro = str(e)

    categorias, sabores_misto = carregar_catalogo_loja()
    return render_template(
        "loja_pedido_novo.html",
        loja=loja,
        categorias=categorias,
        sabores_misto=sabores_misto,
        erro=erro,
        hoje=datetime.now().strftime("%Y-%m-%d"),
    )


@app.route("/loja/pedidos", methods=["GET"])
@loja_obrigatorio
def loja_pedidos():
    loja_id = session.get("loja_id")
    if session.get("perfil") == "admin":
        loja_id = request.args.get("loja_id", "") or None
    status = request.args.get("status", "").strip()
    return render_template(
        "loja_pedidos.html",
        loja=buscar_loja(loja_id) if loja_id else None,
        pedidos=carregar_pedidos_loja(loja_id=loja_id, status=status),
        status_filtro=status,
        status_opcoes=STATUS_PEDIDO_LOJA,
        lojas=carregar_lojas() if session.get("perfil") == "admin" else [],
    )


@app.route("/loja/pedidos/<int:pedido_id>", methods=["GET"])
@login_obrigatorio
def loja_pedido_detalhes(pedido_id):
    loja_id = session.get("loja_id") if session.get("perfil") == "loja" else None
    pedido = buscar_pedido_loja(pedido_id, loja_id=loja_id)
    if not pedido:
        return render_template("acesso_negado.html", mensagem="Pedido não encontrado ou sem permissão de acesso.")
    return render_template(
        "loja_pedido_detalhes.html",
        pedido=pedido,
        itens=carregar_itens_pedido_loja(pedido_id),
        mensagem=request.args.get("mensagem", ""),
        erro=request.args.get("erro", ""),
        pode_cancelar=pedido["status"] == "Recebido" and (not pedido["rodada_id"] or (buscar_rodada_pedidos(pedido["rodada_id"])["status"] == "Aberta")) and (session.get("perfil") in ["loja", "admin"]),
    )


@app.route("/loja/pedidos/<int:pedido_id>/cancelar", methods=["POST"])
@loja_obrigatorio
def loja_pedido_cancelar(pedido_id):
    loja_id = session.get("loja_id") if session.get("perfil") == "loja" else None
    pedido = buscar_pedido_loja(pedido_id, loja_id=loja_id)
    if not pedido:
        return render_template("acesso_negado.html", mensagem="Pedido não encontrado.")
    rodada = buscar_rodada_pedidos(pedido["rodada_id"]) if pedido["rodada_id"] else None
    if pedido["status"] != "Recebido" or (rodada and rodada["status"] != "Aberta"):
        return redirect(url_for("loja_pedido_detalhes", pedido_id=pedido_id, erro="O pedido só pode ser cancelado antes do fechamento dos pedidos para produção."))
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE pedidos_loja
        SET status = 'Cancelado', data_cancelamento = ?, cancelado_por = ?, motivo_cancelamento = ?
        WHERE id = ?
    """, (
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        session.get("usuario", "loja"),
        request.form.get("motivo", "Cancelado pela loja").strip(),
        pedido_id,
    ))
    conn.commit()
    conn.close()
    return redirect(url_for("loja_pedido_detalhes", pedido_id=pedido_id, mensagem="Pedido cancelado."))


@app.route("/expedicao", methods=["GET"])
@expedicao_obrigatorio
def expedicao_home():
    origem_id = request.args.get("origem_id", "").strip()
    status = request.args.get("status", "").strip()
    return render_template(
        "expedicao_home.html",
        pedidos=carregar_pedidos_expedicao(origem_id=origem_id or None, status=status),
        origens=carregar_origens_expedicao(apenas_ativas=False),
        origem_id=origem_id,
        status_filtro=status,
        status_opcoes=STATUS_PEDIDO_LOJA,
        total_pendentes=contar_pedidos_loja_pendentes(),
        resumo_romaneios=contar_romaneios_status(),
    )


@app.route("/expedicao/pedidos/<int:pedido_id>", methods=["GET"])
@expedicao_obrigatorio
def expedicao_pedido_detalhes(pedido_id):
    pedido = buscar_pedido_loja(pedido_id)
    if not pedido:
        return render_template("acesso_negado.html", mensagem="Pedido não encontrado.")
    return render_template(
        "expedicao_pedido_detalhes.html",
        pedido=pedido,
        itens=carregar_itens_pedido_loja(pedido_id),
        romaneio=buscar_romaneio_por_pedido(pedido_id),
        mensagem=request.args.get("mensagem", ""),
        erro=request.args.get("erro", ""),
    )


@app.route("/expedicao/itens/<int:item_id>/separar", methods=["POST"])
@expedicao_obrigatorio
def expedicao_item_separar(item_id):
    try:
        pedido_id = registrar_separacao_item_loja(
            item_id,
            request.form.get("quantidade_separada", "0"),
            indisponivel=request.form.get("indisponivel") == "1",
            observacao=request.form.get("observacao_expedicao", ""),
        )
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, mensagem="Separação atualizada."))
    except Exception as e:
        pedido_id = request.form.get("pedido_id", "")
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, erro=str(e)))


@app.route("/expedicao/pedidos/<int:pedido_id>/separacao/salvar", methods=["POST"])
@expedicao_obrigatorio
def expedicao_pedido_separacao_salvar(pedido_id):
    try:
        registrar_separacao_pedido_loja(pedido_id, request.form)
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, mensagem="Separação do pedido salva em lote."))
    except Exception as e:
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, erro=str(e)))


@app.route("/expedicao/pedidos/<int:pedido_id>/finalizar", methods=["POST"])
@expedicao_obrigatorio
def expedicao_pedido_finalizar(pedido_id):
    # Compatibilidade com a v2.8: o antigo botão de finalização agora prepara
    # a conferência e o romaneio, sem dar baixa prematuramente no estoque.
    try:
        romaneio_id = preparar_romaneio_pedido(pedido_id)
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, mensagem="Romaneio preparado. Faça a conferência antes da saída."))
    except Exception as e:
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, erro=str(e)))


@app.route("/expedicao/pedidos/<int:pedido_id>/romaneio/preparar", methods=["POST"])
@expedicao_obrigatorio
def expedicao_preparar_romaneio(pedido_id):
    try:
        romaneio_id = preparar_romaneio_pedido(pedido_id)
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, mensagem="Romaneio preparado para conferência."))
    except Exception as e:
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, erro=str(e)))


@app.route("/expedicao/romaneios/<int:romaneio_id>", methods=["GET"])
@expedicao_obrigatorio
def romaneio_detalhes(romaneio_id):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio:
        return render_template("acesso_negado.html", mensagem="Romaneio não encontrado."), 404
    return render_template(
        "romaneio_detalhes.html",
        romaneio=romaneio,
        itens=carregar_itens_romaneio(romaneio_id),
        motoristas=carregar_motoristas(apenas_ativos=True),
        veiculos=carregar_veiculos(apenas_ativos=True),
        custos_romaneio=carregar_custos_romaneio(romaneio_id) if session.get("perfil") == "admin" else [],
        mensagem=request.args.get("mensagem", ""),
        erro=request.args.get("erro", ""),
    )


@app.route("/expedicao/romaneios/<int:romaneio_id>/salvar-conferencia", methods=["POST"])
@expedicao_obrigatorio
def romaneio_salvar_conferencia(romaneio_id):
    try:
        salvar_conferencia_romaneio(romaneio_id, request.form)
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, mensagem="Conferência salva. Revise e confirme."))
    except Exception as e:
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, erro=str(e)))


@app.route("/expedicao/romaneios/<int:romaneio_id>/confirmar-conferencia", methods=["POST"])
@expedicao_obrigatorio
def romaneio_confirmar_conferencia(romaneio_id):
    try:
        confirmar_conferencia_romaneio(romaneio_id)
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, mensagem="Conferência confirmada e lotes FEFO reservados para a saída."))
    except Exception as e:
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, erro=str(e)))


@app.route("/expedicao/romaneios/<int:romaneio_id>/saida", methods=["POST"])
@expedicao_obrigatorio
def romaneio_registrar_saida(romaneio_id):
    try:
        registrar_saida_romaneio(romaneio_id)
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, mensagem="Saída registrada. O pedido está em rota."))
    except Exception as e:
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, erro=str(e)))


@app.route("/expedicao/romaneios/<int:romaneio_id>/retorno", methods=["POST"])
@expedicao_obrigatorio
def romaneio_registrar_retorno(romaneio_id):
    try:
        registrar_retorno_romaneio(
            romaneio_id,
            request.form.get("recebido_loja_por", ""),
            request.form.get("observacao_retorno", ""),
        )
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, mensagem="Via retornada e pedido finalizado."))
    except Exception as e:
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, erro=str(e)))


@app.route("/expedicao/romaneios/<int:romaneio_id>/cancelar-preparacao", methods=["POST"])
@expedicao_obrigatorio
def romaneio_cancelar_preparacao(romaneio_id):
    romaneio = buscar_romaneio(romaneio_id)
    pedido_id = romaneio["pedido_id"] if romaneio else None
    try:
        cancelar_preparacao_romaneio(romaneio_id)
        return redirect(url_for("expedicao_pedido_detalhes", pedido_id=pedido_id, mensagem="Preparação do romaneio cancelada. A separação pode ser alterada novamente."))
    except Exception as e:
        return redirect(url_for("romaneio_detalhes", romaneio_id=romaneio_id, erro=str(e)))


@app.route("/expedicao/romaneios/<int:romaneio_id>/imprimir", methods=["GET"])
@expedicao_obrigatorio
def romaneio_imprimir(romaneio_id):
    romaneio = buscar_romaneio(romaneio_id)
    if not romaneio:
        return render_template("acesso_negado.html", mensagem="Romaneio não encontrado."), 404
    itens = carregar_itens_romaneio(romaneio_id, incluir_sugestoes=False)
    return render_template("romaneio_impressao.html", romaneio=romaneio, itens=itens)


@app.route("/lojas", methods=["GET", "POST"])
@admin_obrigatorio
def lojas_admin():
    mensagem = request.args.get("mensagem", "")
    erro = request.args.get("erro", "")
    if request.method == "POST":
        codigo = request.form.get("codigo", "").strip().upper()
        nome = request.form.get("nome", "").strip()
        observacao = request.form.get("observacao", "").strip()
        if not codigo or not nome:
            erro = "Informe código e nome da loja."
        else:
            conn = conectar()
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    INSERT INTO lojas (codigo, nome, ativo, observacao, data_cadastro)
                    VALUES (?, ?, 1, ?, ?)
                """, (codigo, nome, observacao, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
                conn.commit()
                mensagem = "Loja cadastrada com sucesso."
            except sqlite3.IntegrityError:
                erro = "Já existe uma loja com esse código ou nome."
            finally:
                conn.close()
    return render_template("lojas.html", lojas=carregar_lojas(apenas_ativas=False), mensagem=mensagem, erro=erro)


@app.route("/lojas/<int:loja_id>/alternar", methods=["POST"])
@admin_obrigatorio
def loja_alternar(loja_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("UPDATE lojas SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END WHERE id = ?", (loja_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("lojas_admin"))


@app.route("/expedicao/origens", methods=["GET", "POST"])
@admin_obrigatorio
def origens_expedicao_admin():
    mensagem = ""
    erro = ""
    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        observacao = request.form.get("observacao", "").strip()
        if not nome:
            erro = "Informe o nome da origem."
        else:
            conn = conectar()
            cursor = conn.cursor()
            try:
                cursor.execute("INSERT INTO origens_expedicao (nome, ativo, observacao) VALUES (?, 1, ?)", (nome, observacao))
                conn.commit()
                mensagem = "Origem cadastrada."
            except sqlite3.IntegrityError:
                erro = "Essa origem já existe."
            finally:
                conn.close()
    return render_template("origens_expedicao.html", origens=carregar_origens_expedicao(apenas_ativas=False), mensagem=mensagem, erro=erro)


@app.route("/expedicao/origens/<int:origem_id>/alternar", methods=["POST"])
@admin_obrigatorio
def origem_expedicao_alternar(origem_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("UPDATE origens_expedicao SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END WHERE id = ?", (origem_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("origens_expedicao_admin"))


@app.route("/estoque/produtos/embalagens/loja/<int:embalagem_id>", methods=["POST"])
@estoque_obrigatorio
def estoque_alternar_embalagem_loja(embalagem_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT produto_id FROM produto_embalagens WHERE id = ?", (embalagem_id,))
    embalagem = cursor.fetchone()
    if not embalagem:
        conn.close()
        return redirect(url_for("estoque_produtos"))
    cursor.execute("""
        UPDATE produto_embalagens
        SET disponivel_loja = CASE WHEN disponivel_loja = 1 THEN 0 ELSE 1 END
        WHERE id = ?
    """, (embalagem_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("estoque_editar_produto", produto_id=embalagem["produto_id"]))

# ============================================================
# ROTAS DE MOTORISTAS E VEÍCULOS - V3.0.1
# ============================================================

@app.route("/expedicao/motoristas", methods=["GET", "POST"])
@expedicao_obrigatorio
def expedicao_motoristas():
    mensagem = request.args.get("mensagem", "")
    erro = request.args.get("erro", "")
    editar_id = request.args.get("editar", "").strip()
    motorista_edicao = buscar_motorista(editar_id) if editar_id else None
    if request.method == "POST":
        registro_id = request.form.get("id", "").strip()
        nome = request.form.get("nome", "").strip()
        if not nome:
            erro = "Informe o nome do motorista."
        else:
            conn = conectar()
            cursor = conn.cursor()
            try:
                if registro_id:
                    cursor.execute("""
                        UPDATE motoristas
                        SET nome = ?, telefone = ?, categoria_cnh = ?, validade_cnh = ?, observacao = ?
                        WHERE id = ?
                    """, (nome, request.form.get("telefone", "").strip(), request.form.get("categoria_cnh", "").strip(), request.form.get("validade_cnh", "").strip() or None, request.form.get("observacao", "").strip(), registro_id))
                    mensagem = "Motorista atualizado."
                else:
                    cursor.execute("""
                        INSERT INTO motoristas
                        (nome, telefone, categoria_cnh, validade_cnh, ativo, observacao, data_cadastro)
                        VALUES (?, ?, ?, ?, 1, ?, ?)
                    """, (nome, request.form.get("telefone", "").strip(), request.form.get("categoria_cnh", "").strip(), request.form.get("validade_cnh", "").strip() or None, request.form.get("observacao", "").strip(), _agora_texto()))
                    mensagem = "Motorista cadastrado."
                conn.commit()
                motorista_edicao = None
            except sqlite3.IntegrityError:
                conn.rollback()
                erro = "Já existe um motorista com esse nome."
            finally:
                conn.close()
    return render_template("motoristas.html", motoristas=carregar_motoristas(apenas_ativos=False), motorista_edicao=motorista_edicao, mensagem=mensagem, erro=erro)


@app.route("/expedicao/motoristas/<int:motorista_id>/alternar", methods=["POST"])
@expedicao_obrigatorio
def expedicao_motorista_alternar(motorista_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("UPDATE motoristas SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END WHERE id = ?", (motorista_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("expedicao_motoristas"))


@app.route("/expedicao/veiculos", methods=["GET", "POST"])
@expedicao_obrigatorio
def expedicao_veiculos():
    mensagem = request.args.get("mensagem", "")
    erro = request.args.get("erro", "")
    editar_id = request.args.get("editar", "").strip()
    veiculo_edicao = buscar_veiculo(editar_id) if editar_id else None
    if request.method == "POST":
        registro_id = request.form.get("id", "").strip()
        codigo = request.form.get("codigo", "").strip().upper()
        descricao = request.form.get("descricao", "").strip()
        placa = request.form.get("placa", "").strip().upper()
        if not codigo or not descricao or not placa:
            erro = "Informe código, descrição e placa."
        else:
            conn = conectar()
            cursor = conn.cursor()
            try:
                if registro_id:
                    cursor.execute("""
                        UPDATE veiculos
                        SET codigo = ?, descricao = ?, placa = ?, tipo = ?, capacidade = ?, observacao = ?
                        WHERE id = ?
                    """, (codigo, descricao, placa, request.form.get("tipo", "").strip(), request.form.get("capacidade", "").strip(), request.form.get("observacao", "").strip(), registro_id))
                    mensagem = "Veículo atualizado."
                else:
                    cursor.execute("""
                        INSERT INTO veiculos
                        (codigo, descricao, placa, tipo, capacidade, ativo, observacao, data_cadastro)
                        VALUES (?, ?, ?, ?, ?, 1, ?, ?)
                    """, (codigo, descricao, placa, request.form.get("tipo", "").strip(), request.form.get("capacidade", "").strip(), request.form.get("observacao", "").strip(), _agora_texto()))
                    mensagem = "Veículo cadastrado."
                conn.commit()
                veiculo_edicao = None
            except sqlite3.IntegrityError:
                conn.rollback()
                erro = "Código ou placa já cadastrado."
            finally:
                conn.close()
    return render_template("veiculos.html", veiculos=carregar_veiculos(apenas_ativos=False), veiculo_edicao=veiculo_edicao, mensagem=mensagem, erro=erro)


@app.route("/expedicao/veiculos/<int:veiculo_id>/alternar", methods=["POST"])
@expedicao_obrigatorio
def expedicao_veiculo_alternar(veiculo_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("UPDATE veiculos SET ativo = CASE WHEN ativo = 1 THEN 0 ELSE 1 END WHERE id = ?", (veiculo_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("expedicao_veiculos"))


# ============================================================
# ROTAS DE CUSTOS GERENCIAIS - V3.1
# ============================================================

@app.route("/custos", methods=["GET"])
@admin_obrigatorio
def custos_home():
    hoje = datetime.now()
    data_inicio = request.args.get("data_inicio", hoje.replace(day=1).strftime("%Y-%m-%d")).strip()
    data_fim = request.args.get("data_fim", hoje.strftime("%Y-%m-%d")).strip()
    resumo = resumo_custos_gerais(data_inicio=data_inicio, data_fim=data_fim)
    return render_template(
        "custos_home.html",
        resumo=resumo,
        data_inicio=data_inicio,
        data_fim=data_fim,
    )


@app.route("/custos/lojas", methods=["GET"])
@admin_obrigatorio
def custos_lojas():
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()
    loja_id = request.args.get("loja_id", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()
    produto_id = request.args.get("produto_id", "").strip()
    relatorio = carregar_relatorio_custos_lojas(
        data_inicio=data_inicio,
        data_fim=data_fim,
        loja_id=loja_id,
        categoria_id=categoria_id,
        produto_id=produto_id,
    )
    return render_template(
        "custos_lojas.html",
        relatorio=relatorio,
        lojas=carregar_lojas(apenas_ativas=False),
        categorias=carregar_categorias_estoque(apenas_ativas=False),
        produtos=carregar_produtos_estoque(apenas_ativos=False),
        data_inicio=data_inicio,
        data_fim=data_fim,
        loja_id=loja_id,
        categoria_id=categoria_id,
        produto_id=produto_id,
    )


@app.route("/custos/centros", methods=["GET"])
@admin_obrigatorio
def custos_centros():
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()
    centro_custo_id = request.args.get("centro_custo_id", "").strip()
    resumo, detalhes, total_geral = carregar_relatorio_centros_custo(
        data_inicio=_data_html_para_br(data_inicio),
        data_fim=_data_html_para_br(data_fim),
        centro_custo_id=centro_custo_id,
    )
    return render_template(
        "custos_centros.html",
        resumo=resumo,
        detalhes=detalhes,
        total_geral=total_geral,
        centros=carregar_centros_custo(False),
        data_inicio=data_inicio,
        data_fim=data_fim,
        centro_custo_id=centro_custo_id,
    )


@app.route("/custos/produtos", methods=["GET"])
@admin_obrigatorio
def custos_produtos():
    busca = request.args.get("busca", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()
    somente_alertas = request.args.get("somente_alertas", "") == "1"
    produtos = carregar_custos_produtos(
        busca=busca,
        categoria_id=categoria_id,
        somente_alertas=somente_alertas,
    )
    return render_template(
        "custos_produtos.html",
        produtos=produtos,
        categorias=carregar_categorias_estoque(apenas_ativas=False),
        busca=busca,
        categoria_id=categoria_id,
        somente_alertas=somente_alertas,
    )


# ============================================================
# ROTAS DE ANÁLISE PREVISTO X REALIZADO - V2.5
# ============================================================

@app.route("/analises-producao", methods=["GET"])
@producao_obrigatorio
def analises_producao_lista():
    status = request.args.get("status", "").strip()
    return render_template(
        "analises_producao.html",
        producoes=carregar_producoes_para_analise(status=status),
        status=status,
        total_analises=contar_analises_disponiveis()
    )


@app.route("/analises-producao/<int:producao_realizada_id>", methods=["GET"])
@producao_obrigatorio
def analise_producao_detalhes(producao_realizada_id):
    analise = construir_analise_producao(producao_realizada_id)
    if not analise:
        return render_template(
            "acesso_negado.html",
            mensagem="Produção não encontrada para análise."
        )

    return render_template(
        "analise_producao_detalhes.html",
        analise=analise
    )


# ============================================================
# INICIAR SISTEMA
# ============================================================

if __name__ == "__main__":
    inicializar_banco()
    app.run(debug=True)
