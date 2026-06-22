import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flask import session

import app as sistema


class CorrecoesV351TestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tmp.name, "test_v351.db")
        self.old_db_name = sistema.DB_NAME
        sistema.DB_NAME = self.db_path
        sistema.app.config.update(TESTING=True, SECRET_KEY="test-v351")
        self.ctx = sistema.app.test_request_context("/")
        self.ctx.push()
        session["usuario"] = "tester"
        session["perfil"] = "admin"
        sistema.inicializar_banco()

    def tearDown(self):
        self.ctx.pop()
        sistema.DB_NAME = self.old_db_name
        self.tmp.cleanup()

    def query_one(self, sql, params=()):
        conn = sistema.conectar()
        cursor = conn.cursor()
        cursor.execute(sql, params)
        row = cursor.fetchone()
        conn.close()
        return row

    def query_all(self, sql, params=()):
        conn = sistema.conectar()
        cursor = conn.cursor()
        cursor.execute(sql, params)
        rows = cursor.fetchall()
        conn.close()
        return rows

    def execute(self, sql, params=()):
        conn = sistema.conectar()
        cursor = conn.cursor()
        cursor.execute(sql, params)
        conn.commit()
        conn.close()

    def count_table(self, table):
        return self.query_one(f"SELECT COUNT(*) AS total FROM {table}")["total"]

    def produto(self, nome):
        row = self.query_one("SELECT * FROM produtos_estoque WHERE nome = ?", (nome,))
        self.assertIsNotNone(row, f"Produto esperado nao encontrado: {nome}")
        return row

    def configurar_validade(self, nome, controla, dias=None):
        self.execute(
            "UPDATE produtos_estoque SET controla_validade = ?, dias_validade = ? WHERE nome = ?",
            (1 if controla else 0, dias, nome)
        )

    def criar_pre(self, resultado, total_tabuleiros=1, total_empadas=35):
        pre_id = sistema.criar_pre_producao_realizada(
            resultado,
            "Normal",
            "Segunda-feira",
            total_tabuleiros,
            total_empadas,
            origem="Teste v3.5.1"
        )
        self.assertIsNotNone(pre_id)
        return pre_id

    def movimentos_producao(self, producao_id):
        return self.query_all(
            "SELECT * FROM movimentacoes_estoque WHERE origem_id = ? ORDER BY id",
            (producao_id,)
        )

    def test_bug_001_misto_sem_tabuleiro_normal_registra_unidades(self):
        self.configurar_validade("Empada de Bacalhau", False)
        pre_id = self.criar_pre(
            [{
                "nome": "Bacalhau",
                "producao": 0,
                "empadas": 18,
                "misto_unidades": 18,
                "classe": "Prata",
            }],
            total_tabuleiros=1,
            total_empadas=18,
        )

        sistema.confirmar_producao_realizada_estoque(pre_id)

        movimentos = self.movimentos_producao(pre_id)
        self.assertEqual(len(movimentos), 1)
        self.assertAlmostEqual(float(movimentos[0]["quantidade"]), 18)
        self.assertAlmostEqual(float(movimentos[0]["quantidade_embalagem"]), 18)
        self.assertIsNone(movimentos[0]["embalagem_id"])

    def test_bug_001_normal_e_misto_do_mesmo_sabor_registra_total_em_unidades(self):
        self.configurar_validade("Empada de Bacalhau", False)
        pre_id = self.criar_pre(
            [{
                "nome": "Bacalhau",
                "producao": 1,
                "empadas": 53,
                "misto_unidades": 18,
                "classe": "Prata",
            }],
            total_tabuleiros=2,
            total_empadas=53,
        )

        sistema.confirmar_producao_realizada_estoque(pre_id)

        movimentos = self.movimentos_producao(pre_id)
        self.assertEqual(len(movimentos), 1)
        self.assertAlmostEqual(float(movimentos[0]["quantidade"]), 53)
        self.assertAlmostEqual(float(movimentos[0]["quantidade_embalagem"]), 53)
        self.assertIsNone(movimentos[0]["embalagem_id"])

    def test_bug_002_falha_no_segundo_item_desfaz_confirmacao_inteira(self):
        self.configurar_validade("Empada de Bacalhau", False)
        self.configurar_validade("Empada de Calabresa", False)
        pre_id = self.criar_pre(
            [
                {"nome": "Bacalhau", "producao": 1, "empadas": 35, "classe": "Prata"},
                {"nome": "Calabresa", "producao": 1, "empadas": 35, "classe": "Bronze"},
            ],
            total_tabuleiros=2,
            total_empadas=70,
        )

        original = sistema.registrar_movimentacao_estoque_cursor
        chamadas = {"total": 0}

        def falhar_no_segundo(cursor, *args, **kwargs):
            chamadas["total"] += 1
            if chamadas["total"] == 2:
                raise RuntimeError("falha simulada no segundo item")
            return original(cursor, *args, **kwargs)

        with patch.object(sistema, "registrar_movimentacao_estoque_cursor", side_effect=falhar_no_segundo):
            with self.assertRaisesRegex(RuntimeError, "falha simulada"):
                sistema.confirmar_producao_realizada_estoque(pre_id)

        self.assertEqual(chamadas["total"], 2)
        self.assertEqual(len(self.movimentos_producao(pre_id)), 0)
        producao = self.query_one("SELECT status FROM producoes_realizadas WHERE id = ?", (pre_id,))
        self.assertEqual(producao["status"], "Pendente")
        itens = self.query_all(
            "SELECT status, estoque_movimentacao_id FROM itens_producao_realizada WHERE producao_realizada_id = ?",
            (pre_id,)
        )
        self.assertTrue(all(item["status"] == "Pendente" for item in itens))
        self.assertTrue(all(item["estoque_movimentacao_id"] is None for item in itens))

    def test_bug_003_falha_durante_planejamento_desfaz_todos_os_registros(self):
        antes = {
            tabela: self.count_table(tabela)
            for tabela in [
                "planejamentos_producao",
                "itens_planejamento_producao",
                "separacoes_estoque",
                "producoes_realizadas",
            ]
        }
        resultado = [{"nome": "Bacalhau", "producao": 1, "empadas": 35, "classe": "Prata"}]

        with patch.object(sistema, "criar_pre_producao_realizada", side_effect=RuntimeError("falha simulada no planejamento")):
            with self.assertRaisesRegex(RuntimeError, "falha simulada"):
                sistema.criar_planejamento_producao(resultado, "Normal", "Segunda-feira", 1, 35)

        depois = {tabela: self.count_table(tabela) for tabela in antes}
        self.assertEqual(depois, antes)

    def test_bug_007_produto_com_validade_sem_dias_bloqueia_confirmacao(self):
        self.configurar_validade("Empada de Bacalhau", True, None)
        pre_id = self.criar_pre(
            [{"nome": "Bacalhau", "producao": 1, "empadas": 35, "classe": "Prata"}],
            total_tabuleiros=1,
            total_empadas=35,
        )

        with self.assertRaisesRegex(ValueError, "dias de validade"):
            sistema.confirmar_producao_realizada_estoque(pre_id)

        self.assertEqual(len(self.movimentos_producao(pre_id)), 0)
        self.assertEqual(
            self.count_table("lotes_validade"),
            0
        )
        producao = self.query_one("SELECT status FROM producoes_realizadas WHERE id = ?", (pre_id,))
        self.assertEqual(producao["status"], "Pendente")

    def test_fluxo_real_misto_planejamento_pre_producao_confirmacao_e_strings(self):
        self.configurar_validade("Empada de Bacalhau", False)
        self.configurar_validade("Empada de Calabresa", False)
        resultado = [
            {"nome": "Bacalhau", "producao": 0, "empadas": 18, "misto_unidades": 18, "classe": "Prata"},
            {"nome": "Calabresa", "producao": 1, "empadas": 52, "misto_unidades": 17, "classe": "Bronze"},
        ]

        planejamento_id = sistema.criar_planejamento_producao(
            resultado, "Normal", "Segunda-feira", 2, 70
        )
        self.assertIsNotNone(planejamento_id)

        planejamento = self.query_one(
            "SELECT origem, status, producao_realizada_id FROM planejamentos_producao WHERE id = ?",
            (planejamento_id,),
        )
        self.assertEqual(planejamento["origem"], "Meta de Produção")
        self.assertIn(planejamento["status"], ["Rascunho", "Com alertas"])

        itens_planejamento = self.query_all(
            "SELECT sabor, quantidade_unidades FROM itens_planejamento_producao WHERE planejamento_id = ? ORDER BY sabor",
            (planejamento_id,),
        )
        self.assertEqual({item["sabor"] for item in itens_planejamento}, {"Bacalhau", "Calabresa"})

        pre_id = planejamento["producao_realizada_id"]
        self.assertIsNotNone(pre_id)
        itens_pre = self.query_all(
            "SELECT sabor, quantidade_tabuleiros, quantidade_unidades FROM itens_producao_realizada WHERE producao_realizada_id = ? ORDER BY sabor",
            (pre_id,),
        )
        por_sabor = {item["sabor"]: item for item in itens_pre}
        self.assertAlmostEqual(float(por_sabor["Bacalhau"]["quantidade_tabuleiros"]), 0)
        self.assertAlmostEqual(float(por_sabor["Bacalhau"]["quantidade_unidades"]), 18)
        self.assertAlmostEqual(float(por_sabor["Calabresa"]["quantidade_unidades"]), 52)

        sistema.confirmar_producao_realizada_estoque(pre_id)
        movimentos = self.movimentos_producao(pre_id)
        self.assertEqual(len(movimentos), 2)
        self.assertTrue(all(mov["origem"] == "Produção Interna" for mov in movimentos))

        planejamento_final = self.query_one(
            "SELECT status FROM planejamentos_producao WHERE id = ?", (planejamento_id,)
        )
        self.assertEqual(planejamento_final["status"], "Produção confirmada")

        separacao = self.query_one(
            "SELECT status FROM separacoes_estoque WHERE planejamento_id = ?", (planejamento_id,)
        )
        self.assertIsNotNone(separacao)
        self.assertNotIn("Ã", separacao["status"])

    def test_bug_009_link_do_indicador_usa_rota_existente(self):
        template = Path(sistema.__file__).resolve().parent / "templates" / "indicadores.html"
        conteudo = template.read_text(encoding="utf-8")

        self.assertIn('/estoque/produtos/editar/{{ item.id }}', conteudo)
        self.assertNotIn('/estoque/produtos/{{ item.id }}/editar', conteudo)


if __name__ == "__main__":
    unittest.main()
