"""Testes unitários (funções puras; importa app no Windows com pywin32)."""
import os
import struct
import sys
from datetime import datetime
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app  # noqa: E402


class TestSanitizarNomeArquivo:
    def test_remove_invalidos(self):
        assert app._sanitizar_nome_arquivo('a<b>c:d/e') == "a_b_c_d_e"

    def test_vazio_vira_documento(self):
        assert app._sanitizar_nome_arquivo("   ") == "documento"

    def test_trunca(self):
        longo = "x" * 100
        assert len(app._sanitizar_nome_arquivo(longo, max_len=10)) == 10


class TestBytesParaLegivel:
    def test_bytes(self):
        assert app._bytes_para_kb_mb_gb(512) == "512 B"

    def test_kb(self):
        assert "KB" in app._bytes_para_kb_mb_gb(2048)

    def test_invalido(self):
        assert app._bytes_para_kb_mb_gb("x") == "0 B"


class TestChecarTipoSpool:
    def test_emf_header(self, tmp_path):
        p = tmp_path / "t.spl"
        p.write_bytes(struct.pack("<II", 0, 0x00010000) + b"\x00" * 8)
        assert app._checar_tipo_spool(str(p)) == "EMF"

    def test_desconhecido(self, tmp_path):
        p = tmp_path / "t.spl"
        p.write_bytes(b"\xff\xff\xff\xff" * 2)
        assert app._checar_tipo_spool(str(p)) == "DESCONHECIDO"

    def test_inexistente(self):
        assert app._checar_tipo_spool(r"C:\nao_existe_12345.spl") == "DESCONHECIDO"


class TestCaminhoLogDoDia:
    def test_formato_nome(self):
        dt = datetime(2026, 2, 20, 12, 0, 0)
        path = app.caminho_log_do_dia(dt)
        assert path.endswith(os.path.join(app.PASTA_SCRIPT, "log_impressoes_20022026.xlsx"))


class TestCaminhoArquivoCopia:
    def test_extensao_spl(self):
        path = app.caminho_arquivo_copia(42, "doc.pdf", "HP Laser")
        assert path.endswith(".spl")
        assert "42_" in os.path.basename(path)


class TestSplMaisRecente:
    def test_escolhe_mais_novo(self, tmp_path, monkeypatch):
        import time as time_mod

        spool = tmp_path / "spool"
        spool.mkdir()
        antigo = spool / "00001.SPL"
        novo = spool / "00002.SPL"
        antigo.write_bytes(b"x")
        novo.write_bytes(b"y")
        t0 = time_mod.time()
        os.utime(antigo, (t0 - 20, t0 - 20))
        os.utime(novo, (t0 - 1, t0 - 1))
        got = app._spl_mais_recente(str(spool), max_idade_seg=30.0)
        assert got == str(novo)

    def test_colunas_dec_001(self):
        assert app.CABECALHO == [
            "ID_Job",
            "Usuario",
            "Data_Hora",
            "Arquivo",
            "Paginas",
            "Impressora",
            "Tamanho_Bytes",
            "Tamanho_Legivel",
            "Local_Arquivo",
            "Local_Preview",
        ]
