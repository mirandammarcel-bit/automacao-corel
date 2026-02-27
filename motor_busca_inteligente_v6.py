#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Motor de busca inteligente para integração com CorelDRAW.

Melhorias aplicadas sobre a versão original:
- Configuração tipada e persistente com validação.
- Remoção de chave de API hardcoded (segurança).
- Logging estruturado no lugar de prints soltos.
- Funções puras para normalização e score de busca.
- Escrita do arquivo VBA com sanitização e encoding compatível.
- Código pronto para evoluir para GUI (Tkinter) sem acoplamento.
"""

from __future__ import annotations

import json
import logging
import os
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("motor_busca")

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
PASTA_AUTOMACAO = Path(r"C:\AUTOMACAO_COREL")
ARQUIVO_DADOS_VBA = PASTA_AUTOMACAO / "dados.txt"
CONFIG_FILE = BASE_DIR / "config_motor_busca.json"
MAPEAMENTOS_FILE = BASE_DIR / "mapeamentos_aprendidos.json"

EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}

ABREVIACOES = {
    "ling": "linguica",
    "bov": "bovino",
    "sui": "suino",
    "fgo": "frango",
    "temp": "temperado",
    "cong": "congelado",
    "amac": "amaciante",
    "bisc": "biscoito",
    "refri": "refrigerante",
    "pct": "pacote",
    "cx": "caixa",
    "un": "unidade",
}

CORRECOES = {
    "pepeino": "pepino",
    "arrroz": "arroz",
    "smisrnof": "smirnoff",
    "espagute": "espaguete",
    "salgadimho": "salgadinho",
}

PALAVRAS_DISTINTIVAS = {
    "sem",
    "com",
    "inteiro",
    "inteira",
    "temperado",
    "congelado",
    "mignon",
    "sadia",
    "seara",
}

STOP_WORDS = {
    "de",
    "da",
    "do",
    "dos",
    "das",
    "e",
    "a",
    "o",
    "as",
    "os",
    "kg",
    "g",
    "ml",
    "l",
    "unidade",
}

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------


@dataclass
class Config:
    pasta_imagens: str = r"C:\Users\Public\Pictures"
    confianca_alta: int = 80
    confianca_media: int = 55
    remove_bg_key: str = ""  # intencionalmente vazio por segurança
    pastas_extras: list[str] = field(default_factory=list)

    @classmethod
    def load(cls, path: Path) -> "Config":
        if not path.exists():
            return cls()

        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return cls(**{**asdict(cls()), **data})  # type: ignore[arg-type]
        except Exception as exc:
            logger.warning("Falha ao carregar config (%s). Usando padrão.", exc)
            return cls()

    def save(self, path: Path) -> None:
        path.write_text(json.dumps(asdict(self), indent=2, ensure_ascii=False), encoding="utf-8")


# ---------------------------------------------------------------------------
# Texto e busca
# ---------------------------------------------------------------------------


def remover_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )


def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""

    texto = remover_acentos(texto.lower())
    texto = texto.replace("_", " ").replace("-", " ").replace(".", " ")
    texto = re.sub(r"\.(png|jpg|jpeg|gif|bmp|webp)$", "", texto)
    texto = re.sub(r"\s+\d+[,.]?\d*\s*(kg)?\s*$", "", texto)
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def aplicar_correcoes(texto: str) -> str:
    palavras = []
    for token in texto.split():
        if token in CORRECOES:
            palavras.append(CORRECOES[token])
            continue

        melhor = token
        melhor_ratio = 0.0
        for erro, correcao in CORRECOES.items():
            ratio = SequenceMatcher(None, token, erro).ratio()
            if ratio > 0.86 and ratio > melhor_ratio:
                melhor = correcao
                melhor_ratio = ratio
        palavras.append(melhor)

    return " ".join(palavras)


def expandir_abreviacoes(texto: str) -> str:
    return " ".join(ABREVIACOES.get(token, token) for token in texto.split()).strip()


def preprocessar(texto: str) -> str:
    return expandir_abreviacoes(aplicar_correcoes(normalizar_texto(texto)))


def extrair_palavras_chave(texto: str) -> set[str]:
    palavras = set(texto.split())
    return {
        p
        for p in palavras
        if (p in PALAVRAS_DISTINTIVAS or p not in STOP_WORDS) and (len(p) > 1 or p.isdigit())
    }


def calcular_pontuacao(texto_busca: str, texto_arquivo: str) -> float:
    busca = preprocessar(texto_busca)
    arquivo = preprocessar(texto_arquivo)

    palavras_busca = extrair_palavras_chave(busca)
    palavras_arquivo = extrair_palavras_chave(arquivo)
    if not palavras_busca or not palavras_arquivo:
        return 0.0

    comuns = palavras_busca & palavras_arquivo
    fuzzy_hits = 0
    for pb in palavras_busca - comuns:
        if any(SequenceMatcher(None, pb, pa).ratio() >= 0.82 for pa in palavras_arquivo - comuns):
            fuzzy_hits += 1

    cobertura_busca = (len(comuns) + (fuzzy_hits * 0.75)) / len(palavras_busca)
    cobertura_arquivo = (len(comuns) + (fuzzy_hits * 0.75)) / len(palavras_arquivo)
    sequencia = SequenceMatcher(None, busca, arquivo).ratio()
    bonus = len(comuns & PALAVRAS_DISTINTIVAS) * 0.04

    score = (cobertura_busca * 45) + (cobertura_arquivo * 20) + (sequencia * 30) + (bonus * 100)

    if busca == arquivo:
        return 100.0
    if palavras_busca <= palavras_arquivo:
        score += 12

    return round(min(100.0, score), 2)


def extrair_produto_e_preco(linha: str) -> tuple[str, str]:
    linha = linha.strip().replace("*", "")
    if not linha:
        return "", ""

    match = re.match(r"^(.+?)\s+(\d+[.,]\d{2}(?:kg)?)\s*$", linha)
    if match:
        return match.group(1).strip(), match.group(2).strip()

    return linha, ""


# ---------------------------------------------------------------------------
# Motor
# ---------------------------------------------------------------------------


class MotorBusca:
    def __init__(self) -> None:
        self.config = Config.load(CONFIG_FILE)
        self.arquivos: list[str] = []
        self.arquivos_completos: dict[str, Path] = {}
        self.mapeamentos: dict[str, str] = self._carregar_mapeamentos()

    def _carregar_mapeamentos(self) -> dict[str, str]:
        if not MAPEAMENTOS_FILE.exists():
            return {}
        try:
            return json.loads(MAPEAMENTOS_FILE.read_text(encoding="utf-8"))
        except Exception as exc:
            logger.warning("Mapeamentos inválidos, ignorando (%s).", exc)
            return {}

    def salvar_mapeamentos(self) -> None:
        MAPEAMENTOS_FILE.write_text(
            json.dumps(self.mapeamentos, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    def adicionar_mapeamento(self, termo: str, arquivo_rel: str) -> None:
        self.mapeamentos[preprocessar(termo)] = arquivo_rel
        self.salvar_mapeamentos()

    def _iterar_pastas(self) -> Iterable[Path]:
        yield Path(self.config.pasta_imagens)
        for pasta in self.config.pastas_extras:
            yield Path(pasta)

    def carregar_imagens(self) -> int:
        self.arquivos.clear()
        self.arquivos_completos.clear()

        for pasta in self._iterar_pastas():
            if not pasta.exists():
                continue

            for caminho in pasta.rglob("*"):
                if caminho.is_file() and caminho.suffix.lower() in EXTENSOES_IMAGEM:
                    try:
                        relativo = str(caminho.relative_to(pasta))
                    except ValueError:
                        relativo = caminho.name

                    chave = f"{pasta.name}/{relativo}".replace("\\", "/")
                    self.arquivos.append(chave)
                    self.arquivos_completos[chave] = caminho

        logger.info("%s imagens indexadas.", len(self.arquivos))
        return len(self.arquivos)

    def buscar(self, termo: str, top_n: int = 5) -> list[tuple[str, float]]:
        termo_norm = preprocessar(termo)

        if termo_norm in self.mapeamentos and self.mapeamentos[termo_norm] in self.arquivos:
            return [(self.mapeamentos[termo_norm], 100.0)]

        palavras = extrair_palavras_chave(termo_norm)
        candidatos = []

        if palavras:
            for arquivo in self.arquivos:
                arq_lower = arquivo.lower()
                if any(p in arq_lower for p in palavras):
                    candidatos.append(arquivo)
        else:
            candidatos = [a for a in self.arquivos if termo_norm in a.lower()]

        if not candidatos:
            candidatos = self.arquivos

        if len(candidatos) > 500:
            candidatos = candidatos[:500]

        resultados = []
        for arquivo in candidatos:
            score = calcular_pontuacao(termo, arquivo)
            if score >= 30:
                resultados.append((arquivo, score))

        resultados.sort(key=lambda x: x[1], reverse=True)
        return resultados[:top_n]

    def processar_lista(self, texto: str) -> list[dict]:
        resultados = []
        for linha in texto.splitlines():
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue

            produto, preco = extrair_produto_e_preco(linha)
            if not produto:
                continue

            partes = re.split(r"\s+ou\s+", produto, flags=re.IGNORECASE)
            for nome in [p.strip() for p in partes if p.strip()]:
                match = self.buscar(nome)
                melhor = match[0] if match else None
                pontuacao = melhor[1] if melhor else 0

                if pontuacao >= self.config.confianca_alta:
                    status = "OK"
                elif pontuacao >= self.config.confianca_media:
                    status = "REVISAR"
                elif pontuacao > 0:
                    status = "VERIFICAR"
                else:
                    status = "NAO_ENCONTRADO"

                resultados.append(
                    {
                        "produto": nome,
                        "preco": preco,
                        "match": melhor[0] if melhor else None,
                        "pontuacao": pontuacao,
                        "alternativas": match[1:] if len(match) > 1 else [],
                        "status": status,
                    }
                )
        return resultados

    def salvar_dados_vba(self, resultados: list[dict]) -> Path:
        PASTA_AUTOMACAO.mkdir(parents=True, exist_ok=True)

        with ARQUIVO_DADOS_VBA.open("w", encoding="cp1252", errors="replace") as f:
            for item in resultados:
                if item["status"] != "OK" or not item["match"]:
                    continue

                produto = (
                    str(item["produto"])
                    .replace(";", ",")
                    .replace("\n", " ")
                    .replace("\r", "")
                    .strip()
                )
                preco = str(item["preco"] or "0,00").replace(";", ",").strip()
                caminho = str(self.arquivos_completos.get(item["match"], Path(item["match"])))
                f.write(f"{produto};{preco};{caminho}\n")

        logger.info("Arquivo VBA salvo em: %s", ARQUIVO_DADOS_VBA)
        return ARQUIVO_DADOS_VBA


if __name__ == "__main__":
    print("Motor de Busca Inteligente v6 carregado com sucesso.")
    print("Este módulo está pronto para uso via GUI ou integração por script.")
