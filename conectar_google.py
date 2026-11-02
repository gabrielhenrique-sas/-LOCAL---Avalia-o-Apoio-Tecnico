"""Autoriza (uma única vez) o acesso somente leitura à planilha de respostas e testa a leitura.

Uso:  python conectar_google.py
"""
import sys
from nucleo import config, google_planilhas as google


def main():
    cfg = config.carregar()
    try:
        linhas, aba, titulo = google.ler_respostas(cfg, interativo=True)
    except google.ErroGoogle as e:
        print("ERRO:", e)
        return 1
    print(f"OK! Planilha: {titulo} | aba: {aba} | {max(len(linhas) - 1, 0)} resposta(s)")
    if linhas:
        print("Colunas:")
        for i, h in enumerate(linhas[0]):
            print(f"  {i + 1}. {h}")
        try:
            from nucleo import respostas as dados
            mapa, av = dados.mapear_colunas(linhas[0], cfg.get("colunas"))
            print("Perguntas reconhecidas pelo programa: OK")
            for a in av:
                print("  aviso:", a)
        except Exception as e:
            print("ATENÇÃO - perguntas:", e)
    return 0


if __name__ == "__main__":
    sys.exit(main())
