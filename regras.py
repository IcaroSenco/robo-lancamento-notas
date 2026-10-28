"""
Regras de cálculo das notas, compartilhadas pelo Sistema e pelo Robô.

Média do bimestre = (Nota x 8 + TF x 2) / 10
Média final       = (Média01 + Média02) / 2
Aprovado se Média final >= MEDIA_APROVACAO
"""

PESO_NOTA = 8
PESO_TF = 2
MEDIA_APROVACAO = 6.0


def ler_numero(texto):
    """Converte '7,5' ou '7.5' em 7.5 e valida a faixa 0 a 10."""
    if texto is None:
        raise ValueError("valor vazio")
    texto = str(texto).strip().replace(",", ".")
    if texto == "":
        raise ValueError("valor vazio")
    valor = float(texto)
    if not 0 <= valor <= 10:
        raise ValueError(f"nota fora da faixa 0-10: {valor}")
    return valor


def media_bimestre(nota, tf):
    return (nota * PESO_NOTA + tf * PESO_TF) / (PESO_NOTA + PESO_TF)


def media_final(nota1, tf1, nota2, tf2):
    return (media_bimestre(nota1, tf1) + media_bimestre(nota2, tf2)) / 2


def situacao(media):
    return "APROVADO" if round(media, 2) >= MEDIA_APROVACAO else "REPROVADO"
