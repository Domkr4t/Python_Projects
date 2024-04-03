import graphics as gr
import time

window = gr.GraphWin('Russian Game', 600, 600)
alpha = 0.1


def recurserectangle(A, B, C, D, deep=100):
    if deep == 0:
        return
    for M, N in (A, B), (B, C), (C, D), (D, A):
        gr.Line(gr.Point(*M), gr.Point(*N)).draw(window)
    A1 = (A[0] * (1 - alpha) + B[0] * alpha, A[1] * (1 - alpha) + B[1] * alpha)
    B1 = (B[0] * (1 - alpha) + C[0] * alpha, B[1] * (1 - alpha) + C[1] * alpha)
    C1 = (C[0] * (1 - alpha) + D[0] * alpha, C[1] * (1 - alpha) + D[1] * alpha)
    D1 = (D[0] * (1 - alpha) + A[0] * alpha, D[1] * (1 - alpha) + A[1] * alpha)
    recurserectangle(A1, B1, C1, D1, deep - 1)
    time.sleep(10)


recurserectangle((0, 0), (600, 0), (600, 600), (0, 600))

