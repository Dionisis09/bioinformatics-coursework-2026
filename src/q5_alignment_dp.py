from __future__ import annotations

from pathlib import Path


V = "ACTTGGGTG"
W = "GTGTGAATT"
MATCH = 1
MISMATCH = -1
GAP = -1

OUTPUT_DIR = Path("results/question5")
GLOBAL_PATH = OUTPUT_DIR / "global_alignment.txt"
LOCAL_PATH = OUTPUT_DIR / "local_alignment.txt"


def score(a: str, b: str) -> int:
    return MATCH if a == b else MISMATCH


def build_global():
    n, m = len(V), len(W)
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    back = [[[] for _ in range(m + 1)] for _ in range(n + 1)]

    for i in range(1, n + 1):
        dp[i][0] = dp[i - 1][0] + GAP
        back[i][0] = ["up"]
    for j in range(1, m + 1):
        dp[0][j] = dp[0][j - 1] + GAP
        back[0][j] = ["left"]

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            candidates = {
                "diag": dp[i - 1][j - 1] + score(V[i - 1], W[j - 1]),
                "up": dp[i - 1][j] + GAP,
                "left": dp[i][j - 1] + GAP,
            }
            best = max(candidates.values())
            dp[i][j] = best
            back[i][j] = [k for k, v in candidates.items() if v == best]

    i, j = n, m
    av, aw = [], []
    while i > 0 or j > 0:
        move = back[i][j][0]
        if move == "diag":
            av.append(V[i - 1])
            aw.append(W[j - 1])
            i -= 1
            j -= 1
        elif move == "up":
            av.append(V[i - 1])
            aw.append("-")
            i -= 1
        else:
            av.append("-")
            aw.append(W[j - 1])
            j -= 1

    return dp, "".join(reversed(av)), "".join(reversed(aw))


def build_local():
    n, m = len(V), len(W)
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    back = [[[] for _ in range(m + 1)] for _ in range(n + 1)]
    best_score = 0
    best_pos = (0, 0)

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            candidates = {
                "zero": 0,
                "diag": dp[i - 1][j - 1] + score(V[i - 1], W[j - 1]),
                "up": dp[i - 1][j] + GAP,
                "left": dp[i][j - 1] + GAP,
            }
            best = max(candidates.values())
            dp[i][j] = best
            back[i][j] = [k for k, v in candidates.items() if v == best and v != 0]
            if best > best_score:
                best_score = best
                best_pos = (i, j)

    i, j = best_pos
    av, aw = [], []
    while i > 0 and j > 0 and dp[i][j] > 0:
        move = back[i][j][0]
        if move == "diag":
            av.append(V[i - 1])
            aw.append(W[j - 1])
            i -= 1
            j -= 1
        elif move == "up":
            av.append(V[i - 1])
            aw.append("-")
            i -= 1
        else:
            av.append("-")
            aw.append(W[j - 1])
            j -= 1

    return dp, best_score, "".join(reversed(av)), "".join(reversed(aw))


def format_matrix(dp: list[list[int]], row_seq: str, col_seq: str) -> str:
    header = [" "] + list(col_seq)
    rows = ["\t".join(header)]
    rows.append("\t".join(["-"] + [str(x) for x in dp[0]]))
    for i, ch in enumerate(row_seq, start=1):
        rows.append("\t".join([ch] + [str(x) for x in dp[i]]))
    return "\n".join(rows)


def match_line(a: str, b: str) -> str:
    return "".join("|" if x == y else " " for x, y in zip(a, b))


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    global_dp, global_v, global_w = build_global()
    local_dp, local_score, local_v, local_w = build_local()

    global_text = "\n".join(
        [
            "Global alignment (Needleman-Wunsch)",
            f"v = {V}",
            f"w = {W}",
            f"Optimal score = {global_dp[len(V)][len(W)]}",
            "",
            format_matrix(global_dp, V, W),
            "",
            global_v,
            match_line(global_v, global_w),
            global_w,
            "",
        ]
    )
    local_text = "\n".join(
        [
            "Local alignment (Smith-Waterman)",
            f"v = {V}",
            f"w = {W}",
            f"Optimal score = {local_score}",
            "",
            format_matrix(local_dp, V, W),
            "",
            local_v,
            match_line(local_v, local_w),
            local_w,
            "",
        ]
    )

    GLOBAL_PATH.write_text(global_text, encoding="utf-8")
    LOCAL_PATH.write_text(local_text, encoding="utf-8")
    print(global_text)
    print(local_text)


if __name__ == "__main__":
    main()
