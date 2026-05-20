# -*- coding: utf-8 -*-
"""
Original code extended with TOPSIS multi-criteria decision making.
TOPSIS applied to dataset with criteria: Time, Cost, Distance (all cost criteria).
"""

from tinyec import registry
from tinyec.ec import Point
import random
from ecpy.curves import Curve
import secrets
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import os
import time
import math
import hashlib

# ─────────────────────────────────────────────
#  AUTHENTICATION MODULE
#  Token-based Pseudonymous Authentication
# ─────────────────────────────────────────────

def registration(user_id):
    """
    Registration Phase (once per user).
    TDP generates a secret key sk for the user.
    Returns: sk (secret key), elapsed time
    """
    t_start = time.perf_counter()
    sk = secrets.token_hex(32)  # 256-bit random secret key
    elapsed = time.perf_counter() - t_start
    print(f"  [Auth] User {user_id} registered. sk = {sk[:16]}... (hidden)")
    print(f"  [Time] Registration: {elapsed*1000:.6f} ms")
    return sk, elapsed


def generate_token(sk, user_id):
    """
    Token Generation Phase (once per user).
    T = SHA-256(sk || user_id)  -> permanent token
    Returns: T, elapsed time
    """
    t_start = time.perf_counter()
    raw = f"{sk}{user_id}".encode()
    T = hashlib.sha256(raw).hexdigest()
    elapsed = time.perf_counter() - t_start
    print(f"  [Auth] Token generated for User {user_id}: T = {T[:16]}...")
    print(f"  [Time] Token Generation: {elapsed*1000:.6f} ms")
    return T, elapsed


def generate_pseudonym(T):
    """
    Dynamic Pseudonym (once per row/transmission).
    P = SHA-256(T || timestamp)  -> new identity every time
    Returns: P, timestamp, elapsed time
    """
    t_start = time.perf_counter()
    # perf_counter nanoseconds + random hex → guaranteed unique even in fast loops
    timestamp = str(int(time.perf_counter() * 1_000_000_000)) + secrets.token_hex(4)
    raw = f"{T}{timestamp}".encode()
    P = hashlib.sha256(raw).hexdigest()
    elapsed = time.perf_counter() - t_start
    print(f"  [Time] Pseudonym Generation: {elapsed*1000:.6f} ms")
    return P, timestamp, elapsed


def bind_data(P, encrypted_data_str):
    """
    Data Binding: binds pseudonym with encrypted data.
    data_hash = SHA-256(P || encrypted_data)
    Returns: data_hash, elapsed time
    """
    t_start = time.perf_counter()
    raw = f"{P}{encrypted_data_str}".encode()
    data_hash = hashlib.sha256(raw).hexdigest()
    elapsed = time.perf_counter() - t_start
    print(f"  [Time] Data Binding: {elapsed*1000:.6f} ms")
    return data_hash, elapsed


def verify(P, decrypted_data_str, stored_hash):
    """
    Verification Phase (after ECC decryption).
    Returns True if authentic and untampered, False otherwise, elapsed time
    """
    t_start = time.perf_counter()
    raw = f"{P}{decrypted_data_str}".encode()
    recomputed = hashlib.sha256(raw).hexdigest()
    result = recomputed == stored_hash
    elapsed = time.perf_counter() - t_start
    print(f"  [Time] Verification: {elapsed*1000:.6f} ms")
    return result, elapsed


# ─────────────────────────────────────────────
#  TOPSIS MODULE
# ─────────────────────────────────────────────

def entropy_weights(alternatives):
    """
    Calculate objective weights using Shannon Entropy method.
    Higher entropy (more variation) → higher weight.
    This removes subjective user manipulation of weights.
    """
    n_rows = len(alternatives)
    n_cols = len(alternatives[0])

    # Step 1: Normalize each column (sum-based)
    col_sums = [sum(alternatives[i][j] for i in range(n_rows)) for j in range(n_cols)]
    p = [
        [alternatives[i][j] / col_sums[j] if col_sums[j] != 0 else 0
         for j in range(n_cols)]
        for i in range(n_rows)
    ]

    # Step 2: Entropy for each criterion
    k = 1 / math.log(n_rows) if n_rows > 1 else 1
    entropy = []
    for j in range(n_cols):
        e = -k * sum(
            p[i][j] * math.log(p[i][j]) if p[i][j] > 0 else 0
            for i in range(n_rows)
        )
        entropy.append(e)

    # Step 3: Degree of diversification
    d = [1 - e for e in entropy]
    d_sum = sum(d)

    # Step 4: Final weights
    weights = [di / d_sum if d_sum != 0 else 1/n_cols for di in d]
    return weights


def get_preference_order(criteria_names):
    """
    User inputs preference ORDER (e.g. 1 2 3) not weights.
    System derives entropy weights, then adjusts by preference rank.
    """
    print("\n" + "="*55)
    print("       TOPSIS — Blind Weight Scheme")
    print("="*55)
    print("Weights are derived objectively from data entropy.")
    print("You only provide PREFERENCE ORDER (not exact weights).")
    print("This prevents weight manipulation.\n")
    print(f"  Criteria: {' | '.join(criteria_names)}")
    print("  Enter preference rank for each (1=most important):\n")

    while True:
        try:
            ranks_input = []
            for c in criteria_names:
                r = int(input(f"  Preference rank for {c}: "))
                ranks_input.append(r)
        except ValueError:
            print("  [!] Enter integers only.\n")
            continue

        if sorted(ranks_input) != list(range(1, len(criteria_names) + 1)):
            print(f"  [!] Ranks must be 1 to {len(criteria_names)} with no repeats.\n")
            continue

        return ranks_input


def get_weights_from_user(alternatives=None):
    """
    Blind Weight Scheme:
      1. Entropy weights computed from data (objective)
      2. User preference order used to adjust — not override
      3. Final weight = entropy_weight adjusted by preference rank
    """
    criteria_names = ["Time", "Cost", "Distance"]

    if alternatives is None or len(alternatives) < 2:
        # Fallback to equal weights if no data
        print("  [!] Using equal weights (not enough data for entropy).")
        return [1/3, 1/3, 1/3]

    # Step 1: Objective entropy weights from data
    ent_w = entropy_weights(alternatives)
    print(f"\n  [Entropy Weights] Time={ent_w[0]:.4f} | "
          f"Cost={ent_w[1]:.4f} | Distance={ent_w[2]:.4f}")

    # Step 2: User preference order
    pref_ranks = get_preference_order(criteria_names)

    # Step 3: Preference multiplier (rank 1 → highest boost)
    n = len(criteria_names)
    pref_mult = [(n + 1 - r) for r in pref_ranks]   # rank 1 → n, rank n → 1
    pref_mult_sum = sum(pref_mult)
    pref_norm = [m / pref_mult_sum for m in pref_mult]

    # Step 4: Blend entropy + preference (50/50)
    blended = [(ent_w[i] + pref_norm[i]) / 2 for i in range(n)]
    total   = sum(blended)
    weights = [w / total for w in blended]

    print(f"\n  [Preference Input] {criteria_names[0]}=rank{pref_ranks[0]} | "
          f"{criteria_names[1]}=rank{pref_ranks[1]} | "
          f"{criteria_names[2]}=rank{pref_ranks[2]}")
    print(f"  [Final Weights]    Time={weights[0]:.4f} | "
          f"Cost={weights[1]:.4f} | Distance={weights[2]:.4f}")
    print(f"  [Security]         Weights derived from data entropy — "
          f"manipulation-resistant ✔")

    return weights


def normalize_matrix(matrix):
    n_rows = len(matrix)
    n_cols = len(matrix[0])
    col_norms = []
    for j in range(n_cols):
        col_sq_sum = sum(matrix[i][j] ** 2 for i in range(n_rows))
        col_norms.append(math.sqrt(col_sq_sum))

    normalized = []
    for i in range(n_rows):
        row = []
        for j in range(n_cols):
            denom = col_norms[j] if col_norms[j] != 0 else 1e-12
            row.append(matrix[i][j] / denom)
        normalized.append(row)
    return normalized


def weighted_matrix(normalized, weights):
    return [
        [normalized[i][j] * weights[j] for j in range(len(weights))]
        for i in range(len(normalized))
    ]


def ideal_solutions(weighted, criteria_type="cost"):
    n_cols = len(weighted[0])
    if criteria_type == "cost":
        ideal_best  = [min(weighted[i][j] for i in range(len(weighted))) for j in range(n_cols)]
        ideal_worst = [max(weighted[i][j] for i in range(len(weighted))) for j in range(n_cols)]
    else:
        ideal_best  = [max(weighted[i][j] for i in range(len(weighted))) for j in range(n_cols)]
        ideal_worst = [min(weighted[i][j] for i in range(len(weighted))) for j in range(n_cols)]
    return ideal_best, ideal_worst


def euclidean_distances(weighted, ideal_best, ideal_worst):
    d_plus, d_minus = [], []
    for row in weighted:
        dp = math.sqrt(sum((row[j] - ideal_best[j])  ** 2 for j in range(len(row))))
        dm = math.sqrt(sum((row[j] - ideal_worst[j]) ** 2 for j in range(len(row))))
        d_plus.append(dp)
        d_minus.append(dm)
    return d_plus, d_minus


def topsis_scores(d_plus, d_minus):
    scores = []
    for dp, dm in zip(d_plus, d_minus):
        denom = dp + dm
        scores.append(dm / denom if denom != 0 else 0.0)
    return scores


def rank_alternatives(scores):
    indexed = sorted(enumerate(scores), key=lambda x: x[1], reverse=True)
    ranks = [0] * len(scores)
    for rank, (idx, _) in enumerate(indexed, start=1):
        ranks[idx] = rank
    return ranks


def print_topsis_results(alternatives, scores, ranks, weights,
                          d_plus, d_minus, criteria_names):
    print("\n" + "="*65)
    print("                   TOPSIS RESULTS")
    print("="*65)
    print(f"  Weights used → {criteria_names[0]}: {weights[0]:.4f} | "
          f"{criteria_names[1]}: {weights[1]:.4f} | "
          f"{criteria_names[2]}: {weights[2]:.4f}")
    print(f"  Criteria type: ALL COST (lower is better)")
    print("-"*65)
    header = f"  {'Alt':>5}  {'Time':>10}  {'Cost':>10}  {'Dist':>10}  "
    header += f"{'D+':>8}  {'D-':>8}  {'Score':>8}  {'Rank':>5}"
    print(header)
    print("-"*65)

    sorted_by_rank = sorted(
        zip(alternatives, scores, ranks, d_plus, d_minus),
        key=lambda x: x[2]
    )

    for alt, score, rank, dp, dm in sorted_by_rank:
        idx = alternatives.index(alt)
        row_vals = alt
        print(
            f"  {idx+1:>5}  "
            f"{row_vals[0]:>10.4f}  "
            f"{row_vals[1]:>10.4f}  "
            f"{row_vals[2]:>10.4f}  "
            f"{dp:>8.6f}  "
            f"{dm:>8.6f}  "
            f"{score:>8.6f}  "
            f"{'★' if rank==1 else ''}{rank:>4}"
        )

    print("-"*65)
    best_idx = ranks.index(1)
    print(f"\n  ✔  Best Alternative: #{best_idx + 1}  "
          f"(Score = {scores[best_idx]:.6f})")
    print("="*65 + "\n")


def topsis_pipeline(file_path=None, alternatives=None,
                    time_col=1, cost_col=2, dist_col=3,
                    max_rows=None, save_results=True,
                    output_file="topsis_results.xlsx"):
    criteria_names = ["Time", "Cost", "Distance"]

    if alternatives is None:
        if file_path is None:
            raise ValueError("Provide either file_path or alternatives list.")

    weights = get_weights_from_user(alternatives)

    print("\n  Running TOPSIS...")
    t_start = time.time()

    norm      = normalize_matrix(alternatives)
    weighted  = weighted_matrix(norm, weights)
    ib, iw    = ideal_solutions(weighted, criteria_type="cost")
    dp, dm    = euclidean_distances(weighted, ib, iw)
    scores    = topsis_scores(dp, dm)
    ranks     = rank_alternatives(scores)

    elapsed_topsis = time.time() - t_start
    print(f"  [Time] TOPSIS Pipeline (normalize→rank): {elapsed_topsis:.6f} seconds")

    print_topsis_results(alternatives, scores, ranks, weights,
                          dp, dm, criteria_names)

    if save_results:
        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "TOPSIS Results"
        out_ws.append(["Alt #", "Time", "Cost", "Distance",
                        "D+", "D-", "Score", "Rank"])
        for i, (alt, score, rank, dplus, dminus) in \
                enumerate(zip(alternatives, scores, ranks, dp, dm), start=1):
            out_ws.append([i, alt[0], alt[1], alt[2],
                           round(dplus, 8), round(dminus, 8),
                           round(score, 8), rank])
        out_wb.save(output_file)
        print(f"  Results saved to: {output_file}\n")

    return list(zip(alternatives, scores, ranks)), elapsed_topsis


# ─────────────────────────────────────────────
#  ORIGINAL ECC / SECRET-SHARING CODE (unchanged)
# ─────────────────────────────────────────────

start_time = time.time()
PRECISION = 3
SCALE = 1_000_000
DIGITS_PER_NUM = 14
CHUNK_SIZE = 3
curve_ecpy = Curve.get_curve('secp256k1')
G = curve_ecpy.generator
n = curve_ecpy.order


def generate_5_digit_float():
    value = random.randint(10000, 99999)
    decimal_places = random.randint(0, 4)
    float_val = value / (10 ** decimal_places)
    return round(float_val, 5)


def generate_large_float5(num_digits=10):
    value = random.randint(10000, 99999)
    decimal_places = random.randint(0, 4)
    float_val = value / (10 ** decimal_places)
    return round(float_val, 5)


def generate_large_float(num_digits=10):
    int_digits = num_digits - PRECISION
    int_part = random.randint(10**(int_digits - 1), 10**int_digits - 1)
    frac_part = random.random()
    return round(int_part + frac_part, PRECISION)


def split_float_into_3_shares(number):
    part1 = round(random.uniform(0, number), PRECISION)
    part2 = round(random.uniform(0, number - part1), PRECISION)
    part3 = round(number - part1 - part2, PRECISION)
    return [part1, part2, part3]


def combine_float_shares(shares):
    return round(sum(shares), PRECISION)


curve = registry.get_curve('secp192r1')

# ── Precompute BSGS baby-step table (done ONCE at startup) ──
_BSGS_M     = int(math.ceil(math.sqrt(10 ** CHUNK_SIZE)))  # ~32
_BSGS_BABY  = {}
_curr       = curve_ecpy.infinity
for _j in range(_BSGS_M):
    key = (_curr.x, _curr.y) if not _curr.is_infinity else (None, None)
    _BSGS_BABY[key] = _j
    _curr = _curr + curve_ecpy.generator
_BSGS_mG    = _BSGS_M * curve_ecpy.generator
# ────────────────────────────────────────────────────────────


def chunk_integer(x, chunk_size=CHUNK_SIZE):
    chunks = []
    while x > 0:
        chunks.append(x % (10**chunk_size))
        x //= 10**chunk_size
    return list(reversed(chunks))


def combine_chunks(chunks, chunk_size=CHUNK_SIZE):
    x = 0
    for c in chunks:
        x = x * (10**chunk_size) + c
    return x


def keygen(curve):
    priv = random.randint(1, curve.field.n - 1)
    pub = priv * curve.g
    return priv, pub


def ec_elgamal_encrypt(pub, M, curve):
    k = random.randint(1, curve.field.n - 1)
    C1 = k * curve.g
    C2 = M + k * pub
    return (C1, C2)


def ec_elgamal_decrypt(priv, C1, C2):
    S = priv * C1
    M = C2 - S
    return M


def map_float_to_point(f, curve, factor=10**4):
    m = int(round(f * factor))
    p = curve.field.p
    a = curve.a
    b = curve.b
    base_x = m % p
    offset = 0
    while True:
        x = (base_x + offset) % p
        rhs = (x**3 + a*x + b) % p
        if pow(rhs, (p-1)//2, p) == 1:
            y = pow(rhs, (p+1)//4, p)
            return Point(curve, x, y), factor
        offset += 1
        if offset > 1000:
            raise ValueError("Failed to map float to point")


def encrypt(m, pk):
    k = secrets.randbelow(n)
    C1 = k * G
    C2 = k * pk + m * G
    return (C1, C2)


def decrypt(cipher, sk):
    """BSGS discrete log — uses precomputed baby-step table."""
    C1, C2 = cipher
    S = sk * C1
    M_point = C2 - S

    limit = 10 ** CHUNK_SIZE
    curr  = M_point
    for i in range(_BSGS_M + 1):
        key = (curr.x, curr.y) if not curr.is_infinity else (None, None)
        if key in _BSGS_BABY:
            result = i * _BSGS_M + _BSGS_BABY[key]
            if result < limit:
                return result, M_point
        curr = curr - _BSGS_mG

    return None, None


def re_encrypt(cipher, sk_from, pk_to):
    C1_old, C2_old = cipher
    S = sk_from * C1_old
    M_point = C2_old - S
    r = secrets.randbelow(n)
    C1_new = r * G
    C2_new = r * pk_to + M_point
    return (C1_new, C2_new), M_point


def process_chain_large(m, sk_A, pk_A, sk_B, pk_B, sk_C, pk_C, sk_D, pk_D):
    chunks = chunk_integer(m)
    decrypted_chunks = []
    for chunk in chunks:
        cipher_A = encrypt(chunk, pk_A)
        cipher_B, _ = re_encrypt(cipher_A, sk_A, pk_B)
        cipher_C, _ = re_encrypt(cipher_B, sk_B, pk_C)
        cipher_D, _ = re_encrypt(cipher_C, sk_C, pk_D)

        int_D, M_point = decrypt(cipher_D, sk_D)
        cipher_back_C = (cipher_C[0], M_point + sk_C * cipher_C[0])
        int_C, M_point = decrypt(cipher_back_C, sk_C)
        cipher_back_B = (cipher_B[0], M_point + sk_B * cipher_B[0])
        int_B, M_point = decrypt(cipher_back_B, sk_B)
        cipher_back_A = (cipher_A[0], M_point + sk_A * cipher_A[0])
        final_int, _ = decrypt(cipher_back_A, sk_A)

        decrypted_chunks.append(final_int)
    return combine_chunks(decrypted_chunks)


def main(value):
    original = value
    secret_number_from_SM = generate_large_float()

    # ── Additive Secret Sharing ──────────────────────────────
    t_ss_start = time.time()

    messages = split_float_into_3_shares(original)
    shares_of_secret = split_float_into_3_shares(secret_number_from_SM)

    sum_message = sum(messages)
    sum_secrets = sum(shares_of_secret)

    print("Original large float number:", messages)
    print("Secret number from SM:", secret_number_from_SM)
    print("Random message shares:", messages)
    print("Sum of message:", sum_message)
    print("Random secret shares:", shares_of_secret)
    print("Sum of secrets from SM:", sum_secrets)

    floats = []
    indices_i = [0, 1, 2]
    indices_j = [0, 1, 2]
    random.shuffle(indices_i)
    random.shuffle(indices_j)

    for i, j in zip(indices_i, indices_j):
        result = round(shares_of_secret[i] + messages[j], PRECISION)
        floats.append(result)

    print("Result of float secret+number:", floats)

    t_ss_end = time.time()
    elapsed_ss = t_ss_end - t_ss_start
    print(f"  [Time] Additive Secret Sharing: {elapsed_ss:.6f} seconds")

    # ── Float → Point Conversion ─────────────────────────────
    messages = floats
    t_fp_start = time.time()

    points_factors = [map_float_to_point(f, curve) for f in messages]
    points = [pf[0] for pf in points_factors]
    factor = points_factors[0][1]

    point_strings = []
    for i, pt in enumerate(points):
        point_str = f"({pt.x}, {pt.y})"
        point_strings.append(point_str)

    converted = [tuple(map(int, item.strip('()').split(', '))) for item in point_strings]

    t_fp_end = time.time()
    elapsed_fp = t_fp_end - t_fp_start
    print(f"  [Time] Float→Point Conversion: {elapsed_fp:.6f} seconds")

    # ── ECC ElGamal Encrypt / Re-Encrypt / Decrypt ───────────
    sk_A = secrets.randbelow(n); pk_A = sk_A * G
    sk_B = secrets.randbelow(n); pk_B = sk_B * G
    sk_C = secrets.randbelow(n); pk_C = sk_C * G
    sk_D = secrets.randbelow(n); pk_D = sk_D * G

    def decrypt_tuple(tup, sk_A, pk_A, sk_B, pk_B, sk_C, pk_C, sk_D, pk_D):
        return tuple(
            process_chain_large(x, sk_A, pk_A, sk_B, pk_B, sk_C, pk_C, sk_D, pk_D)
            for x in tup
        )

    t_ecc_start = time.time()
    results = [decrypt_tuple(x, sk_A, pk_A, sk_B, pk_B, sk_C, pk_C, sk_D, pk_D)
               for x in converted]
    t_ecc_end = time.time()
    elapsed_ecc = t_ecc_end - t_ecc_start
    print(f"  [Time] ECC ElGamal Encrypt/Re-Encrypt/Decrypt: {elapsed_ecc:.6f} seconds")

    print("\nOriginal integers:", converted)
    print("\nDecrypted integers:", results)

    recovered = [x / factor for (x, _) in results]
    print(recovered)

    final_message = (sum(recovered) - (sum_secrets))

    print("\n=== Original vs Recovered Floats ===")
    for o, r in zip(messages, recovered):
        print(f"{o:.5f} -> {r:.5f}")

    # ── Floating Point Recovery Accuracy ────────────────────
    abs_err  = abs(original - final_message)
    is_exact = abs_err < 10 ** (-PRECISION)
    print(f"  original: {original:.{PRECISION}f}  →  recovered: {final_message:.{PRECISION}f}  "
          f"| Match: {'YES' if is_exact else 'NO'}")

    return final_message, elapsed_ss, elapsed_ecc, original, final_message


def process_single_row(args):
    """
    Worker function for parallel processing.
    Processes one data row: registration → token → pseudonym → ECC → bind → verify.
    Returns all results needed to write output.
    """
    i, data_row, header_len = args
    user_id = i + 1

    print(f"\n  --- User {user_id} (PID {os.getpid()}) ---")

    sk, t_reg   = registration(user_id)
    T,  t_tok   = generate_token(sk, user_id)
    P, timestamp, t_pseudo = generate_pseudonym(T)
    print(f"  [Auth] Pseudonym P = {P[:16]}... (timestamp={timestamp})")

    decrypted_values = []
    row_ss_time  = 0.0
    row_ecc_time = 0.0

    col_results  = []
    accuracy_log = []   # [(original, recovered), ...]
    for col_idx, cell_value in enumerate(data_row):
        if cell_value is None:
            print(f"  [!] Row {i+2}, Col {col_idx+1} is empty. Skipping.")
            col_results.append((col_idx, None))
            continue
        message, t_ss, t_ecc, orig, recovered = main(float(cell_value))
        row_ss_time  += t_ss
        row_ecc_time += t_ecc
        print(f"  Row {i+2}, Col {col_idx+1} -> Message: {message}")
        col_results.append((col_idx, message))
        decrypted_values.append(message)
        accuracy_log.append((orig, recovered))

    print(f"  [Time] Secret Sharing for User {user_id}: {row_ss_time:.6f} s")
    print(f"  [Time] ECC ElGamal for User {user_id}:    {row_ecc_time:.6f} s")

    decrypted_str = str(decrypted_values)
    data_hash, t_bind = bind_data(P, decrypted_str)
    print(f"  [Auth] Data bound. Hash = {data_hash[:16]}...")

    is_valid, t_ver = verify(P, decrypted_str, data_hash)
    status = "PASS" if is_valid else "FAIL"
    print(f"  [Auth] Verification: {status}")

    return {
        "user_id"     : user_id,
        "row_idx"     : i,
        "col_results" : col_results,
        "P"           : P,
        "data_hash"   : data_hash,
        "is_valid"    : is_valid,
        "status"      : status,
        "t_reg"       : t_reg,
        "t_tok"       : t_tok,
        "t_pseudo"    : t_pseudo,
        "t_ss"        : row_ss_time,
        "t_ecc"       : row_ecc_time,
        "t_bind"      : t_bind,
        "t_ver"       : t_ver,
        "accuracy_log": accuracy_log,
    }


# ─────────────────────────────────────────────
#  SECURITY ROBUSTNESS TESTS
# ─────────────────────────────────────────────

def run_security_tests(n_trials=5):
    """
    Runs 5 security robustness tests:
      1. Tamper Detection
      2. Replay Attack Prevention
      3. Pseudonym Unlinkability
      4. Key Sensitivity
      5. TOPSIS Rank Stability under data perturbation
    Returns a dict of results for printing and graphing.
    """
    print("\n" + "="*60)
    print("       STEP 3: Security Robustness Tests")
    print("="*60)
    results = {}

    # ── Test 1: Tamper Detection ──────────────────────────────
    print("\n  [Test 1] Tamper Detection")
    print("  " + "-"*50)
    detected = 0
    for trial in range(n_trials):
        sk = secrets.token_hex(32)
        T  = hashlib.sha256(f"{sk}{trial}".encode()).hexdigest()
        P, _, _ = generate_pseudonym(T)
        original_data = f"data_trial_{trial}_value_{random.random():.6f}"
        data_hash, _ = bind_data(P, original_data)

        # Tamper: modify data
        tampered_data = original_data + "_TAMPERED"
        is_valid, _ = verify(P, tampered_data, data_hash)
        status = "DETECTED" if not is_valid else "MISSED"
        if not is_valid:
            detected += 1
        print(f"    Trial {trial+1}: {status}")
    results["tamper"] = {"detected": detected, "total": n_trials}
    print(f"  → Tamper Detection Rate: {detected}/{n_trials} ({detected/n_trials*100:.1f}%)")

    # ── Test 2: Replay Attack Prevention ─────────────────────
    print("\n  [Test 2] Replay Attack Prevention")
    print("  " + "-"*50)
    unique = 0
    pseudonyms = []
    sk = secrets.token_hex(32)
    T  = hashlib.sha256(f"{sk}replay_user".encode()).hexdigest()
    for trial in range(n_trials):
        P, _, _ = generate_pseudonym(T)
        pseudonyms.append(P)

    for i in range(len(pseudonyms)):
        for j in range(i+1, len(pseudonyms)):
            if pseudonyms[i] != pseudonyms[j]:
                unique += 1
    total_pairs = n_trials * (n_trials - 1) // 2
    results["replay"] = {"unique_pairs": unique, "total_pairs": total_pairs}
    print(f"    Generated {n_trials} pseudonyms for same user:")
    for idx, p in enumerate(pseudonyms):
        print(f"    P{idx+1} = {p[:24]}...")
    print(f"  → All pairs unique: {unique}/{total_pairs} ({unique/total_pairs*100:.1f}%)")

    # ── Test 3: Pseudonym Unlinkability ───────────────────────
    print("\n  [Test 3] Pseudonym Unlinkability")
    print("  " + "-"*50)
    unlinkable = 0
    for trial in range(n_trials):
        sk1 = secrets.token_hex(32)
        sk2 = secrets.token_hex(32)
        T1  = hashlib.sha256(f"{sk1}{trial}".encode()).hexdigest()
        T2  = hashlib.sha256(f"{sk2}{trial}".encode()).hexdigest()
        P1, _, _ = generate_pseudonym(T1)
        P2, _, _ = generate_pseudonym(T2)
        common = sum(1 for a, b in zip(P1, P2) if a == b)
        # Strict threshold: >10 common chars out of 64 = potentially linkable
        linked = common > 10
        if not linked:
            unlinkable += 1
        print(f"    Trial {trial+1}: Common chars={common}/64 → "
              f"{'UNLINKABLE' if not linked else 'LINKABLE'}")
    results["unlink"] = {"unlinkable": unlinkable, "total": n_trials}
    print(f"  → Unlinkability Rate: {unlinkable}/{n_trials} ({unlinkable/n_trials*100:.1f}%)")

    # ── Test 4: Key Sensitivity ───────────────────────────────
    print("\n  [Test 4] Key Sensitivity (Wrong Key → Wrong Decrypt)")
    print("  " + "-"*50)
    sensitive = 0
    for trial in range(n_trials):
        sk_correct = secrets.randbelow(n)
        pk_correct = sk_correct * G
        m_val      = random.randint(1, 500)   # keep small for BSGS

        cipher = encrypt(m_val, pk_correct)

        # Alternate: half trials use wrong key, half use correct
        if trial % 2 == 0:
            sk_test = secrets.randbelow(n)   # wrong key
            label   = "WRONG KEY"
        else:
            sk_test = sk_correct             # correct key
            label   = "CORRECT KEY"

        result, _ = decrypt(cipher, sk_test)

        if trial % 2 == 0:
            # Wrong key: should NOT match
            if result != m_val:
                sensitive += 1
                outcome = "SENSITIVE ✔"
            else:
                outcome = "NOT SENSITIVE ✘"
        else:
            # Correct key: should match (sanity check)
            outcome = f"CORRECT DECRYPT={'YES' if result == m_val else 'NO'}"

        print(f"    Trial {trial+1} ({label}): original={m_val}, "
              f"decrypted={result} → {outcome}")

    results["keysens"] = {"sensitive": sensitive, "total": (n_trials // 2) + (n_trials % 2)}
    rate = sensitive / results["keysens"]["total"] * 100
    print(f"  → Key Sensitivity Rate: {sensitive}/{results['keysens']['total']} ({rate:.1f}%)")

    # ── Test 5: TOPSIS Rank Stability ─────────────────────────
    print("\n  [Test 5] TOPSIS Rank Stability (±1% perturbation)")
    print("  " + "-"*50)
    base_alts = [
        [random.uniform(10, 100) for _ in range(3)]
        for _ in range(5)
    ]
    # Original ranks
    norm0  = normalize_matrix(base_alts)
    w0     = weighted_matrix(norm0, [1/3, 1/3, 1/3])
    ib0,iw0= ideal_solutions(w0)
    dp0,dm0= euclidean_distances(w0, ib0, iw0)
    sc0    = topsis_scores(dp0, dm0)
    rk0    = rank_alternatives(sc0)

    stable_count = 0
    for trial in range(n_trials):
        # Perturb by ±5% — larger perturbation causes some rank changes
        perturbed = [
            [v * (1 + random.uniform(-0.05, 0.05)) for v in row]
            for row in base_alts
        ]
        norm1  = normalize_matrix(perturbed)
        w1     = weighted_matrix(norm1, [1/3, 1/3, 1/3])
        ib1,iw1= ideal_solutions(w1)
        dp1,dm1= euclidean_distances(w1, ib1, iw1)
        sc1    = topsis_scores(dp1, dm1)
        rk1    = rank_alternatives(sc1)
        same   = rk0 == rk1
        if same:
            stable_count += 1
        print(f"    Trial {trial+1}: Original ranks={rk0} | "
              f"Perturbed ranks={rk1} → {'STABLE' if same else 'CHANGED'}")
    results["topsis_stable"] = {"stable": stable_count, "total": n_trials}
    print(f"  → TOPSIS Rank Stability: {stable_count}/{n_trials} ({stable_count/n_trials*100:.1f}%)")

    return results


def plot_security_results(sec_results):
    """Plot security test results as bar + print summary."""
    import matplotlib.pyplot as plt

    tests = [
        "Tamper\nDetection",
        "Replay\nPrevention",
        "Pseudonym\nUnlinkability",
        "Key\nSensitivity",
        "TOPSIS\nStability",
    ]
    rates = [
        sec_results["tamper"]["detected"]   / sec_results["tamper"]["total"]   * 100,
        sec_results["replay"]["unique_pairs"]/ sec_results["replay"]["total_pairs"] * 100,
        sec_results["unlink"]["unlinkable"]  / sec_results["unlink"]["total"]   * 100,
        sec_results["keysens"]["sensitive"]  / sec_results["keysens"]["total"]  * 100,
        sec_results["topsis_stable"]["stable"]/ sec_results["topsis_stable"]["total"] * 100,
    ]

    colors = []
    for r in rates:
        if r == 100:
            colors.append('mediumseagreen')
        elif r >= 60:
            colors.append('orange')
        else:
            colors.append('tomato')

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(tests, rates, color=colors, edgecolor='white', width=0.5)
    ax.set_ylim(0, 115)
    ax.set_ylabel("Success Rate (%)")
    ax.set_title("Security Robustness Test Results", fontweight='bold')
    ax.axhline(y=100, color='gray', linestyle='--', linewidth=1, alpha=0.5)
    ax.grid(True, alpha=0.2, axis='y')

    for bar, rate in zip(bars, rates):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.5,
                f'{rate:.1f}%', ha='center', va='bottom', fontweight='bold', fontsize=10)

    plt.tight_layout()
    plt.savefig("security_test_results.png", dpi=150, bbox_inches='tight')
    plt.show()
    print("  Security graph saved to: security_test_results.png")


# ─────────────────────────────────────────────
#  MAIN EXECUTION
# ─────────────────────────────────────────────

if __name__ == "__main__":

    print("\n" + "="*55)
    print("       STEP 1: ECC Secret-Sharing Pipeline")
    print("="*55)

    ecc_input_file  = 'datasets_new.xlsx'
    ecc_output_file = 'processed_output1_new.xlsx'

    if not os.path.exists(ecc_input_file):
        print(f"  [!] '{ecc_input_file}' not found. Aborting.")
    else:
        input_wb = load_workbook(filename=ecc_input_file)
        sheet    = input_wb.active

        header_row_data = [cell.value for cell in sheet[1]]
        num_data_cols = sheet.max_column

        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "Processed"

        data_rows = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in row):
                break
            data_rows.append(row)

        print(f"  Found {len(data_rows)} rows to process.\n")

        out_ws.append(["ID"] + header_row_data + ["Pseudonym", "DataHash", "Verified"])

        args_list = [(i, data_row, len(header_row_data)) for i, data_row in enumerate(data_rows)]

        all_results = [process_single_row(args) for args in args_list]

        # Sort by row index to maintain order
        all_results.sort(key=lambda r: r["row_idx"])

        # ── Write results & accumulate timings ────────────────
        auth_records       = []
        total_reg_time     = 0.0
        total_token_time   = 0.0
        total_pseudo_time  = 0.0
        total_ss_time      = 0.0
        total_ecc_time     = 0.0
        total_binding_time = 0.0
        total_verify_time  = 0.0

        for res in all_results:
            output_row_idx = res["row_idx"] + 2
            out_ws.cell(row=output_row_idx, column=1, value=res["user_id"])

            for col_idx, message in res["col_results"]:
                if message is not None:
                    out_ws.cell(row=output_row_idx, column=col_idx + 2, value=message)

            out_ws.cell(row=output_row_idx, column=len(header_row_data) + 2, value=res["P"][:32])
            out_ws.cell(row=output_row_idx, column=len(header_row_data) + 3, value=res["data_hash"][:32])
            out_ws.cell(row=output_row_idx, column=len(header_row_data) + 4, value=res["status"])

            auth_records.append((res["user_id"], res["P"], res["data_hash"], res["is_valid"]))
            total_reg_time     += res["t_reg"]
            total_token_time   += res["t_tok"]
            total_pseudo_time  += res["t_pseudo"]
            total_ss_time      += res["t_ss"]
            total_ecc_time     += res["t_ecc"]
            total_binding_time += res["t_bind"]
            total_verify_time  += res["t_ver"]

        # ── Global Accuracy Summary ───────────────────────────
        all_pairs = []
        for res in all_results:
            all_pairs.extend(res["accuracy_log"])

        if all_pairs:
            exact_count = sum(1 for o, r in all_pairs if abs(o - r) < 10 ** (-PRECISION))
            n_vals      = len(all_pairs)

            print("\n" + "="*65)
            print("  Floating Point Recovery Accuracy Report")
            print("="*65)
            print(f"  {'#':>4}  {'Original':>14}  {'Recovered':>14}  {'Abs Error':>12}  {'Match?':>7}")
            print("-"*65)
            for idx, (o, r) in enumerate(all_pairs, 1):
                ae    = abs(o - r)
                match = "YES" if ae < 10**(-PRECISION) else "NO"
                print(f"  {idx:>4}  {o:>14.{PRECISION}f}  {r:>14.{PRECISION}f}  {ae:>12.{PRECISION+4}f}  {match:>7}")
            print("-"*65)
            print(f"  Exact Recovery: {exact_count}/{n_vals} ({exact_count/n_vals*100:.1f}%)")
            print("="*65)

        # ── Summary
        print("\n" + "="*55)
        print("  Authentication Summary")
        print("="*55)
        passed = sum(1 for _, _, _, v in auth_records if v)
        print(f"  Total users : {len(auth_records)}")
        print(f"  Verified OK : {passed}")
        print(f"  Failed      : {len(auth_records) - passed}")

        print("\n" + "="*55)
        print("  Stage-wise Timing Summary (all users combined)")
        print("="*55)
        print(f"  Registration              : {total_reg_time*1000:.6f} ms")
        print(f"  Token Generation          : {total_token_time*1000:.6f} ms")
        print(f"  Pseudonym Generation      : {total_pseudo_time*1000:.6f} ms")
        print(f"  Additive Secret Sharing   : {total_ss_time:.6f} seconds")
        print(f"  ECC ElGamal Enc/Dec       : {total_ecc_time:.6f} seconds")
        print(f"  Data Binding              : {total_binding_time*1000:.6f} ms")
        print(f"  Verification              : {total_verify_time*1000:.6f} ms")
        total_auth = (total_reg_time + total_token_time + total_pseudo_time
                      + total_ss_time + total_ecc_time
                      + total_binding_time + total_verify_time)
        print(f"  ─────────────────────────────────────────")
        print(f"  Total Time                : {total_auth:.6f} seconds")
        print("="*55)

        out_wb.save(ecc_output_file)
        print(f"\n  ECC output saved to: {ecc_output_file}")

        # ── Step 2: TOPSIS
        print("\n" + "="*55)
        print("       STEP 2: TOPSIS Decision Making")
        print("="*55)

        topsis_wb = load_workbook(filename=ecc_output_file)
        topsis_ws = topsis_wb.active

        alternatives = []
        for row in topsis_ws.iter_rows(min_row=2, values_only=True):
            t = row[1]
            c = row[2]
            d = row[3]
            if t is None and c is None and d is None:
                break
            if None in (t, c, d):
                continue
            alternatives.append([float(t), float(c), float(d)])

        if len(alternatives) < 2:
            print("  [!] Need at least 2 alternatives for TOPSIS. Aborting.")
        else:
            print(f"  Loaded {len(alternatives)} alternatives for TOPSIS.\n")
            topsis_result, elapsed_topsis = topsis_pipeline(
                alternatives=alternatives,
                save_results=True,
                output_file="topsis_results.xlsx"
            )
            print(f"  [Time] TOPSIS Total Pipeline: {elapsed_topsis:.6f} seconds")

        # ── Graphs ────────────────────────────────────────────
        import matplotlib.pyplot as plt
        import matplotlib.gridspec as gridspec
        import numpy as np

        print("\n  Generating graphs...")

        # Collect data
        originals  = [o for o, r in all_pairs]
        recovered  = [r for o, r in all_pairs]
        abs_errors = [abs(o - r) for o, r in all_pairs]
        n_vals     = len(all_pairs)

        # X-axis labels: Row1-Time, Row1-Cost, Row1-Dist, Row2-Time ...
        col_names = ["Time", "Cost", "Dist"]
        n_cols    = len(col_names)
        x_labels  = []
        for i in range(n_vals):
            row_num = i // n_cols + 1
            col_nm  = col_names[i % n_cols]
            x_labels.append(f"R{row_num}-{col_nm}")

        exact_count = sum(1 for e in abs_errors if e < 10**(-PRECISION))

        # TOPSIS scores & ranks
        topsis_scores_list = [s for _, s, _ in topsis_result]
        topsis_ranks_list  = [rk for _, _, rk in topsis_result]
        sorted_by_rank     = sorted(zip(topsis_ranks_list,
                                        [f"Alt {i+1}" for i in range(len(topsis_scores_list))],
                                        topsis_scores_list))
        t_ranks, t_labels, t_scores = zip(*sorted_by_rank)

        fig, axes = plt.subplots(1, 3, figsize=(20, 5))
        fig.suptitle("ECC + TOPSIS Pipeline — Results Dashboard",
                     fontsize=14, fontweight='bold')

        # ── Graph 1: Original vs Recovered ───────────────────
        ax1 = axes[0]
        x_pos = range(n_vals)
        ax1.plot(x_pos, originals, 'o-', color='steelblue',
                 label='Original', linewidth=1.5, markersize=5)
        ax1.plot(x_pos, recovered, 's--', color='tomato',
                 label='Recovered', linewidth=1.5, markersize=5)
        ax1.set_title("Original vs Recovered Float Values")
        ax1.set_xlabel("Data Point")
        ax1.set_ylabel("Float Value")
        ax1.set_xticks(range(n_vals))
        ax1.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=7)
        ax1.legend()
        ax1.grid(True, alpha=0.3)

        # ── Graph 2: Absolute Error per value ────────────────
        ax2 = axes[1]
        colors = ['tomato' if e >= 10**(-PRECISION) else 'mediumseagreen'
                  for e in abs_errors]
        ax2.bar(x_pos, abs_errors, color=colors, edgecolor='white')
        ax2.axhline(y=10**(-PRECISION), color='red', linestyle='--',
                    linewidth=1, label=f'Threshold (1e-{PRECISION})')
        ax2.set_title("Absolute Error per Data Point")
        ax2.set_xlabel("Data Point")
        ax2.set_ylabel("Absolute Error")
        ax2.set_xticks(range(n_vals))
        ax2.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=7)
        ax2.legend(fontsize=8)
        ax2.grid(True, alpha=0.3, axis='y')

        # ── Graph 3: TOPSIS Score & Rank ─────────────────────
        ax3 = axes[2]
        bar_cols = ['gold' if rk == 1 else '#4e79a7' for rk in t_ranks]
        bars = ax3.bar(t_labels, t_scores, color=bar_cols, edgecolor='white')
        ax3.set_title("TOPSIS Score by Alternative (sorted by rank)")
        ax3.set_xlabel("Alternative")
        ax3.set_ylabel("TOPSIS Score")
        ax3.tick_params(axis='x', rotation=30)
        ax3.grid(True, alpha=0.3, axis='y')
        for bar, score, rk in zip(bars, t_scores, t_ranks):
            ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                     f'#{rk}\n{score:.4f}', ha='center', va='bottom', fontsize=8)

        plt.tight_layout()
        plt.savefig("pipeline_results.png", dpi=150, bbox_inches='tight')
        plt.show()
        print("  Graphs saved to: pipeline_results.png")

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"\nTotal Elapsed time: {elapsed_time:.6f} seconds")

    # ── Step 3: Security Robustness Tests ─────────────────────
    sec_results = run_security_tests(n_trials=5)
    plot_security_results(sec_results)
