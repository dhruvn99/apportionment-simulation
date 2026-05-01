# ============================================================
# APPORTIONMENT SIMULATION
# ============================================================

import openpyxl
import math
import random
import csv
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as pltn
from collections import defaultdict

# ── Reproducibility ──────────────────────────────────────────
RANDOM_SEED = 42
random.seed(RANDOM_SEED)
np.random.seed(RANDOM_SEED)

# ── Global Configuration ─────────────────────────────────────
DATA_PATH      = '/Users/dhruv/Downloads/apportionment_sims/census_data.xlsx'
OUTPUT_DIR     = '/Users/dhruv/Downloads/apportionment_sims/'
TARGET_YEARS   = [1990, 2000, 2010, 2020]
H_CENSUS       = 435
TOTAL_POP      = 331_449_281
H_PER_STATE    = 435 / 50
SYNTH_N_STATES = [50, 100, 150, 200]
SIGMA_VALUES   = {
    'Low (s=0.80)':        0.80,
    'Calibrated (s=1.18)': 1.18,
    'High (s=1.60)':       1.60,
}
N_DRAWS = 5000


# ============================================================
# STEP 1 — LOAD CENSUS DATA
# ============================================================

def load_census_data(filepath):
    """
    Reads the census Excel file and returns a dictionary:
    { year: { state: population } }
    Only includes states with a valid population for that year.
    """
    wb   = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws   = wb['Population']
    rows = list(ws.iter_rows(values_only=True))

    header = rows[0]
    years  = [int(y) for y in header[1:] if y is not None]

    census = defaultdict(dict)
    for row in rows[1:]:
        state = row[0]
        if state is None:
            continue
        for i, year in enumerate(years):
            val = row[i + 1]
            if val is not None and val > 0:
                census[year][state] = int(val)

    wb.close()
    return census


# ============================================================
# STEP 2 — APPORTIONMENT METHODS
# ============================================================

def priority_method(populations, H, priority_fn):
    """
    Generic priority queue method.
    Assigns seats one at a time to the state with the
    highest priority value at each step.
    priority_fn(population, seats_already_held) -> priority
    """
    import heapq
    states = list(populations.keys())
    seats  = {s: 0 for s in states}
    heap   = [(-priority_fn(populations[s], 0), s) for s in states]
    heapq.heapify(heap)
    for _ in range(H):
        neg_pri, s = heapq.heappop(heap)
        seats[s]  += 1
        new_pri    = priority_fn(populations[s], seats[s])
        heapq.heappush(heap, (-new_pri, s))
    return seats


def jefferson(populations, H):
    # Floor rounding. Priority = p/n. Large-state bias.
    return priority_method(
        populations, H,
        lambda p, n: p / (n if n > 0 else 1)
    )


def adams(populations, H):
    # Ceiling rounding. Priority = p/(n+1). Small-state bias.
    return priority_method(
        populations, H,
        lambda p, n: p / (n + 1)
    )


def webster(populations, H):
    # Arithmetic mean rounding. Priority = p/(2n+1). No systematic bias.
    return priority_method(
        populations, H,
        lambda p, n: p / (2 * n + 1)
    )


def hill_huntington(populations, H):
    # Geometric mean rounding. Slight small-state bias.
    def pri(p, n):
        return float('inf') if n == 0 else p / math.sqrt(n * (n + 1))
    return priority_method(populations, H, pri)


def dean(populations, H):
    # Harmonic mean rounding. Small-state bias more than Hill.
    def pri(p, n):
        return float('inf') if n == 0 else p * (2 * n + 1) / (2 * n * (n + 1))
    return priority_method(populations, H, pri)


def hamilton(populations, H):
    # Largest remainder method. Always satisfies quota.
    total_pop  = sum(populations.values())
    quotas     = {s: p * H / total_pop for s, p in populations.items()}
    seats      = {s: int(q) for s, q in quotas.items()}
    remaining  = H - sum(seats.values())
    remainders = {s: quotas[s] - seats[s] for s in quotas}
    top        = sorted(remainders, key=lambda x: remainders[x], reverse=True)
    for s in top[:remaining]:
        seats[s] += 1
    return seats


METHODS = {
    'Jefferson':       jefferson,
    'Adams':           adams,
    'Webster':         webster,
    'Hill-Huntington': hill_huntington,
    'Dean':            dean,
    'Hamilton':        hamilton,
}

METHOD_ORDER = ['Jefferson', 'Adams', 'Webster', 'Hill-Huntington', 'Dean', 'Hamilton']


# ============================================================
# STEP 3 — VALIDATE METHODS AGAINST KNOWN 2020 RESULTS
# ============================================================

def validate_methods(census):
    """
    Runs all six methods on the 2020 census and prints
    seat allocations for key states.
    Known correct values for 2020:
        California = 52, Texas = 38, Wyoming = 1
    """
    pops = census[2020]
    print("\nVALIDATION — 2020 Census (H=435)")
    print(f"{'State':<20}", end="")
    for name in METHOD_ORDER:
        print(f"{name[:8]:>10}", end="")
    print()
    print("-" * (20 + 10 * len(METHOD_ORDER)))

    all_seats = {name: METHODS[name](pops, H_CENSUS) for name in METHOD_ORDER}

    check_states = ['California', 'Texas', 'Florida', 'New York', 'Wyoming']
    for state in check_states:
        print(f"{state:<20}", end="")
        for name in METHOD_ORDER:
            print(f"{all_seats[name][state]:>10}", end="")
        print()

    print("\nTotal seats check:")
    for name in METHOD_ORDER:
        total = sum(all_seats[name].values())
        print(f"  {name:<18} {total}")


# ============================================================
# STEP 4 — DEVIATION METRICS
# ============================================================

def compute_metrics(populations, seats, H):
    """
    Computes all deviation metrics for one apportionment.
    populations : dict {state: population}
    seats       : dict {state: seats_allocated}
    H           : total house size
    """
    total_pop   = sum(populations.values())
    states      = list(populations.keys())
    S           = len(states)
    quotas      = {s: populations[s] * H / total_pop for s in states}

    abs_devs    = [abs(seats[s] - quotas[s]) for s in states]
    rel_devs    = [abs(seats[s] - quotas[s]) / quotas[s] for s in states]
    seat_shares = [seats[s] / H for s in states]
    pop_shares  = [populations[s] / total_pop for s in states]

    return {
        'rel_mad':          sum(rel_devs) / S,
        'abs_mad':          sum(abs_devs) / S,
        'max_rel':          max(rel_devs),
        'max_abs':          max(abs_devs),
        'gallagher':        math.sqrt(
                                0.5 * sum(
                                    (seat_shares[i] - pop_shares[i]) ** 2
                                    for i in range(S)
                                )
                            ),
        'loosemore_hanby':  0.5 * sum(
                                abs(seat_shares[i] - pop_shares[i])
                                for i in range(S)
                            ),
        'quota_violations': sum(
                                1 for s in states
                                if seats[s] < math.floor(quotas[s])
                                or seats[s] > math.ceil(quotas[s])
                            ),
    }


def run_scenario(populations, H):
    """Applies all six methods and returns metrics for each."""
    return {
        name: compute_metrics(populations, fn(populations, H), H)
        for name, fn in METHODS.items()
    }


# ============================================================
# STEP 5 — CENSUS RUNS
# ============================================================

def run_census_scenarios(census):
    records = []
    for year in TARGET_YEARS:
        pops = {s: p for s, p in census[year].items() if p > 0}
        for method, metrics in run_scenario(pops, H_CENSUS).items():
            row = {
                'scenario':    f'Census {year}',
                'type':        'census',
                'year':        year,
                'n_states':    len(pops),
                'sigma_label': 'Actual Census',
                'H':           H_CENSUS,
                'method':      method,
            }
            row.update(metrics)
            records.append(row)
        print(f"  Census {year}: {len(pops)} states")
    return records


# ============================================================
# STEP 6 — SYNTHETIC POPULATION GENERATOR
# ============================================================

def generate_lognormal_populations(n_states, total_pop, sigma):
    """
    Generates n_states integer populations summing exactly to total_pop.
    Drawn from LogNormal(mu, sigma) where:
        mu = log(total_pop / n_states) - sigma^2 / 2
    """
    mu       = math.log(total_pop / n_states) - (sigma ** 2) / 2
    raw      = [math.exp(random.gauss(mu, sigma)) for _ in range(n_states)]
    scale    = total_pop / sum(raw)
    pops_int = [max(1, round(r * scale)) for r in raw]
    pops_int[-1] += total_pop - sum(pops_int)
    pops_int[-1]  = max(1, pops_int[-1])
    return {f'State_{i+1:03d}': p for i, p in enumerate(pops_int)}


# ============================================================
# STEP 7 — MONTE CARLO RUNS
# ============================================================

def run_monte_carlo(n_states, sigma, sigma_label):
    H       = round(H_PER_STATE * n_states)
    label   = f'Synthetic_{n_states}states'
    records = []

    for draw in range(N_DRAWS):
        pops = generate_lognormal_populations(n_states, TOTAL_POP, sigma)
        for method, metrics in run_scenario(pops, H).items():
            row = {
                'scenario':    label,
                'type':        'synthetic',
                'n_states':    n_states,
                'sigma':       sigma,
                'sigma_label': sigma_label,
                'H':           H,
                'draw':        draw,
                'method':      method,
            }
            row.update(metrics)
            records.append(row)

    print(f"  {N_DRAWS} draws | n_states={n_states} | H={H} | {sigma_label}")
    return records


# ============================================================
# STEP 8 — SUMMARIZE RESULTS
# ============================================================

METRIC_LABELS = {
    'rel_mad':          'Relative MAD',
    'abs_mad':          'Absolute MAD',
    'max_rel':          'Max Relative Dev',
    'max_abs':          'Max Absolute Dev',
    'gallagher':        'Gallagher Index',
    'loosemore_hanby':  'Loosemore-Hanby',
    'quota_violations': 'Quota Violations',
}


def summarize(records, group_keys):
    metrics = list(METRIC_LABELS.keys())
    groups  = defaultdict(list)
    for r in records:
        key = tuple(r[k] for k in group_keys)
        groups[key].append(r)
    summaries = []
    for key, rows in groups.items():
        s = dict(zip(group_keys, key))
        for m in metrics:
            vals           = [r[m] for r in rows]
            s[f'{m}_mean'] = float(np.mean(vals))
            s[f'{m}_std']  = float(np.std(vals, ddof=1)) if len(vals) > 1 else 0.0
        s['n_draws'] = len(rows)
        summaries.append(s)
    return summaries


# ============================================================
# STEP 9 — PRINT RESULTS TABLE
# ============================================================

def print_table(summaries, scenario_label, sigma_label=None):
    rows = [s for s in summaries if s.get('scenario') == scenario_label]
    if sigma_label:
        rows = [s for s in rows if s.get('sigma_label') == sigma_label]
    if not rows:
        return
    metrics = ['rel_mad', 'abs_mad', 'max_rel', 'gallagher',
               'loosemore_hanby', 'quota_violations']
    tag     = scenario_label + (f" | {sigma_label}" if sigma_label else "")
    print(f"\n{'='*85}")
    print(f"  {tag}")
    print(f"{'='*85}")
    header = f"{'Method':<18}" + "".join(
        f"{METRIC_LABELS[m][:12]:>13}" for m in metrics
    )
    print(header)
    print("-" * len(header))
    for method in METHOD_ORDER:
        row  = next((s for s in rows if s['method'] == method), None)
        if not row:
            continue
        line = f"{method:<18}"
        for m in metrics:
            v = row.get(f'{m}_mean', row.get(m, float('nan')))
            line += f"{v:>13.5f}"
        print(line)


# ============================================================
# STEP 10 — PLOTS
# ============================================================

METHOD_COLORS = {
    'Jefferson':       '#E63946',
    'Adams':           '#F4A261',
    'Webster':         '#2A9D8F',
    'Hill-Huntington': '#457B9D',
    'Dean':            '#6A4C93',
    'Hamilton':        '#264653',
}

SIGMA_COLORS = {
    'Low (s=0.80)':        '#74B3CE',
    'Calibrated (s=1.18)': '#2A9D8F',
    'High (s=1.60)':       '#E76F51',
}

plt.rcParams.update({
    'font.family':       'serif',
    'axes.spines.top':   False,
    'axes.spines.right': False,
    'axes.grid':         True,
    'grid.alpha':        0.3,
    'grid.linestyle':    '--',
    'figure.dpi':        150,
})


def plot_census_bars(census_summaries, metric, outpath):
    years   = TARGET_YEARS
    x       = np.arange(len(years))
    width   = 0.13
    fig, ax = plt.subplots(figsize=(11, 5))
    for i, method in enumerate(METHOD_ORDER):
        vals = [
            next((s[f'{metric}_mean'] for s in census_summaries
                  if s.get('year') == y and s['method'] == method), 0)
            for y in years
        ]
        ax.bar(
            x + (i - len(METHOD_ORDER) / 2 + 0.5) * width,
            vals, width,
            label=method,
            color=METHOD_COLORS[method],
            alpha=0.88, zorder=3
        )
    ax.set_xticks(x)
    ax.set_xticklabels([str(y) for y in years], fontsize=11)
    ax.set_xlabel('Census Year', fontsize=12)
    ax.set_ylabel(METRIC_LABELS[metric], fontsize=12)
    ax.set_title(
        f'{METRIC_LABELS[metric]} by Method — Historical Census Data',
        fontsize=13, fontweight='bold'
    )
    ax.legend(fontsize=9, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(outpath, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {outpath}")


def plot_growth_trend(mc_summaries, metric, outpath):
    sigma_labels = list(SIGMA_VALUES.keys())
    fig, axes    = plt.subplots(1, 3, figsize=(15, 5), sharey=True)
    for ax, sigma_label in zip(axes, sigma_labels):
        for method in METHOD_ORDER:
            vals, errs = [], []
            for n in SYNTH_N_STATES:
                row = next(
                    (s for s in mc_summaries
                     if s.get('n_states') == n
                     and s.get('sigma_label') == sigma_label
                     and s['method'] == method),
                    None
                )
                vals.append(row[f'{metric}_mean'] if row else np.nan)
                errs.append(row[f'{metric}_std']  if row else 0.0)
            ax.plot(
                SYNTH_N_STATES, vals,
                marker='o', linewidth=2,
                color=METHOD_COLORS[method],
                label=method, zorder=3
            )
            ax.fill_between(
                SYNTH_N_STATES,
                [v - e for v, e in zip(vals, errs)],
                [v + e for v, e in zip(vals, errs)],
                alpha=0.10, color=METHOD_COLORS[method]
            )
        ax.set_xticks(SYNTH_N_STATES)
        ax.set_xlabel('Number of States', fontsize=11)
        ax.set_title(
            sigma_label, fontsize=11, fontweight='bold',
            color=SIGMA_COLORS[sigma_label]
        )
    axes[0].set_ylabel(METRIC_LABELS[metric], fontsize=12)
    axes[-1].legend(fontsize=8, framealpha=0.9, loc='upper left')
    fig.suptitle(
        f'{METRIC_LABELS[metric]} vs. State Count — Sensitivity to Population Skew\n'
        f'({N_DRAWS:,} Monte Carlo draws per scenario)',
        fontsize=13, fontweight='bold'
    )
    fig.tight_layout()
    fig.savefig(outpath, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {outpath}")


def plot_sigma_sensitivity(mc_summaries, metric, n_states, outpath):
    sigma_labels = list(SIGMA_VALUES.keys())
    x     = np.arange(len(METHOD_ORDER))
    width = 0.25
    fig, ax = plt.subplots(figsize=(11, 5))
    for i, sigma_label in enumerate(sigma_labels):
        vals = []
        for method in METHOD_ORDER:
            row = next(
                (s for s in mc_summaries
                 if s.get('n_states') == n_states
                 and s.get('sigma_label') == sigma_label
                 and s['method'] == method),
                None
            )
            vals.append(row[f'{metric}_mean'] if row else np.nan)
        ax.bar(
            x + (i - 1) * width, vals, width,
            label=sigma_label,
            color=SIGMA_COLORS[sigma_label],
            alpha=0.88, zorder=3
        )
    ax.set_xticks(x)
    ax.set_xticklabels(METHOD_ORDER, rotation=20, ha='right', fontsize=10)
    ax.set_ylabel(METRIC_LABELS[metric], fontsize=12)
    ax.set_title(
        f'{METRIC_LABELS[metric]} by Method and Population Skew — {n_states} States\n'
        f'({N_DRAWS:,} Monte Carlo draws per scenario)',
        fontsize=12, fontweight='bold'
    )
    ax.legend(fontsize=9, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(outpath, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {outpath}")


def plot_ranking_heatmap(all_summaries, metric, sigma_label, outpath):
    census_labels = [f'Census {y}' for y in TARGET_YEARS]
    synth_labels  = [f'Synthetic_{n}states' for n in SYNTH_N_STATES]
    all_labels    = census_labels + synth_labels
    disp_labels   = [str(y) for y in TARGET_YEARS] + \
                    [f'{n} St.' for n in SYNTH_N_STATES]
    rank_matrix = np.zeros((len(METHOD_ORDER), len(all_labels)))
    for col_i, label in enumerate(all_labels):
        if 'Census' in label:
            sc_rows = [s for s in all_summaries if s.get('scenario') == label]
        else:
            sc_rows = [s for s in all_summaries
                       if s.get('scenario') == label
                       and s.get('sigma_label') == sigma_label]
        sorted_rows = sorted(sc_rows, key=lambda r: r[f'{metric}_mean'])
        for row_i, method in enumerate(METHOD_ORDER):
            rank_matrix[row_i, col_i] = next(
                (i + 1 for i, r in enumerate(sorted_rows)
                 if r['method'] == method), 6
            )
    fig, ax = plt.subplots(figsize=(11, 4.5))
    im = ax.imshow(rank_matrix, cmap='RdYlGn_r', vmin=1, vmax=6, aspect='auto')
    ax.set_xticks(range(len(disp_labels)))
    ax.set_xticklabels(disp_labels, fontsize=10)
    ax.set_yticks(range(len(METHOD_ORDER)))
    ax.set_yticklabels(METHOD_ORDER, fontsize=10)
    for i in range(len(METHOD_ORDER)):
        for j in range(len(all_labels)):
            ax.text(
                j, i, str(int(rank_matrix[i, j])),
                ha='center', va='center',
                fontsize=12, fontweight='bold',
                color='white' if rank_matrix[i, j] >= 4 else 'black'
            )
    ax.axvline(x=3.5, color='white', linewidth=2.5)
    ax.set_title(
        f'Method Rankings — {METRIC_LABELS[metric]}\n'
        f'{sigma_label}  |  1 = lowest deviation',
        fontsize=12, fontweight='bold'
    )
    fig.colorbar(im, ax=ax, shrink=0.8).set_label('Rank', fontsize=9)
    fig.tight_layout()
    fig.savefig(outpath, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {outpath}")


def plot_violin(mc_records, metric, n_states, outpath):
    sigma_labels = list(SIGMA_VALUES.keys())
    fig, axes    = plt.subplots(1, 3, figsize=(15, 5), sharey=True)
    for ax, sigma_label in zip(axes, sigma_labels):
        draws = defaultdict(list)
        for r in mc_records:
            if r['n_states'] == n_states and r.get('sigma_label') == sigma_label:
                draws[r['method']].append(r[metric])
        data   = [draws[m] for m in METHOD_ORDER]
        colors = [METHOD_COLORS[m] for m in METHOD_ORDER]
        parts  = ax.violinplot(data, positions=range(len(METHOD_ORDER)),
                               showmedians=True, showextrema=False)
        for pc, color in zip(parts['bodies'], colors):
            pc.set_facecolor(color)
            pc.set_alpha(0.75)
        parts['cmedians'].set_color('black')
        parts['cmedians'].set_linewidth(1.8)
        ax.set_xticks(range(len(METHOD_ORDER)))
        ax.set_xticklabels(METHOD_ORDER, rotation=25, ha='right', fontsize=9)
        ax.set_title(
            sigma_label, fontsize=11, fontweight='bold',
            color=SIGMA_COLORS[sigma_label]
        )
    axes[0].set_ylabel(METRIC_LABELS[metric], fontsize=12)
    fig.suptitle(
        f'Distribution of {METRIC_LABELS[metric]} — {n_states} States\n'
        f'({N_DRAWS:,} Monte Carlo draws per scenario)',
        fontsize=13, fontweight='bold'
    )
    fig.tight_layout()
    fig.savefig(outpath, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {outpath}")


def plot_quota_violations(mc_summaries, outpath):
    sigma_cal       = 'Calibrated (s=1.18)'
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5))
    for method in METHOD_ORDER:
        vals = [
            next((s['quota_violations_mean'] for s in mc_summaries
                  if s.get('n_states') == n
                  and s.get('sigma_label') == sigma_cal
                  and s['method'] == method), np.nan)
            for n in SYNTH_N_STATES
        ]
        ax1.plot(
            SYNTH_N_STATES, vals,
            marker='o', linewidth=2,
            color=METHOD_COLORS[method], label=method
        )
    ax1.set_xticks(SYNTH_N_STATES)
    ax1.set_xlabel('Number of States', fontsize=11)
    ax1.set_ylabel('Mean Quota Violations per Draw', fontsize=11)
    ax1.set_title(
        f'Quota Violations vs. State Count\n({sigma_cal})',
        fontsize=11, fontweight='bold'
    )
    ax1.legend(fontsize=8, framealpha=0.9)
    n_states     = 50
    sigma_labels = list(SIGMA_VALUES.keys())
    x     = np.arange(len(METHOD_ORDER))
    width = 0.25
    for i, sigma_label in enumerate(sigma_labels):
        vals = [
            next((s['quota_violations_mean'] for s in mc_summaries
                  if s.get('n_states') == n_states
                  and s.get('sigma_label') == sigma_label
                  and s['method'] == method), np.nan)
            for method in METHOD_ORDER
        ]
        ax2.bar(
            x + (i - 1) * width, vals, width,
            label=sigma_label,
            color=SIGMA_COLORS[sigma_label],
            alpha=0.88, zorder=3
        )
    ax2.set_xticks(x)
    ax2.set_xticklabels(METHOD_ORDER, rotation=20, ha='right', fontsize=9)
    ax2.set_ylabel('Mean Quota Violations per Draw', fontsize=11)
    ax2.set_title(
        f'Quota Violations vs. Population Skew\n({n_states} States)',
        fontsize=11, fontweight='bold'
    )
    ax2.legend(fontsize=8, framealpha=0.9)
    fig.suptitle(
        'Quota Violations: State Count and Population Skew Effects',
        fontsize=13, fontweight='bold'
    )
    fig.tight_layout()
    fig.savefig(outpath, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {outpath}")


# ============================================================
# STEP 11 — SAVE CSV
# ============================================================

def save_csv(records, filepath):
    if not records:
        return
    with open(filepath, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)
    print(f"  Saved: {filepath}")


# ============================================================
# STEP 12 — MAIN
# ============================================================

def main():
    print("\n" + "=" * 70)
    print("  APPORTIONMENT MONTE CARLO SIMULATION")
    print("=" * 70)

    # ── Load data ─────────────────────────────────────────────
    print("\n[1] Loading census data...")
    census = load_census_data(DATA_PATH)
    print(f"    Years available: {sorted(census.keys())}")

    # ── Validate methods ──────────────────────────────────────
    print("\n[2] Validating methods against 2020 census...")
    validate_methods(census)

    # ── Census runs ───────────────────────────────────────────
    print("\n[3] Running census scenarios...")
    census_records   = run_census_scenarios(census)
    census_summaries = summarize(
        census_records,
        ['scenario', 'year', 'n_states', 'sigma_label', 'H', 'method']
    )

    # ── Monte Carlo ───────────────────────────────────────────
    print("\n[4] Running Monte Carlo synthetic scenarios...")
    all_mc_records = []
    for sigma_label, sigma in SIGMA_VALUES.items():
        print(f"\n  --- {sigma_label} ---")
        for n_states in SYNTH_N_STATES:
            all_mc_records += run_monte_carlo(n_states, sigma, sigma_label)

    mc_summaries  = summarize(
        all_mc_records,
        ['scenario', 'n_states', 'sigma_label', 'H', 'method']
    )
    all_summaries = census_summaries + mc_summaries

    # ── Print tables ──────────────────────────────────────────
    print("\n[5] Results...")
    for year in TARGET_YEARS:
        print_table(census_summaries, f'Census {year}')
    for sigma_label in SIGMA_VALUES:
        for n in SYNTH_N_STATES:
            print_table(mc_summaries, f'Synthetic_{n}states', sigma_label)

    # ── Ranking stability ─────────────────────────────────────
    print(f"\n{'='*85}")
    print("  RANKING STABILITY — Relative MAD — Calibrated s=1.18")
    print(f"{'='*85}")
    scenarios = [f'Census {y}' for y in TARGET_YEARS] + \
                [f'Synthetic_{n}states' for n in SYNTH_N_STATES]
    short     = [str(y) for y in TARGET_YEARS] + \
                [f'{n}St' for n in SYNTH_N_STATES]
    sigma_cal = 'Calibrated (s=1.18)'
    print(f"  {'Method':<18}" + "".join(f"{s:>8}" for s in short))
    print("  " + "-" * (18 + 8 * len(scenarios)))
    for method in METHOD_ORDER:
        line = f"  {method:<18}"
        for sc in scenarios:
            if 'Census' in sc:
                sc_rows = [s for s in all_summaries
                           if s.get('scenario') == sc]
            else:
                sc_rows = [s for s in all_summaries
                           if s.get('scenario') == sc
                           and s.get('sigma_label') == sigma_cal]
            sorted_sc = sorted(sc_rows, key=lambda r: r['rel_mad_mean'])
            rank = next(
                (i + 1 for i, r in enumerate(sorted_sc)
                 if r['method'] == method), '-'
            )
            line += f"{rank:>8}"
        print(line)

    # ── Plots ─────────────────────────────────────────────────
    print("\n[6] Generating plots...")
    for metric in ['rel_mad', 'abs_mad', 'gallagher', 'max_rel']:
        plot_census_bars(
            census_summaries, metric,
            OUTPUT_DIR + f'census_bar_{metric}.png'
        )
    for metric in ['rel_mad', 'gallagher', 'max_rel', 'quota_violations']:
        plot_growth_trend(
            mc_summaries, metric,
            OUTPUT_DIR + f'trend_{metric}.png'
        )
    for n in SYNTH_N_STATES:
        plot_sigma_sensitivity(
            mc_summaries, 'rel_mad', n,
            OUTPUT_DIR + f'sigma_sensitivity_{n}states.png'
        )
    for sigma_label in SIGMA_VALUES:
        slug = sigma_label.replace('(', '').replace(')', '') \
                          .replace('=', '').replace(' ', '_')
        plot_ranking_heatmap(
            all_summaries, 'rel_mad', sigma_label,
            OUTPUT_DIR + f'heatmap_relmad_{slug}.png'
        )
        plot_ranking_heatmap(
            all_summaries, 'gallagher', sigma_label,
            OUTPUT_DIR + f'heatmap_gallagher_{slug}.png'
        )
    for n in [50, 200]:
        plot_violin(
            all_mc_records, 'rel_mad', n,
            OUTPUT_DIR + f'violin_relmad_{n}states.png'
        )
    plot_quota_violations(
        mc_summaries,
        OUTPUT_DIR + 'quota_violations.png'
    )

    # ── Save CSVs ─────────────────────────────────────────────
    print("\n[7] Saving CSVs...")
    save_csv(census_records,   OUTPUT_DIR + 'census_raw.csv')
    save_csv(census_summaries, OUTPUT_DIR + 'census_summary.csv')
    save_csv(mc_summaries,     OUTPUT_DIR + 'synthetic_summary.csv')

    print("\n  All done.\n" + "=" * 70)


# ── Entry point ───────────────────────────────────────────────
if __name__ == '__main__':
    main()