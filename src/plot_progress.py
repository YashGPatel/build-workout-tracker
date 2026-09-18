#!/usr/bin/env python3
"""Plot Yash's lifting history by muscle group, 2023-2026."""
import csv, datetime, collections
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

import sys
CSV = sys.argv[1] if len(sys.argv) > 1 else 'history.csv'
OUT = sys.argv[2] if len(sys.argv) > 2 else 'progress.png'

# merge name variants
ALIAS = {
    'Seated Dumbbell Shoulder Press': 'Seated DB Shoulder Press',
    'Paused Seated Dumbbell Shoulder Press': 'Seated DB Shoulder Press',
    'Lateral Pulldown': 'Lat Pulldown',
    'Seated Cable Row (mid/upper back)': 'Seated Cable Row',
    'Seated Leg Curls': 'Seated Leg Curl',
    'Leg Extensions': 'Seated Leg Extensions',
    'Incline Dumbbell Curls': 'Incline DB Curl (30-45°)',
    'Overhead Tricep  Extension': 'Overhead Tricep Extension',
    'Overhead Rope Extensions': 'Overhead Tricep Extension',
    'Lying Incline Lateral Raise': 'DB Lateral Raise',
    'Dumbbell Lateral Raise': 'DB Lateral Raise',
    'DB Chest Supported Row (mid/upper back)': 'Chest Supported DB Row',
    'Chest Supported DB Row (lat focus)': 'Chest Supported DB Row',
    'Standing Weighted Calf Raise': 'Weighted Calf Raise',
    'Seated Weighted Calf Raise': 'Weighted Calf Raise',
    'Wrist Curl + Rev. Curl (cable)': 'Wrist Curl + Reverse Curl',
    'Cable Tricep Pushdown': 'Cable Rope Pushdown',
    'Paused Flat Dumbbell Press': 'Flat Dumbbell Press',
    'Standing Mid-Chest Cable Fly': 'Cable Fly',
    'Standing High To Low Cable Fly': 'Cable Fly',
    'Seated Mid-Chest Cable Fly': 'Cable Fly',
    'Lying Leg Curls': 'Seated Leg Curl',
    'Barbell Romanian Deadlift': 'Romanian Deadlift',
    'Dumbbell Romanian Deadlift': 'Romanian Deadlift',
}

GROUP = {
    'Side delts': ['Cable Lateral Raise', 'DB Lateral Raise', 'LU Raises'],
    'Shoulders (press)': ['Seated DB Shoulder Press', 'Machine Shoulder Press',
                          'Standing Barbell Overhead Press'],
    'Rear delts': ['Rear Delt Cable Fly', 'Reverse Pec Deck', 'Standing Face Pulls'],
    'Biceps': ['Incline DB Curl (30-45°)', 'Standing Cable Curl', 'Bayesian Cable Curl',
               'Rope Cable Curls (neutral grip)', 'Cable Hammer Curl (rope)',
               'Reverse Cable Curl (straight bar)', 'Wrist Curl + Reverse Curl'],
    'Triceps': ['Overhead Tricep Extension', 'Cable Rope Pushdown', 'Close-Grip Bench Press'],
    'Chest': ['Barbell Bench Press', 'Flat Dumbbell Press', 'Low Incline Dumbbell Press',
              'Decline Dumbbell Press', 'Dumbbell Fly', 'Cable Fly',
              'Seated Flat Cable Press'],
    'Back': ['Lat Pulldown', 'Seated Cable Row', 'Barbell Row (mid/upper back)',
             'Chest Supported DB Row', 'Banded Pull-Ups', 'Pull-Ups'],
    'Legs': ['Barbell Back Squat', 'Barbell Deadlift', 'Romanian Deadlift',
             'Bulgarian Split Squat (quad focus)', 'Bulgarian Split Squat (glute focus)',
             'Seated Leg Extensions', 'Seated Leg Curl', 'Barbell Hip Thrust',
             'Weighted Calf Raise', 'Walking Lunges (quad-focus)', 'Reverse Lunges*',
             'Leg Press Calf Raise'],
}

PHASE_BANDS = [
    ('Beginner\nfull body',  '2023-05-01', '2023-08-10', '#dbe7f5'),
    ('Chest\n3-day',         '2023-08-11', '2023-11-05', '#f3dcef'),
    ('Upper/Lower\n4-day',   '2024-02-01', '2024-10-31', '#e5f0dc'),
    ('Shoulder\n3-day',      '2024-11-01', '2025-01-31', '#fdeacd'),
    ('Break',                '2025-02-01', '2026-04-30', '#f0f0f0'),
    ('Return',               '2026-05-01', '2026-08-15', '#f7dbdb'),
]

rows = []
for r in csv.DictReader(open(CSV)):
    d = datetime.date.fromisoformat(r['date'])
    if d.year < 2023:                       # one mistyped 2001 date, actually Aug 2023
        d = d.replace(year=2023)
    name = ALIAS.get(r['exercise'], r['exercise'])
    rows.append((d, name, float(r['weight'])))

# best (heaviest) working weight per exercise per session-date
best = collections.defaultdict(dict)
for d, name, w in rows:
    if name not in best or d not in best[name] or w > best[name][d]:
        best[name][d] = max(w, best[name].get(d, 0))

plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9,
                     'axes.grid': True, 'grid.alpha': .25, 'grid.linewidth': .6})
fig, axes = plt.subplots(4, 2, figsize=(15, 15), sharex=True)
fig.suptitle('Yash — lifting history by muscle group, 2023-2026',
             fontsize=17, fontweight='bold', y=.985)
fig.text(.5, .958, 'Heaviest working weight per session. Dumbbell lifts = weight per '
                   'dumbbell; barbell = total bar load; cables = stack setting.',
         ha='center', fontsize=9.5, color='#555')

xmin, xmax = datetime.date(2023, 4, 1), datetime.date(2026, 8, 20)

for ax, (group, exercises) in zip(axes.ravel(), GROUP.items()):
    for lo, hi, _, colr in [(p[1], p[2], p[0], p[3]) for p in PHASE_BANDS]:
        ax.axvspan(datetime.date.fromisoformat(lo), datetime.date.fromisoformat(hi),
                   color=colr, zorder=0)
    plotted = 0
    for ex in exercises:
        series = best.get(ex)
        if not series:
            continue
        pts = sorted(series.items())
        # break the line wherever there's a gap of more than 8 weeks, so the
        # year-long layoff isn't drawn as a smooth trend
        segments, seg = [], [pts[0]]
        for prev, cur in zip(pts, pts[1:]):
            if (cur[0] - prev[0]).days > 56:
                segments.append(seg); seg = []
            seg.append(cur)
        segments.append(seg)
        colr = None
        for i, s in enumerate(segments):
            line, = ax.plot([p[0] for p in s], [p[1] for p in s], marker='o',
                            ms=3.5, lw=1.6, color=colr,
                            label=ex if i == 0 else None)
            colr = line.get_color()
        plotted += 1
    ax.set_title(group, fontsize=12, fontweight='bold', loc='left')
    ax.set_ylabel('kg')
    ax.set_xlim(xmin, xmax)
    ax.margins(y=.22)
    if plotted:
        ax.legend(fontsize=7, loc='lower center', framealpha=.9,
                  ncol=2 if plotted > 4 else 1, borderpad=.4)

# phase labels along the top of the first row
for ax in axes[0]:
    for label, lo, hi, _c in [(p[0], p[1], p[2], p[3]) for p in PHASE_BANDS]:
        mid = (datetime.date.fromisoformat(lo) +
               (datetime.date.fromisoformat(hi) - datetime.date.fromisoformat(lo)) / 2)
        ax.annotate(label, xy=(mid, 1.02), xycoords=('data', 'axes fraction'),
                    ha='center', va='bottom', fontsize=7.5, color='#666')

for ax in axes[-1]:
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
    plt.setp(ax.get_xticklabels(), rotation=45, ha='right')

axes[-1][-1].axis('off') if len(GROUP) < axes.size else None
fig.tight_layout(rect=[0, 0, 1, .945])
out = OUT
fig.savefig(out, dpi=150, facecolor='white')
print('saved', out)

# --- console summary: first vs best vs latest -------------------------------
print(f"\n{'exercise':38s} {'first':>16s} {'peak':>14s} {'now (2026)':>14s}")
for group, exercises in GROUP.items():
    print('--', group)
    for ex in exercises:
        s = best.get(ex)
        if not s:
            continue
        pts = sorted(s.items())
        peak = max(pts, key=lambda p: p[1])
        recent = [p for p in pts if p[0].year == 2026]
        now = f'{recent[-1][1]:g} ({recent[-1][0]:%b %y})' if recent else '-'
        print(f'   {ex[:36]:36s} {pts[0][1]:6g} ({pts[0][0]:%b %y})  '
              f'{peak[1]:5g} ({peak[0]:%b %y})  {now:>16s}')
