# workout-tracker

Tooling for building and analysing hypertrophy programmes, modelled on Jeff
Nippard's tracker layout and designed to be usable in Google Sheets on a phone.

A programme is a YAML file. The generator turns it into a tracker; the parser
turns filled-in trackers back into data; the plotter shows what actually
happened.

```
programs/     one YAML per programme
src/          build, extract, plot
docs/         the rules, and the gym's physical constraints
```

## Usage

**Build a tracker**

```bash
python src/build_tracker.py programs/arm-shoulder.yaml "Arm & Shoulder v7.xlsx"

# returning from a layoff — run 2 sets for the first 4 weeks
python src/build_tracker.py programs/arm-shoulder.yaml out.xlsx --taper-weeks 4
```

Import the result into Google Sheets with File → Import → Upload. Effort
dropdowns, colour coding, frozen panes and numeric keypad formats all survive.

**Pull the data back out**

```bash
python src/extract_history.py history.csv "Arm & Shoulder v7.xlsx"

# older files whose week headers say "WEEK 3" rather than a date
python src/extract_history.py history.csv old.xlsx --anchor 2023-08-14
```

**Plot it**

```bash
python src/plot_progress.py history.csv progress.png
```

## Writing a new programme

Copy `programs/arm-shoulder.yaml` and edit. Each session lists blocks, each
block is either a `superset` (two exercises) or a `single`.

```yaml
- type: superset
  label: SUPERSET A
  note: one cable station, rope stays on
  exercises:
    - {code: A1, name: Cable Rope Pushdown, sets: 3, reps: "10-15", start: 27.5}
    - {code: A2, name: Cable Rope Hammer Curl, sets: 3, reps: "10-14", start: 25}

- {type: single, name: Rear Delt Machine, sets: 3, reps: "12-15", start: 15}
```

`start` pre-fills the week-1 weight column. Use a string for paired movements
(`"25,15"` for wrist curl + reverse curl) or drop sets (`"11.25,10,7.5"`).

Set `alternate_supersets: true` on a session to print an A FIRST / B first
banner over each week, so the same exercise doesn't always run pre-fatigued.

### Before you write one, read

- **[docs/progression-rules.md](docs/progression-rules.md)** — how to choose rep
  ranges, when to add load, what a real stall looks like. In particular: rep
  range has to match the equipment's increment size, or the programme will
  quietly stall no matter how hard you train.
- **[docs/gym-constraints.md](docs/gym-constraints.md)** — rooms, what pairs
  with what, and which weights physically exist.

## Provenance

The layout is derived from Jeff Nippard's Intermediate Workout Tracker. The
rules in `docs/` come from four years of personal logs (2023–2026): a beginner
full-body block, a chest specialisation, an upper/lower split, a shoulder
specialisation, and the current arm/shoulder programme.
