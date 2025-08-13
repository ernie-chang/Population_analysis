from flask import Flask, render_template, request
import os

app = Flask(__name__)


def get_regions_and_files(charts_dir: str):
    if not os.path.isdir(charts_dir):
        return [], []
    files = [f for f in os.listdir(charts_dir) if f.endswith('_attendance.png')]
    regions = sorted(list({f.split('_')[0] for f in files}))
    for extra in ['國中大區', '青年大區']:
        if extra not in regions:
            regions.append(extra)
    return regions, files


def find_subdistricts_for_region(charts_dir: str, region: str):
    subdistricts = []
    if not os.path.isdir(charts_dir):
        return subdistricts
    suffix = '_attendance.png'
    prefix = f'{region}_'
    for f in os.listdir(charts_dir):
        if not f.endswith(suffix):
            continue
        if not f.startswith(prefix):
            continue
        if f == f'{region}{suffix}':
            continue
        sub = f[len(prefix):-len(suffix)]
        if sub and sub not in subdistricts:
            subdistricts.append(sub)
    return sorted(subdistricts)


def build_subdistrict_cards(charts_dir: str, region: str, subdistricts: list):
    cards = []
    for s in subdistricts:
        att = f'charts/{region}_{s}_attendance.png'
        bur = f'charts/{region}_{s}_burden.png'
        att_path = os.path.join(charts_dir, f'{region}_{s}_attendance.png')
        bur_path = os.path.join(charts_dir, f'{region}_{s}_burden.png')
        cards.append({
            'name': s,
            'attendance_chart': att,
            'burden_chart': bur,
            'has_attendance': os.path.exists(att_path),
            'has_burden': os.path.exists(bur_path),
        })
    return cards


@app.route('/')
def index():
    charts_dir = os.path.join(app.static_folder, 'charts')
    regions, _ = get_regions_and_files(charts_dir)
    # Prefer '總計' as default region if present
    requested_region = request.args.get('region', '').strip()
    if requested_region:
        region = requested_region
    else:
        region = '總計' if '總計' in regions else (regions[0] if regions else '')

    subdistricts = find_subdistricts_for_region(charts_dir, region)
    subdistrict_cards = build_subdistrict_cards(charts_dir, region, subdistricts)

    # Always show region-level charts on main page
    attendance_chart = f'charts/{region}_attendance.png'
    burden_chart = f'charts/{region}_burden.png'

    return render_template(
        'index.html',
        regions=regions,
        region=region,
        subdistricts=subdistricts,
        subdistrict_cards=subdistrict_cards,
        attendance_chart=attendance_chart,
        burden_chart=burden_chart,
    )


if __name__ == '__main__':
    app.run(debug=True)
