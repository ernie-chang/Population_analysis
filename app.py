from flask import Flask, render_template, request, flash, redirect, url_for, jsonify
import os
from werkzeug.utils import secure_filename
import analysis
import pandas as pd

app = Flask(__name__)

# --- Configuration ---
# Folder to store uploaded report files
UPLOAD_FOLDER = 'date'
DATA_CACHE_PATH = os.path.join(UPLOAD_FOLDER, 'aggregated_data.pkl')
# Secret key for flashing messages (important for security)
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'a-default-secret-key-for-development-only')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

@app.context_processor
def inject_now():
    """Injects a timestamp into templates for cache busting image assets."""
    return {'now': pd.Timestamp.now().timestamp()}

def get_regions_and_files(charts_dir: str):
    if not os.path.isdir(charts_dir):
        return [], []
    files = [f for f in os.listdir(charts_dir) if f.endswith('_attendance.png')]
    detected = {f.split('_')[0] for f in files}

    # Always include extras
    detected.update({'國中大區', '青年大區'})

    regions = sorted(list(detected))

    # Move '總計' to the front if present; otherwise keep order
    if '總計' in regions:
        regions = ['總計'] + [r for r in regions if r != '總計']

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


def allowed_file(filename):
    """Checks if the file's extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def run_analysis():
    """
    Runs the full analysis pipeline: aggregates reports from the UPLOAD_FOLDER
    and generates all corresponding charts in the static/charts directory.
    Returns a status message.
    """
    reports_dir = app.config['UPLOAD_FOLDER']
    static_charts_dir = os.path.join(app.static_folder, 'charts')

    try:
        # 1. Aggregate data from all .xls* files
        df_reports = analysis.aggregate_reports(reports_dir)

        # Save aggregated data to cache for rate calculations
        if not df_reports.empty:
            df_reports.to_pickle(DATA_CACHE_PATH)
        elif os.path.exists(DATA_CACHE_PATH):
            # If no data, remove old cache
            os.remove(DATA_CACHE_PATH)

        # 2. Generate charts if data was found
        if not df_reports.empty:
            # Clear old charts to ensure a clean state
            if os.path.exists(static_charts_dir):
                for f in os.listdir(static_charts_dir):
                    if f.endswith('.png'):
                        os.remove(os.path.join(static_charts_dir, f))
            else:
                os.makedirs(static_charts_dir, exist_ok=True)

            # Generate '總計' (Total) chart
            analysis.generate_region_charts(df_reports, "總計", static_charts_dir)

            # Generate charts for each unique region found in the data
            unique_regions = df_reports["大區"].dropna().unique()
            for region in unique_regions:
                if str(region) != "總計":
                    analysis.generate_region_charts(df_reports, str(region), static_charts_dir)
            return f"分析完成，共處理 {len(df_reports['週末日'].unique())} 週的資料。"
        else:
            return "找不到可分析的報告檔案。"

    except Exception as e:
        app.logger.error(f"An error occurred during analysis: {e}", exc_info=True)
        return f"分析時發生錯誤: {e}"


@app.route('/')
def index():
    charts_dir = os.path.join(app.static_folder, 'charts')

    # Ensure directories for uploads and charts exist
    os.makedirs(charts_dir, exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    regions, _ = get_regions_and_files(charts_dir)
    requested_region = request.args.get('region', '').strip()
    if requested_region and requested_region in regions:
        region = requested_region
    else:
        region = '總計' if '總計' in regions else (regions[0] if regions else '')

    base_number = request.args.get('base_number', 100, type=int)
    latest_attendance = {}
    average_attendance = {}

    if os.path.exists(DATA_CACHE_PATH):
        try:
            df = pd.read_pickle(DATA_CACHE_PATH)
            if region == '總計':
                region_df = df
            else:
                region_df = df[df['大區'] == region]

            if not region_df.empty:
                ts = analysis.build_region_timeseries(region_df, region)
                if not ts.empty:
                    latest_data = ts.iloc[-1]
                    latest_attendance = latest_data.to_dict()
                    average_data = ts.mean()
                    average_attendance = average_data.to_dict()
        except Exception as e:
            app.logger.error(f"Error reading or processing cache for rates: {e}")

    subdistricts = find_subdistricts_for_region(charts_dir, region)
    subdistrict_cards = build_subdistrict_cards(charts_dir, region, subdistricts)

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
        base_number=base_number,
        latest_attendance=latest_attendance,
        average_attendance=average_attendance,
    )


@app.route('/upload', methods=['POST'])
def upload_files():
    """Handles file uploads, triggers analysis, and returns JSON status."""
    if 'files[]' not in request.files:
        return jsonify({'status': 'error', 'messages': ['請求中找不到檔案部分。']})

    files = request.files.getlist('files[]')
    if not files or files[0].filename == '':
        return jsonify({'status': 'error', 'messages': ['沒有選擇檔案。']})

    uploaded_count = 0
    for file in files:
        if file and allowed_file(file.filename):
            # filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            uploaded_count += 1

    messages = []
    if uploaded_count > 0:
        messages.append(f'成功上傳 {uploaded_count} 個檔案。')
        analysis_message = run_analysis()
        messages.append(analysis_message)
        return jsonify({'status': 'success', 'messages': messages})
    else:
        messages.append('沒有上傳有效的檔案。請選擇副檔名為 ' + ", ".join(ALLOWED_EXTENSIONS) + ' 的檔案。')
        return jsonify({'status': 'error', 'messages': messages})


@app.route('/region_data')
def region_data():
    """Provides chart and rate data for a specific region as JSON."""
    region = request.args.get('region', '總計', type=str)
    charts_dir = os.path.join(app.static_folder, 'charts')

    # Chart paths
    attendance_chart_path = f'charts/{region}_attendance.png'
    burden_chart_path = f'charts/{region}_burden.png'

    # Check if charts exist
    has_attendance_chart = os.path.exists(os.path.join(app.static_folder, attendance_chart_path))
    has_burden_chart = os.path.exists(os.path.join(app.static_folder, burden_chart_path))

    # Find subdistricts and build cards
    subdistricts = find_subdistricts_for_region(charts_dir, region)
    subdistrict_cards = build_subdistrict_cards(charts_dir, region, subdistricts)

    # Generate full URLs for the frontend
    for card in subdistrict_cards:
        card['attendance_chart'] = url_for('static', filename=card['attendance_chart'])
        card['burden_chart'] = url_for('static', filename=card['burden_chart'])

    return jsonify({
        'region': region,
        'attendance_chart_url': url_for('static', filename=attendance_chart_path) if has_attendance_chart else None,
        'burden_chart_url': url_for('static', filename=burden_chart_path) if has_burden_chart else None,
        'subdistrict_cards': subdistrict_cards,
    })

@app.route('/calculate_rates')
def calculate_rates():
    """
    Calculates attendance rates based on cached data and a base number.
    This is a read-only endpoint that does not modify any data.
    """
    base_number = request.args.get('base_number', 100, type=int)
    region = request.args.get('region', '總計', type=str)

    if not os.path.exists(DATA_CACHE_PATH):
        return jsonify({'status': 'error', 'message': '找不到彙整後的資料，請先執行分析。'}), 404

    try:
        df = pd.read_pickle(DATA_CACHE_PATH)
        region_df = df if region == '總計' else df[df['大區'] == region]

        if region_df.empty:
            return jsonify({'status': 'success', 'rates': {}})

        ts = analysis.build_region_timeseries(region_df, region)
        if ts.empty:
            return jsonify({'status': 'success', 'rates': {}})

        latest_data = ts.iloc[-1].to_dict()
        average_data = ts.mean().to_dict()

        rates = {}
        metrics = ['主日', '小排', '晨興', '禱告']

        for metric in metrics:
            latest_count = latest_data.get(metric, 0)
            avg_count = average_data.get(metric, 0)
            
            # Apply custom formulas
            formula = ((latest_count - base_number) / base_number * 100) if metric == '主日' else (latest_count / base_number * 100)
            avg_formula = ((avg_count - base_number) / base_number * 100) if metric == '主日' else (avg_count / base_number * 100)

            rates[metric] = {'latest_rate': formula if base_number > 0 else float('nan'), 'avg_rate': avg_formula if base_number > 0 else float('nan'), 'latest_count': latest_count, 'avg_count': avg_count}
        return jsonify({'status': 'success', 'rates': rates})
    except Exception as e:
        app.logger.error(f"Error during rate calculation: {e}", exc_info=True)
        return jsonify({'status': 'error', 'message': f'計算時發生錯誤: {e}'}), 500

if __name__ == '__main__':
    app.run(debug=True)
