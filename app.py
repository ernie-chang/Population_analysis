from flask import Flask, render_template, request, url_for
import os

app = Flask(__name__)

@app.route('/')
def index():
    charts_dir = os.path.join(app.static_folder, 'charts')
    attendance_files = [f for f in os.listdir(charts_dir) if f.endswith('_attendance.png')]
    regions = [f.replace('_attendance.png', '') for f in attendance_files]
    # Add extra regions if needed
    for extra in ['國中大區', '青年大區']:
        if extra not in regions:
            regions.append(extra)
    # Get selected region from query param, default to first
    region = request.args.get('region', regions[0] if regions else '')
    attendance_chart = f'charts/{region}_attendance.png'
    burden_chart = f'charts/{region}_burden.png'
    return render_template('index.html', regions=regions, region=region, attendance_chart=attendance_chart, burden_chart=burden_chart)

if __name__ == '__main__':
    app.run(debug=True)
