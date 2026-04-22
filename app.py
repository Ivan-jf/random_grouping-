import os
import uuid
import random
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

# 自动创建必要目录
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def random_grouping(df, group_num, group_names, id_col, var_col, filter_col=None):
    """
    随机分组核心逻辑
    - group_num: 分组数
    - group_names: 分组名称列表
    - id_col: 样品ID列名
    - var_col: 用于排序分组的变量列名
    - filter_col: "要不要"列名，只取Y的行（可选）
    """
    # 过滤"要不要"列
    if filter_col and filter_col in df.columns:
        df_excluded = df[df[filter_col].astype(str).str.strip().str.upper() != 'Y'].copy()
        df = df[df[filter_col].astype(str).str.strip().str.upper() == 'Y'].copy()
        df = df.reset_index(drop=True)
    else:
        df_excluded = pd.DataFrame()

    n = len(df)
    expected = group_num * len(group_names)
    if n != expected:
        raise ValueError(
            f"筛选后有效行数为 {n}，但 分组数({group_num}) × 组名数量({len(group_names)}) = {expected}，两者必须相等，请检查数据或参数。"
        )

    # 按 var_col 降序排列
    df_sorted = df.sort_values(by=var_col, ascending=False).reset_index(drop=True)

    # 分块
    df_sorted['block'] = [i // (n/group_num) + 1 for i in range(n)]

    # 每块内随机排序并分配组名
    df_sorted['group'] = None

    for block_id in df_sorted['block'].unique():
        idx = df_sorted[df_sorted['block'] == block_id].index.tolist()
        shuffled = idx.copy()
        random.shuffle(shuffled)
        names_for_block = group_names[:len(idx)]
        for pos, original_idx in enumerate(shuffled):
            df_sorted.at[original_idx, 'group'] = names_for_block[pos]

    # 统计摘要
    summary_df = df_sorted.groupby('group')[var_col].agg(
        mean='mean',
        sd='std',
        min='min',
        max='max'
    ).reset_index()
    summary_df.columns = ['分组', '均值', '标准差', '最小值', '最大值']
    summary_df = summary_df.round(4)

    return df_sorted, summary_df, df_excluded


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/get_columns', methods=['POST'])
def get_columns():
    """上传文件后返回列名供用户选择"""
    if 'file' not in request.files:
        return jsonify({'error': '未上传文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': '仅支持 .xlsx / .xls 格式'}), 400

    filename = secure_filename(file.filename)
    uid = str(uuid.uuid4())[:8]
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], uid + '_' + filename)
    file.save(save_path)

    try:
        df = pd.read_excel(save_path)
        columns = df.columns.tolist()
        return jsonify({'columns': columns, 'filepath': save_path})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/run', methods=['POST'])
def run():
    data = request.get_json()
    filepath   = data.get('filepath')
    group_num  = int(data.get('group_num'))
    group_names = [g.strip() for g in data.get('group_names', '').split(',') if g.strip()]
    id_col     = data.get('id_col')
    var_col    = data.get('var_col')
    filter_col = data.get('filter_col') or None

    if not filepath or not os.path.exists(filepath):
        return jsonify({'error': '找不到上传的文件，请重新上传'}), 400
    if not group_names:
        return jsonify({'error': '请输入分组名称'}), 400

    try:
        df = pd.read_excel(filepath)
        df_result, summary_df, df_excluded = random_grouping(
            df, group_num, group_names, id_col, var_col, filter_col
        )

        # 保存结果
        out_name = 'result_' + str(uuid.uuid4())[:8] + '.xlsx'
        out_path = os.path.join(app.config['OUTPUT_FOLDER'], out_name)

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='分组结果', index=False)
            summary_df.to_excel(writer, sheet_name='统计摘要', index=False)
            if not df_excluded.empty:
                df_excluded.to_excel(writer, sheet_name='未参与样本', index=False)

        # 返回预览数据
        summary_html = summary_df.to_dict(orient='records')
        preview = df_result[[id_col, var_col, 'block', 'group']].head(20).to_dict(orient='records')

        return jsonify({
            'success': True,
            'out_file': out_name,
            'summary': summary_html,
            'preview': preview,
            'total': len(df_result)
        })

    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'运行出错：{str(e)}'}), 500


@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if not os.path.exists(path):
        return '文件不存在', 404
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    app.run(debug=True)
