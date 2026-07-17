"""
公文格式助手 - Flask 服务器
"""

import os
import uuid
import time
import json
import logging
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

from formatter.engine import FormatEngine
from formatter.converter import convert_file, get_conversion_options, find_libreoffice
from formatter.image_utils import (
    resize_images_in_docx, validate_dimensions,
    get_document_images_info,
    A4_CONTENT_WIDTH_MM, A4_CONTENT_HEIGHT_MM
)

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 上传限制

# 目录配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
TEMP_DIR = os.path.join(BASE_DIR, 'temp')
META_DIR = os.path.join(BASE_DIR, 'meta')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(META_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {
    'docx', 'doc', 'pdf', 'pptx', 'ppt', 'xlsx', 'xls',
    'odt', 'ods', 'odp', 'rtf', 'txt', 'csv', 'html'
}


def allowed_file(filename):
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
    return ext in ALLOWED_EXTENSIONS


def get_file_ext(filename):
    return filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''


# ══════════════════════════════════════════════
# 文件存储（文件系统版，支持多 worker 和重启持久化）
# ══════════════════════════════════════════════
def save_meta(file_id, data):
    """保存文件元数据到磁盘"""
    path = os.path.join(META_DIR, f'{file_id}.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)


def load_meta(file_id):
    """从磁盘加载文件元数据"""
    path = os.path.join(META_DIR, f'{file_id}.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def delete_meta(file_id):
    """删除文件元数据"""
    path = os.path.join(META_DIR, f'{file_id}.json')
    if os.path.exists(path):
        os.remove(path)


def cleanup_expired_files():
    """清理超过1小时的临时文件"""
    now = time.time()
    try:
        for fname in os.listdir(META_DIR):
            if not fname.endswith('.json'):
                continue
            file_id = fname[:-5]
            meta = load_meta(file_id)
            if meta and now - meta.get('upload_time', 0) > 3600:
                # 删除文件
                file_path = meta.get('path', '')
                if file_path and os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                    except Exception:
                        pass
                delete_meta(file_id)
    except Exception:
        pass


# ══════════════════════════════════════════════
# 页面路由
# ══════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')


# ══════════════════════════════════════════════
# API: 文件上传
# ══════════════════════════════════════════════
@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "没有找到文件"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "error": "没有选择文件"}), 400

    if not allowed_file(file.filename):
        return jsonify({"success": False, "error": f"不支持的文件格式。支持: {', '.join(sorted(ALLOWED_EXTENSIONS))}"}), 400

    file_id = str(uuid.uuid4())[:8]
    ext = get_file_ext(file.filename)
    # 使用 file_id 作为文件名前缀，避免同名文件冲突
    safe_name = f"{file_id}_{secure_filename(file.filename)}"
    if not safe_name or not safe_name.endswith(f'.{ext}'):
        safe_name = f"{file_id}_upload.{ext}"
    save_path = os.path.join(UPLOAD_DIR, safe_name)

    try:
        file.save(save_path)
        meta = {
            "path": save_path,
            "original_name": file.filename,
            "ext": ext,
            "upload_time": time.time()
        }
        save_meta(file_id, meta)
        logger.info(f"文件上传成功: {file.filename} -> {file_id}")

        return jsonify({
            "success": True,
            "file_id": file_id,
            "filename": file.filename,
            "ext": ext,
            "size": os.path.getsize(save_path)
        })
    except Exception as e:
        logger.exception("文件上传失败")
        return jsonify({"success": False, "error": str(e)}), 500


# ══════════════════════════════════════════════
# API: 文件格式转换
# ══════════════════════════════════════════════
@app.route('/api/convert', methods=['POST'])
def api_convert():
    data = request.json or request.form
    file_id = data.get('file_id')
    target_format = data.get('target_format', '').lower()

    meta = load_meta(file_id)
    if not file_id or not meta:
        return jsonify({"success": False, "error": "文件不存在，请重新上传"}), 400

    if not target_format:
        return jsonify({"success": False, "error": "请指定目标格式"}), 400

    input_path = meta['path']
    source_ext = meta['ext']

    # 如果源格式和目标格式相同
    if source_ext == target_format:
        return jsonify({
            "success": True,
            "output_file_id": file_id,
            "message": "源文件已是目标格式，无需转换"
        })

    try:
        result = convert_file(input_path, target_format, TEMP_DIR)
        if result['success']:
            output_id = str(uuid.uuid4())[:8]
            output_meta = {
                "path": result['output_path'],
                "original_name": os.path.basename(result['output_path']),
                "ext": target_format,
                "upload_time": time.time()
            }
            save_meta(output_id, output_meta)
            return jsonify({
                "success": True,
                "output_file_id": output_id,
                "filename": os.path.basename(result['output_path'])
            })
        else:
            return jsonify({"success": False, "error": result['error']}), 500
    except Exception as e:
        logger.exception("格式转换异常")
        return jsonify({"success": False, "error": str(e)}), 500


# ══════════════════════════════════════════════
# API: 字体段落格式修改（公文排版）
# ══════════════════════════════════════════════
@app.route('/api/format', methods=['POST'])
def api_format():
    data = request.json or request.form
    file_id = data.get('file_id')

    meta = load_meta(file_id)
    if not file_id or not meta:
        return jsonify({"success": False, "error": "文件不存在，请重新上传"}), 400

    if meta['ext'] not in ('docx', 'doc'):
        return jsonify({"success": False, "error": "格式修改仅支持 Word 文档（.docx/.doc）"}), 400

    input_path = meta['path']
    output_name = f"formatted_{os.path.basename(input_path)}"
    output_path = os.path.join(TEMP_DIR, output_name)

    try:
        engine = FormatEngine()
        result = engine.process(input_path, output_path)

        output_id = str(uuid.uuid4())[:8]
        output_meta = {
            "path": result['output_path'],
            "original_name": output_name,
            "ext": "docx",
            "upload_time": time.time()
        }
        save_meta(output_id, output_meta)

        return jsonify({
            "success": True,
            "output_file_id": output_id,
            "filename": output_name,
            "confidence_marks": result['confidence_marks']
        })
    except Exception as e:
        logger.exception("格式修改异常")
        return jsonify({"success": False, "error": str(e)}), 500


# ══════════════════════════════════════════════
# API: 图片尺寸统一
# ══════════════════════════════════════════════
@app.route('/api/resize-images', methods=['POST'])
def api_resize_images():
    data = request.json or request.form
    file_id = data.get('file_id')
    target_width = data.get('target_width')  # mm
    target_height = data.get('target_height')  # mm
    mode = data.get('mode', 'fit')  # fit / exact / max

    meta = load_meta(file_id)
    if not file_id or not meta:
        return jsonify({"success": False, "error": "文件不存在，请重新上传"}), 400

    if meta['ext'] not in ('docx', 'doc'):
        return jsonify({"success": False, "error": "图片尺寸统一仅支持 Word 文档（.docx/.doc）"}), 400

    # 参数处理
    width_mm = float(target_width) if target_width else None
    height_mm = float(target_height) if target_height else None

    # 尺寸验证
    if mode == "exact" and width_mm and height_mm:
        validation = validate_dimensions(width_mm, height_mm)
        if not validation['valid']:
            return jsonify({"success": False, "error": validation['error']}), 400

    if width_mm and (width_mm > A4_CONTENT_WIDTH_MM or width_mm <= 0):
        return jsonify({"success": False, "error": f"宽度需在 1-{A4_CONTENT_WIDTH_MM}mm 范围内"}), 400
    if height_mm and (height_mm > A4_CONTENT_HEIGHT_MM or height_mm <= 0):
        return jsonify({"success": False, "error": f"高度需在 1-{A4_CONTENT_HEIGHT_MM}mm 范围内"}), 400

    input_path = meta['path']
    output_name = f"resized_{os.path.basename(input_path)}"
    output_path = os.path.join(TEMP_DIR, output_name)

    try:
        result = resize_images_in_docx(input_path, output_path, width_mm, height_mm, mode)

        if result['success']:
            output_id = str(uuid.uuid4())[:8]
            output_meta = {
                "path": output_path,
                "original_name": output_name,
                "ext": "docx",
                "upload_time": time.time()
            }
            save_meta(output_id, output_meta)
            return jsonify({
                "success": True,
                "output_file_id": output_id,
                "filename": output_name,
                "processed_count": result['processed_count']
            })
        else:
            return jsonify({"success": False, "error": result.get('error', '未知错误')}), 500
    except Exception as e:
        logger.exception("图片尺寸统一异常")
        return jsonify({"success": False, "error": str(e)}), 500


# ══════════════════════════════════════════════
# API: 获取图片信息
# ══════════════════════════════════════════════
@app.route('/api/images-info', methods=['POST'])
def api_images_info():
    data = request.json or request.form
    file_id = data.get('file_id')

    meta = load_meta(file_id)
    if not file_id or not meta:
        return jsonify({"success": False, "error": "文件不存在"}), 400

    if meta['ext'] not in ('docx', 'doc'):
        return jsonify({"success": False, "error": "仅支持 Word 文档"}), 400

    images = get_document_images_info(meta['path'])
    return jsonify({"success": True, "images": images, "count": len(images)})


# ══════════════════════════════════════════════
# API: 下载文件
# ══════════════════════════════════════════════
@app.route('/api/download/<file_id>')
def download_file(file_id):
    meta = load_meta(file_id)
    if not meta:
        return jsonify({"error": "文件不存在"}), 404

    return send_file(
        meta['path'],
        as_attachment=True,
        download_name=meta['original_name']
    )


# ══════════════════════════════════════════════
# API: 获取可转换格式
# ══════════════════════════════════════════════
@app.route('/api/conversion-options/<ext>')
def api_conversion_options(ext):
    options = get_conversion_options(ext)
    libre_available = find_libreoffice() is not None
    return jsonify({
        "options": options,
        "libreoffice_available": libre_available
    })


# ══════════════════════════════════════════════
# API: 页面尺寸限制
# ══════════════════════════════════════════════
@app.route('/api/page-limits')
def api_page_limits():
    return jsonify({
        "max_width_mm": A4_CONTENT_WIDTH_MM,
        "max_height_mm": A4_CONTENT_HEIGHT_MM,
        "paper": "A4",
        "margins": {"top": 37, "bottom": 35, "left": 28, "right": 26}
    })


# ══════════════════════════════════════════════
# API: 服务器状态
# ══════════════════════════════════════════════
@app.route('/api/status')
def api_status():
    return jsonify({
        "status": "running",
        "libreoffice": find_libreoffice() is not None,
        "version": "2.0.0"
    })


# ══════════════════════════════════════════════
# 清理过期文件（每次请求概率触发，避免高频遍历）
# ══════════════════════════════════════════════
_cleanup_counter = 0

@app.before_request
def maybe_cleanup():
    global _cleanup_counter
    _cleanup_counter += 1
    # 每 50 次请求清理一次
    if _cleanup_counter % 50 == 0:
        cleanup_expired_files()


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5800))
    logger.info(f"公文格式助手启动: http://localhost:{port}")
    logger.info(f"LibreOffice: {'已安装' if find_libreoffice() else '未安装（格式转换部分不可用）'}")
    app.run(host='0.0.0.0', port=port, debug=False)
