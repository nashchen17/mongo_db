from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from pymongo import MongoClient
import pandas as pd
import os
import io

app = Flask(__name__, static_folder="static", template_folder="templates")

# Config via environment variables
MONGO_URI = os.environ.get("MONGO_URI", "mongodb://mongo:27017/")
DB_NAME = os.environ.get("DB_NAME", "mydb")
COLLECTION_NAME = os.environ.get("COLLECTION_NAME", "items")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db[COLLECTION_NAME]

# 新增採購與出貨表、庫存與採購需求表的匯入 API
PURCHASE_SHIPPING_COLLECTION_NAME = os.environ.get("PURCHASE_SHIPPING_COLLECTION_NAME", "purchase_shipping")
INVENTORY_NEED_COLLECTION_NAME = os.environ.get("INVENTORY_NEED_COLLECTION_NAME", "inventory_need")
purchase_shipping_collection = db[PURCHASE_SHIPPING_COLLECTION_NAME]
inventory_need_collection = db[INVENTORY_NEED_COLLECTION_NAME]

# 新增客戶需求表匯入 API
CUSTOMER_NEED_COLLECTION_NAME = os.environ.get("CUSTOMER_NEED_COLLECTION_NAME", "customer_need")
customer_need_collection = db[CUSTOMER_NEED_COLLECTION_NAME]


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/items", methods=["GET"])
def get_items():
    # Return all documents (limit to 1000 by default to avoid huge responses)
    limit = int(request.args.get("limit", "1000"))
    docs = list(collection.find({}, {"_id": 0}).limit(limit))
    return jsonify({"count": len(docs), "items": docs})


@app.route("/api/upload", methods=["POST"])
def upload_excel():
    """
    Expects a form-data request with a file field named 'file'.
    Reads the first sheet of the Excel file and inserts rows into MongoDB.
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"ok": False, "error": "No selected file"}), 400

    filename = secure_filename(file.filename)
    try:
        # Read file bytes into pandas
        in_memory = io.BytesIO(file.read())
        df = pd.read_excel(in_memory, engine="openpyxl")

        # Normalize dataframe: convert NaN to None
        df = df.where(pd.notnull(df), None)

        # Convert to list of dicts
        records = df.to_dict(orient="records")
        if len(records) == 0:
            return jsonify({"ok": False, "error": "Excel file contains no rows"}), 400

        # Insert into MongoDB
        result = collection.insert_many(records)
        inserted = len(result.inserted_ids)
        return jsonify({"ok": True, "inserted": inserted})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/clear", methods=["POST"])
def clear_db():
    """
    Clears the configured collection (drops all documents).
    Safety: requires JSON body { "confirm": true } to perform the operation.
    """
    data = request.get_json(force=True, silent=True)
    if not data or not data.get("confirm"):
        return jsonify({"ok": False, "error": "Missing confirmation. Send JSON {\"confirm\": true}"}), 400

    try:
        # Option 1: drop the collection
        collection.drop()
        return jsonify({"ok": True, "message": f"Collection '{COLLECTION_NAME}' dropped."})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/upload_purchase_shipping", methods=["POST"])
def upload_purchase_shipping_excel():
    """
    上傳 Excel 檔案並匯入採購與出貨表。
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"ok": False, "error": "No selected file"}), 400
    try:
        in_memory = io.BytesIO(file.read())
        df = pd.read_excel(in_memory, engine="openpyxl")
        df = df.where(pd.notnull(df), None)
        records = df.to_dict(orient="records")
        if len(records) == 0:
            return jsonify({"ok": False, "error": "Excel file contains no rows"}), 400
        result = purchase_shipping_collection.insert_many(records)
        inserted = len(result.inserted_ids)
        return jsonify({"ok": True, "inserted": inserted})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/upload_inventory_need", methods=["POST"])
def upload_inventory_need_excel():
    """
    上傳 Excel 檔案並匯入庫存與採購需求表。
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"ok": False, "error": "No selected file"}), 400
    try:
        in_memory = io.BytesIO(file.read())
        df = pd.read_excel(in_memory, engine="openpyxl")
        df = df.where(pd.notnull(df), None)
        records = df.to_dict(orient="records")
        if len(records) == 0:
            return jsonify({"ok": False, "error": "Excel file contains no rows"}), 400
        result = inventory_need_collection.insert_many(records)
        inserted = len(result.inserted_ids)
        return jsonify({"ok": True, "inserted": inserted})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/upload_customer_need", methods=["POST"])
def upload_customer_need_excel():
    """
    上傳 Excel 檔案並匯入客戶需求表。
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"ok": False, "error": "No selected file"}), 400
    try:
        in_memory = io.BytesIO(file.read())
        df = pd.read_excel(in_memory, engine="openpyxl")
        # 將 NaT 及 datetime 欄位全部轉為 None 或 ISO 格式字串
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].apply(lambda x: x.isoformat() if not pd.isna(x) and hasattr(x, 'isoformat') else None)
            else:
                df[col] = df[col].apply(lambda x: None if pd.isna(x) else x)
        records = df.to_dict(orient="records")
        if len(records) == 0:
            return jsonify({"ok": False, "error": "Excel file contains no rows"}), 400
        result = customer_need_collection.insert_many(records)
        inserted = len(result.inserted_ids)
        return jsonify({"ok": True, "inserted": inserted})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# 撿貨資訊表搜尋 API
@app.route("/api/search_pick", methods=["GET"])
def search_pick():
    """
    以 MIC需求起日 為條件，分別搜尋三個資料庫，回傳指定欄位。
    Query param: mic_start (MIC需求起日)
    """
    mic_start = request.args.get("mic_start")
    mic_end = request.args.get("mic_end")
    if not mic_start or not mic_end:
        return jsonify({"ok": False, "error": "缺少 MIC需求起日區間 參數"}), 400
    # 欄位
    fields = ["MIC需求起日", "MIC需求訖日", "料號", "版本", "產品中文名稱", "數量", "單價", "PO單號", "庫存"]
    result = {}
    # 日期區間搜尋
    from datetime import datetime
    def parse_date(s):
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%dT%H:%M:%S", "%Y/%m/%dT%H:%M:%S"]:
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                continue
        try:
            return datetime.fromisoformat(s)
        except Exception:
            return None
    start_dt = parse_date(mic_start)
    end_dt = parse_date(mic_end)
    pick_results = []
    import sys
    for name, coll in [
        ("purchase_shipping", purchase_shipping_collection),
        ("inventory_need", inventory_need_collection),
        ("customer_need", customer_need_collection)
    ]:
        query = {"$or": []}
        for fmt in [lambda s: s, lambda s: s.replace("-", "/"), lambda s: s.replace("/", "-")]:
            start_str = fmt(mic_start)
            end_str = fmt(mic_end)
            query["$or"].append({"MIC需求起日": {"$gte": start_str, "$lte": end_str}})
            query["$or"].append({"MIC需求起日": {"$in": [start_str, end_str]}})
        if start_dt and end_dt:
            query["$or"].append({"MIC需求起日": {"$gte": start_dt, "$lte": end_dt}})
            query["$or"].append({"MIC需求起日": {"$in": [start_dt, end_dt]}})
        # Debug print
        print(f"[DEBUG] Searching {name} with query: {query}", file=sys.stderr)
        docs = list(coll.find(query, {f: 1 for f in fields}))
        print(f"[DEBUG] Found {len(docs)} records in {name} for MIC需求起日 between {mic_start} and {mic_end}", file=sys.stderr)
        import math
        for d in docs:
            d.pop("_id", None)
            # 日期欄位格式化，移除 T 之後內容
            for date_field in ["MIC需求起日", "MIC需求訖日"]:
                if date_field in d and isinstance(d[date_field], str):
                    d[date_field] = d[date_field].split('T')[0]
            # 將 NaN 轉為 None，避免 JSON 錯誤
            for k, v in d.items():
                if isinstance(v, float) and math.isnan(v):
                    d[k] = None
            pick_results.append(d)
    # 取得所有料號
    partnos = list({row.get("料號") for row in pick_results if row.get("料號")})
    # 以料號搜尋所有資料庫，取得 產品中文名稱、單價、庫存
    enrich_fields = ["產品中文名稱", "單價", "庫存"]
    enrich_data = {}
    search_collections = []
    # 產品資料庫
    PRODUCTS_COLLECTION_NAME = os.environ.get("PRODUCTS_COLLECTION_NAME", "products")
    products_collection = db[PRODUCTS_COLLECTION_NAME]
    search_collections.append(products_collection)
    # 其他三個資料庫
    search_collections.extend([purchase_shipping_collection, inventory_need_collection, customer_need_collection])
    import math
    for partno in partnos:
        enrich_data[partno] = {}
        for coll in search_collections:
            for key in [partno, None]:
                if key is None:
                    try:
                        partno_float = float(partno)
                        keys_to_try = [partno_float, str(partno_float)]
                    except Exception:
                        keys_to_try = []
                else:
                    keys_to_try = [key]
                for k in keys_to_try:
                    doc = coll.find_one({"料號": k}, {f: 1 for f in enrich_fields})
                    if doc:
                        for f in enrich_fields:
                            val = doc.get(f)
                            if isinstance(val, float) and math.isnan(val):
                                val = None
                            if val is not None and f not in enrich_data[partno]:
                                enrich_data[partno][f] = val
                    if all(f in enrich_data[partno] for f in enrich_fields):
                        break
                if all(f in enrich_data[partno] for f in enrich_fields):
                    break
    # 合併 enrich_data 到 pick_results
    for row in pick_results:
        partno = row.get("料號")
        if partno and partno in enrich_data:
            for f in enrich_fields:
                val = enrich_data[partno].get(f)
                if isinstance(val, float) and math.isnan(val):
                    val = None
                if val is not None:
                    row[f] = val
    # 分組回傳
    result = {"pick": pick_results}
    return jsonify({"ok": True, "data": result})


# Static files (optional)
@app.route("/static/<path:path>")
def send_static(path):
    return send_from_directory("static", path)


if __name__ == "__main__":
    # For local testing only. In production use gunicorn in Dockerfile.
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)