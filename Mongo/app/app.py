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
    if not mic_start:
        return jsonify({"ok": False, "error": "缺少 MIC需求起日 參數"}), 400
    # 欄位
    fields = ["MIC需求起日", "MIC需求訖日", "料號", "版本", "產品中文名稱", "數量", "單價", "PO單號", "庫存"]
    result = {}
    # 各資料庫搜尋
    # 先依日期搜尋
    pick_results = []
    for name, coll in [
        ("purchase_shipping", purchase_shipping_collection),
        ("inventory_need", inventory_need_collection),
        ("customer_need", customer_need_collection)
    ]:
        # 支援多種日期格式搜尋，包含 ISODate 格式與時間部分
        try:
            from datetime import datetime
            iso_date = None
            if len(mic_start) in [8, 10]:
                # yyyy/mm/dd 或 yyyy-mm-dd
                fmt = "%Y/%m/%d" if "/" in mic_start else "%Y-%m-%d"
                iso_date = datetime.strptime(mic_start, fmt)
                # 也搜尋該日的 00:00:00
                iso_date_full = datetime.strptime(mic_start, fmt).replace(hour=0, minute=0, second=0)
            elif len(mic_start) == 19:
                # yyyy-mm-ddTHH:MM:SS
                iso_date = datetime.fromisoformat(mic_start)
                iso_date_full = iso_date
            else:
                iso_date_full = None
        except Exception:
            iso_date = None
            iso_date_full = None
        query = {"$or": [
            {"MIC需求起日": mic_start},
            {"MIC需求起日": mic_start.replace("-", "/")},
            {"MIC需求起日": mic_start.replace("/", "-")},
            {"MIC需求起日": iso_date} if iso_date else {},
            {"MIC需求起日": iso_date_full} if iso_date_full else {},
            {"MIC需求起日": {"$regex": f"^{mic_start}"}}  # 支援部分比對
        ]}
        docs = list(coll.find(query, {f: 1 for f in fields}))
        for d in docs:
            d.pop("_id", None)
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
    for partno in partnos:
        enrich_data[partno] = {}
        for coll in search_collections:
            # 嘗試三種型態查詢
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
                            if f in doc and doc[f] is not None and f not in enrich_data[partno]:
                                enrich_data[partno][f] = doc[f]
                    # 若三欄都找到就跳出
                    if all(f in enrich_data[partno] for f in enrich_fields):
                        break
                if all(f in enrich_data[partno] for f in enrich_fields):
                    break
    # 合併 enrich_data 到 pick_results
    for row in pick_results:
        partno = row.get("料號")
        if partno and partno in enrich_data:
            for f in enrich_fields:
                if enrich_data[partno].get(f) is not None:
                    row[f] = enrich_data[partno][f]
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