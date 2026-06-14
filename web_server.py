import socket
import threading
from datetime import datetime
from flask import Flask, jsonify, request, render_template_string

app = Flask(__name__)
_db = None
_get_vendors = None

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mobile Inventory</title>
    <script src="https://unpkg.com/html5-qrcode"></script>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 10px; background: #f4f4f9; }
        h2 { text-align: center; color: #333; margin-top: 5px; }
        .controls { background: white; padding: 10px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 15px; }
        select, input { width: 100%; padding: 10px; margin: 5px 0; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; font-size: 16px; }
        .search-row { display: flex; gap: 10px; align-items: center; }
        .btn-scan { background: #17a2b8; color: white; border: none; padding: 10px 15px; border-radius: 5px; cursor: pointer; font-size: 16px; white-space: nowrap; }
        .btn-scan-small { background: #17a2b8; color: white; border: none; padding: 8px; border-radius: 5px; cursor: pointer; font-size: 14px; margin-left: 5px; }
        .product-card { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 10px; }
        .product-card h3 { margin: 0 0 10px 0; font-size: 16px; color: #0056b3; }
        .product-card p { margin: 4px 0; font-size: 14px; color: #555; }
        .btn-edit { background: #007bff; color: white; border: none; padding: 8px 12px; border-radius: 5px; cursor: pointer; width: 100%; font-size: 16px; margin-top: 10px; }
        .btn-edit:hover { background: #0056b3; }
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; overflow-y: auto; }
        .modal-content { background: white; margin: 20px auto; padding: 20px; width: 90%; max-width: 400px; border-radius: 8px; box-sizing: border-box; }
        .field-row { margin-bottom: 10px; }
        .field-row label { display: block; font-weight: bold; margin-bottom: 5px; font-size: 14px; }
        .btn-save { background: #28a745; color: white; border: none; padding: 12px; border-radius: 5px; cursor: pointer; width: 100%; font-size: 16px; margin-top: 10px; }
        .btn-cancel { background: #dc3545; color: white; border: none; padding: 12px; border-radius: 5px; cursor: pointer; width: 100%; font-size: 16px; margin-top: 10px; }
        .toast { visibility: hidden; min-width: 250px; background-color: #333; color: #fff; text-align: center; border-radius: 2px; padding: 16px; position: fixed; z-index: 1; left: 50%; bottom: 30px; font-size: 17px; transform: translateX(-50%); }
        .toast.show { visibility: visible; animation: fadein 0.5s, fadeout 0.5s 2.5s; }
        @keyframes fadein { from {bottom: 0; opacity: 0;} to {bottom: 30px; opacity: 1;} }
        @keyframes fadeout { from {bottom: 30px; opacity: 1;} to {bottom: 0; opacity: 0;} }
    </style>
</head>
<body>
    <h2>Mobile Inventory</h2>
    <div class="controls">
        <select id="vendor-select" onchange="loadProducts()">
            <option value="">Select Vendor...</option>
            {% for v in vendors %}
            <option value="{{ v }}">{{ v }}</option>
            {% endfor %}
        </select>
        <div class="search-row">
            <input type="text" id="search-input" placeholder="Search by name or SKU..." onkeyup="filterProducts()">
            <button class="btn-scan" onclick="startGlobalScanner()">Scan</button>
        </div>
        <select id="brand-select" onchange="filterProducts()">
            <option value="">All Brands</option>
        </select>
    </div>
    
    <div id="product-list"></div>

    <div id="edit-modal" class="modal">
        <div class="modal-content">
            <h3 style="margin-top: 0;">Edit Product</h3>
            <div class="field-row"><label>SKU (Required):</label><input type="text" id="edit-sku"></div>
            <div class="field-row"><label>Product Name:</label><input type="text" id="edit-name"></div>
            <div class="field-row"><label>Brand:</label><input type="text" id="edit-brand"></div>
            <div class="field-row"><label>Pack:</label><input type="text" id="edit-pack"></div>
            <div class="field-row"><label>Unit:</label><input type="text" id="edit-unit"></div>
            <div class="field-row"><label>CTN Qty:</label><input type="number" id="edit-ctn" min="0"></div>
            <div class="field-row"><label>Price (€):</label><input type="number" id="edit-price" step="0.01"></div>
            <div class="field-row"><label>Stock:</label><input type="text" id="edit-stock"></div>
            <div class="field-row">
                <label>Barcode:</label>
                <div style="display: flex;">
                    <input type="text" id="edit-barcode">
                    <button class="btn-scan-small" onclick="startFieldScanner()">Scan</button>
                </div>
            </div>
            <div class="field-row"><label>SR-SKU:</label><input type="text" id="edit-srsku"></div>
            <button class="btn-save" onclick="saveProduct()">Save Changes</button>
            <button class="btn-cancel" onclick="closeModal()">Cancel</button>
        </div>
    </div>

    <div id="scanner-modal" class="modal" style="z-index: 1050;">
        <div class="modal-content">
            <h3 style="margin-top: 0;">Scan Barcode</h3>
            <div id="reader" style="width: 100%;"></div>
            <button class="btn-cancel" onclick="closeScanner()">Cancel</button>
        </div>
    </div>

    <div id="toast" class="toast">Saved Successfully</div>

    <script>
        let products = [];
        let html5QrCode = null;
        let scanMode = 'global';
        
        async function loadProducts() {
            const vendor = document.getElementById('vendor-select').value;
            if (!vendor) return;
            try {
                const response = await fetch(`/api/inventory/${vendor}`);
                products = await response.json();
                
                const brands = new Set(products.map(p => p.brand).filter(b => b));
                const brandSelect = document.getElementById('brand-select');
                brandSelect.innerHTML = '<option value="">All Brands</option>';
                [...brands].sort().forEach(b => {
                    brandSelect.innerHTML += `<option value="${b}">${b}</option>`;
                });
                
                filterProducts();
            } catch (e) { alert('Error loading products: ' + e); }
        }
        
        function filterProducts() {
            const term = document.getElementById('search-input').value.toLowerCase();
            const brand = document.getElementById('brand-select').value;
            
            const filtered = products.filter(p => {
                const matchTerm = !term || 
                                  (p.product_name && p.product_name.toLowerCase().includes(term)) || 
                                  (p.sku && p.sku.toLowerCase().includes(term)) ||
                                  (p.barcode && p.barcode.toLowerCase().includes(term));
                const matchBrand = !brand || p.brand === brand;
                return matchTerm && matchBrand;
            });
            renderProducts(filtered);
        }
        
        function renderProducts(list) {
            const container = document.getElementById('product-list');
            container.innerHTML = '';
            list.forEach(p => {
                const div = document.createElement('div');
                div.className = 'product-card';
                div.innerHTML = `
                    <h3>${p.product_name || 'No Name'}</h3>
                    <p><b>SKU:</b> ${p.sku} | <b>Brand:</b> ${p.brand || 'N/A'}</p>
                    <p><b>Pack:</b> ${p.pack || '-'} ${p.unit || ''} | <b>CTN Qty:</b> ${p.ctn_qty || 0}</p>
                    <p><b>Price:</b> €${parseFloat(p.price || 0).toFixed(2)} | <b>Stock:</b> ${p.stock || 'N/A'}</p>
                    <button class='btn-edit' onclick='openEdit(${JSON.stringify(p).replace(/'/g, "&#39;")})'>Edit</button>
                `;
                container.appendChild(div);
            });
        }
        
        function openEdit(p) {
            document.getElementById('edit-sku').value = p.sku || '';
            document.getElementById('edit-sku').readOnly = !!p.sku; // Disable editing SKU if it exists
            document.getElementById('edit-name').value = p.product_name || '';
            document.getElementById('edit-brand').value = p.brand || '';
            document.getElementById('edit-pack').value = p.pack || '';
            document.getElementById('edit-unit').value = p.unit || '';
            document.getElementById('edit-ctn').value = p.ctn_qty || 0;
            document.getElementById('edit-price').value = p.price || 0;
            document.getElementById('edit-stock').value = p.stock || '';
            document.getElementById('edit-barcode').value = p.barcode || '';
            document.getElementById('edit-srsku').value = p.sr_sku || '';
            document.getElementById('edit-modal').style.display = 'block';
        }

        function startGlobalScanner() {
            const vendor = document.getElementById('vendor-select').value;
            if (!vendor) { alert("Please select a vendor first."); return; }
            scanMode = 'global';
            openScannerModal();
        }

        function startFieldScanner() {
            scanMode = 'field';
            openScannerModal();
        }

        function openScannerModal() {
            document.getElementById('scanner-modal').style.display = 'block';
            if (!html5QrCode) {
                html5QrCode = new Html5Qrcode("reader");
            }
            html5QrCode.start(
                { facingMode: "environment" },
                { fps: 10, qrbox: { width: 250, height: 150 } },
                (decodedText) => {
                    closeScanner();
                    if (scanMode === 'global') {
                        handleGlobalScan(decodedText);
                    } else {
                        document.getElementById('edit-barcode').value = decodedText;
                    }
                },
                (errorMessage) => { }
            ).catch(err => {
                alert("Camera error: " + err);
                closeScanner();
            });
        }

        function closeScanner() {
            document.getElementById('scanner-modal').style.display = 'none';
            if (html5QrCode) {
                html5QrCode.stop().then(() => {
                    html5QrCode.clear();
                    html5QrCode = null;
                }).catch(err => console.log(err));
            }
        }

        function handleGlobalScan(barcode) {
            document.getElementById('search-input').value = barcode;
            const product = products.find(p => p.barcode == barcode);
            if (product) {
                filterProducts();
            } else {
                if (confirm(`Barcode ${barcode} not found.\\nDo you want to create a new product with this barcode?`)) {
                    openEdit({ barcode: barcode });
                } else {
                    filterProducts();
                }
            }
        }
        
        function closeModal() {
            document.getElementById('edit-modal').style.display = 'none';
        }
        
        async function saveProduct() {
            const vendor = document.getElementById('vendor-select').value;
            const sku = document.getElementById('edit-sku').value;
            if(!sku) { alert("SKU is required."); return; }

            const data = {
                product_name: document.getElementById('edit-name').value,
                brand: document.getElementById('edit-brand').value,
                pack: document.getElementById('edit-pack').value,
                unit: document.getElementById('edit-unit').value,
                ctn_qty: parseInt(document.getElementById('edit-ctn').value || 0),
                price: parseFloat(document.getElementById('edit-price').value || 0),
                stock: document.getElementById('edit-stock').value,
                barcode: document.getElementById('edit-barcode').value,
                sr_sku: document.getElementById('edit-srsku').value
            };
            
            try {
                const res = await fetch(`/api/inventory/${vendor}/${sku}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                });
                
                if (res.ok) {
                    closeModal();
                    showToast();
                    loadProducts();
                } else { alert('Failed to save product'); }
            } catch (e) { alert('Error: ' + e); }
        }

        function showToast() {
            const x = document.getElementById("toast");
            x.className = "toast show";
            setTimeout(function(){ x.className = x.className.replace("show", ""); }, 3000);
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    vendors = _get_vendors() if _get_vendors else []
    return render_template_string(HTML_TEMPLATE, vendors=vendors)

@app.route('/api/inventory/<vendor>', methods=['GET'])
def get_inventory(vendor):
    try:
        products = _db.get_vendor_products(vendor)
        return jsonify(products)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/inventory/<vendor>/<sku>', methods=['POST'])
def update_inventory(vendor, sku):
    data = request.json
    data['last_updated'] = datetime.utcnow().isoformat()
    try:
        products = _db.get_vendor_products(vendor)
        exists = any(str(p.get("sku")) == str(sku) for p in products)
        if exists:
            _db.update_vendor_product(vendor, sku, data)
        else:
            data['sku'] = sku
            _db.upsert_vendor_products(vendor, [data])
        return jsonify({"status": "success"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def start_web_server(db, get_vendors_callback):
    global _db, _get_vendors
    _db = db
    _get_vendors = get_vendors_callback
    
    try:
        import cryptography
        use_https = True
    except ImportError:
        use_https = False
    
    def run():
        if use_https:
            try:
                app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False, ssl_context='adhoc')
            except Exception as e:
                print("HTTPS failed, falling back to HTTP:", e)
                app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
        else:
            app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
        
    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        protocol = "https" if use_https else "http"
        url = f"{protocol}://{ip}:5000"
        if not use_https:
            url += " (Install 'cryptography')"
        return url
    except Exception:
        protocol = "https" if use_https else "http"
        return f"{protocol}://127.0.0.1:5000"