import socket
import threading
from datetime import datetime
from flask import Flask, jsonify, request, render_template_string
import pandas as pd
import io
import requests

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
        .modal-content-large { background: white; margin: 20px auto; padding: 20px; width: 95%; max-width: 1200px; border-radius: 8px; box-sizing: border-box; max-height: 90vh; overflow-y: auto; }
        .field-row { margin-bottom: 10px; }
        .field-row label { display: block; font-weight: bold; margin-bottom: 5px; font-size: 14px; }
        .btn-save { background: #28a745; color: white; border: none; padding: 12px; border-radius: 5px; cursor: pointer; width: 100%; font-size: 16px; margin-top: 10px; }
        .btn-cancel { background: #dc3545; color: white; border: none; padding: 12px; border-radius: 5px; cursor: pointer; width: 100%; font-size: 16px; margin-top: 10px; }
        .btn-archive { background: #6c757d; color: white; border: none; padding: 10px 15px; border-radius: 5px; cursor: pointer; font-size: 16px; white-space: nowrap; margin-top: 10px; width: 100%; }
        .btn-archive:hover { background: #5a6268; }
        .btn-load { background: #007bff; color: white; border: none; padding: 12px; border-radius: 5px; cursor: pointer; width: 100%; font-size: 16px; margin-top: 10px; }
        .btn-load:hover { background: #0056b3; }
        .sheet-tabs { display: flex; gap: 5px; margin-bottom: 15px; flex-wrap: wrap; }
        .sheet-tab { padding: 10px 15px; background: #e9ecef; border: 1px solid #ccc; border-radius: 5px 5px 0 0; cursor: pointer; font-size: 14px; }
        .sheet-tab.active { background: #007bff; color: white; border-color: #007bff; }
        .sheet-content { display: none; }
        .sheet-content.active { display: block; }
        .data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .data-table th { background: #f8f9fa; border: 1px solid #ddd; padding: 8px; text-align: left; font-weight: bold; }
        .data-table td { border: 1px solid #ddd; padding: 8px; }
        .data-table tr:nth-child(even) { background: #f9f9f9; }
        .data-table tr:hover { background: #f0f0f0; }
        .toast { visibility: hidden; min-width: 250px; background-color: #333; color: #fff; text-align: center; border-radius: 2px; padding: 16px; position: fixed; z-index: 1; left: 50%; bottom: 30px; font-size: 17px; transform: translateX(-50%); }
        .toast.show { visibility: visible; animation: fadein 0.5s, fadeout 0.5s 2.5s; }
        .loading { text-align: center; padding: 20px; }
        .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #007bff; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
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
        <button class="btn-archive" onclick="openArchiveWidget()">SR Products Archive</button>
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

    <div id="archive-modal" class="modal" style="z-index: 1050;">
        <div class="modal-content-large">
            <h3 style="margin-top: 0; display: flex; justify-content: space-between; align-items: center;">
                SR Products Archive
                <button onclick="closeArchiveWidget()" style="background: none; border: none; font-size: 24px; cursor: pointer;">&times;</button>
            </h3>
            <div class="field-row">
                <label>Dropbox Excel Sheet URL:</label>
                <input type="text" id="dropbox-url" placeholder="https://www.dropbox.com/..." style="font-size: 14px;">
                <small style="color: #666; margin-top: 5px; display: block;">Paste the Dropbox Excel file link here</small>
            </div>
            <button class="btn-load" onclick="loadDropboxSheet()">Load Sheet</button>
            <div id="archive-content" style="margin-top: 20px;"></div>
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

        // Archive Widget Functions
        function openArchiveWidget() {
            document.getElementById('archive-modal').style.display = 'block';
        }

        function closeArchiveWidget() {
            document.getElementById('archive-modal').style.display = 'none';
            document.getElementById('archive-content').innerHTML = '';
            document.getElementById('dropbox-url').value = '';
        }

        async function loadDropboxSheet() {
            const url = document.getElementById('dropbox-url').value.trim();
            if (!url) {
                alert('Please enter a Dropbox URL');
                return;
            }

            const archiveContent = document.getElementById('archive-content');
            archiveContent.innerHTML = '<div class="loading"><div class="spinner"></div><p>Loading...</p></div>';

            try {
                const response = await fetch('/api/load-excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: url })
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.error || 'Failed to load file');
                }

                const data = await response.json();
                displaySheets(data.sheets);
            } catch (e) {
                archiveContent.innerHTML = `<div style="color: red; padding: 20px; background: #ffe0e0; border-radius: 5px;"><b>Error:</b> ${e.message}</div>`;
            }
        }

        function displaySheets(sheets) {
            const archiveContent = document.getElementById('archive-content');
            archiveContent.innerHTML = '';

            if (!sheets || Object.keys(sheets).length === 0) {
                archiveContent.innerHTML = '<p>No sheets found in the Excel file.</p>';
                return;
            }

            const sheetNames = Object.keys(sheets);
            
            // Create tabs
            const tabsContainer = document.createElement('div');
            tabsContainer.className = 'sheet-tabs';

            sheetNames.forEach((sheetName, index) => {
                const tab = document.createElement('button');
                tab.className = `sheet-tab ${index === 0 ? 'active' : ''}`;
                tab.textContent = sheetName;
                tab.onclick = () => switchSheet(sheetName);
                tabsContainer.appendChild(tab);
            });

            archiveContent.appendChild(tabsContainer);

            // Create sheet contents
            sheetNames.forEach((sheetName, index) => {
                const sheetDiv = document.createElement('div');
                sheetDiv.id = `sheet-${sheetName}`;
                sheetDiv.className = `sheet-content ${index === 0 ? 'active' : ''}`;
                sheetDiv.innerHTML = renderTable(sheets[sheetName]);
                archiveContent.appendChild(sheetDiv);
            });
        }

        function switchSheet(sheetName) {
            // Hide all sheets
            document.querySelectorAll('.sheet-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.sheet-tab').forEach(el => el.classList.remove('active'));

            // Show selected sheet
            document.getElementById(`sheet-${sheetName}`).classList.add('active');
            event.target.classList.add('active');
        }

        function renderTable(data) {
            if (!data || data.length === 0) {
                return '<p>No data in this sheet.</p>';
            }

            const headers = Object.keys(data[0]);
            let html = '<table class="data-table"><thead><tr>';
            
            headers.forEach(header => {
                html += `<th>${escapeHtml(header)}</th>`;
            });
            html += '</tr></thead><tbody>';

            data.forEach(row => {
                html += '<tr>';
                headers.forEach(header => {
                    const value = row[header];
                    html += `<td>${escapeHtml(value !== null && value !== undefined ? value : '')}</td>`;
                });
                html += '</tr>';
            });

            html += '</tbody></table>';
            return html;
        }

        function escapeHtml(text) {
            const map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return String(text).replace(/[&<>"']/g, m => map[m]);
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

@app.route('/api/load-excel', methods=['POST'])
def load_excel():
    try:
        data = request.json
        url = data.get('url', '').strip()
        
        if not url:
            return jsonify({"error": "URL is required"}), 400
        
        # Convert Dropbox sharing URL to direct download URL
        if 'dropbox.com' in url:
            # Replace the end parameter from ?dl=0 to ?dl=1 for direct download
            if '?dl=0' in url:
                url = url.replace('?dl=0', '?dl=1')
            elif '?dl=1' not in url:
                url = url + '?dl=1'
        
        # Fetch the file
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        # Parse Excel file
        excel_file = io.BytesIO(response.content)
        xls = pd.ExcelFile(excel_file)
        
        sheets_data = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            # Convert DataFrame to list of dictionaries, handling NaN values
            records = df.where(pd.notna(df), None).to_dict('records')
            sheets_data[sheet_name] = records
        
        return jsonify({"sheets": sheets_data})
    
    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to fetch file: {str(e)}"}), 400
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel file: {str(e)}"}), 400

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