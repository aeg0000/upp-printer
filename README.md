<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>UPP Web App - Fix Preview</title>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js"></script>
    <style>
        :root { --primary: #2c3e50; --accent: #3498db; --success: #27ae60; --bg: #f4f7f9; }
        body { font-family: 'Tahoma', sans-serif; background: var(--bg); margin: 0; padding-bottom: 70px; }
        .tab-bar { position: fixed; bottom: 0; width: 100%; background: white; display: flex; box-shadow: 0 -2px 10px rgba(0,0,0,0.1); z-index: 1000; }
        .tab-btn { flex: 1; padding: 15px 5px; border: none; background: none; font-size: 11px; color: #7f8c8d; cursor: pointer; }
        .tab-btn.active { color: var(--accent); border-top: 3px solid var(--accent); font-weight: bold; }
        .content-section { display: none; padding: 20px; box-sizing: border-box; min-height: 100vh; }
        .content-section.active { display: block; }
        .card { background: white; padding: 15px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); margin-bottom: 15px; }
        h3 { margin-top: 0; font-size: 18px; color: var(--primary); border-left: 5px solid var(--accent); padding-left: 12px; margin-bottom: 20px; }
        label { display: block; font-size: 13px; font-weight: bold; margin-bottom: 6px; color: #444; }
        input, select { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 15px; box-sizing: border-box; margin-bottom: 10px; }
        .btn { width: 100%; padding: 15px; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 10px; }
        .btn-primary { background: var(--primary); color: white; }
        .btn-bt { background: var(--success); color: white; }
        .preview-wrapper { display: flex; justify-content: center; padding-top: 10px; }
        .preview-box { 
            background: white; border: 1px solid #000; width: 100mm; height: 100mm; 
            padding: 5mm; box-sizing: border-box; display: flex; flex-direction: column;
            transform: scale(0.7); transform-origin: top center;
        }
        .lbl-header { font-size: 14pt; font-weight: bold; text-align: center; border-bottom: 2px solid #000; padding-bottom: 4px; margin-bottom: 10px; }
        .lbl-row { display: flex; border-bottom: 0.5px solid #eee; padding: 4px 0; font-size: 11pt; }
        .lbl-label { font-weight: bold; min-width: 32mm; }
        .lbl-footer { display: grid; grid-template-columns: 1fr 1.2fr; border-top: 1px solid #000; padding-top: 8px; align-items: center; margin-top: auto; }
        .lbl-bc { width: 100%; text-align: center; margin-top: 8px; }
        #p_barcode { width: 100%; height: 50px; }
        @media print {
            .no-print, .tab-bar { display: none !important; }
            #printSection { display: block !important; }
            @page { size: 100mm 100mm; margin: 0; }
            .print-page { width: 100mm; height: 100mm; padding: 5mm; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; }
        }
    </style>
</head>
<body>

<div id="section1" class="content-section active">
    <h3>📝 1. ข้อมูลสินค้า</h3>
    <div class="card">
        <label>อัปโหลด Excel:</label><input type="file" id="excelFile" accept=".xlsx">
        <label>เลือกลูกค้า:</label><select id="custSelect" onchange="filterProducts()"><option>-- เลือกไฟล์ --</option></select>
        <label>เลือกสินค้า:</label><select id="prodSelect" onchange="updateUI()"><option>-- เลือกลูกค้า --</option></select>
        <div style="display:flex; gap:10px;">
            <div style="flex:1"><label>ช่าง:</label><input type="text" id="worker" placeholder="ชื่อช่าง" oninput="updateUI()"></div>
            <div style="flex:1"><label>เครื่อง:</label><input type="text" id="machine" placeholder="เลขเครื่อง" oninput="updateUI()"></div>
        </div>
        <label>วันที่ผลิต:</label><input type="date" id="mfgDate" onchange="updateUI()">
    </div>
</div>

<div id="section2" class="content-section">
    <h3>⚙️ 2. เครื่องพิมพ์</h3>
    <div class="card">
        <button class="btn btn-bt" onclick="connectBluetooth()">🔵 ค้นหา Bluetooth</button>
        <label style="margin-top:20px;">หัวข้อบริษัท:</label><input type="text" id="headerText" value="บริษัท ยูพีพี อัลทิเมท แพคกิ้ง จำกัด" oninput="updateUI()">
        <label>จำนวนที่พิมพ์:</label><input type="number" id="printQty" value="1" min="1">
        <button class="btn btn-primary" onclick="doPrint()">🖨️ สั่งพิมพ์</button>
    </div>
</div>

<div id="section3" class="content-section">
    <h3>🖼️ 3. Preview</h3>
    <div class="preview-wrapper">
        <div id="mainPreview" class="preview-box">
            <div class="lbl-header" id="ph">บริษัท ยูพีพี อัลทิเมท แพคกิ้ง จำกัด</div>
            <div style="flex-grow:1">
                <div class="lbl-row"><span class="lbl-label">CUSTOMER:</span><span id="p_cust">-</span></div>
                <div class="lbl-row"><span class="lbl-label">PRODUCT:</span><span id="p_prod">-</span></div>
                <div class="lbl-row"><span class="lbl-label">รหัสสินค้า:</span><span id="p_id">-</span></div>
                <div class="lbl-row"><span class="lbl-label">Q.T.Y:</span><span id="p_qty">-</span></div>
                <div class="lbl-row"><span class="lbl-label">ผลิต /MG:</span><span id="p_mfg">-</span></div>
            </div>
            <div class="lbl-footer">
                <div style="font-size: 9pt;">
                    <div><b>ช่าง:</b> <span id="p_wk">-</span></div>
                    <div><b>เครื่อง:</b> <span id="p_mc">-</span></div>
                </div>
                <div style="display:flex; justify-content:center;"><div id="p_qr"></div></div>
            </div>
            <div class="lbl-bc"><svg id="p_barcode"></svg></div>
        </div>
    </div>
</div>

<div class="tab-bar no-print">
    <button class="tab-btn active" onclick="switchTab(1)">📝 ข้อมูล</button>
    <button class="tab-btn" onclick="switchTab(2)">⚙️ พิมพ์</button>
    <button class="tab-btn" onclick="switchTab(3)">🖼️ Preview</button>
</div>
<div id="printSection" style="display:none;"></div>

<script>
    let db = [], activeItem = null;

    function switchTab(n) {
        document.querySelectorAll('.content-section, .tab-btn').forEach(el => el.classList.remove('active'));
        document.getElementById('section'+n).classList.add('active');
        document.querySelectorAll('.tab-btn')[n-1].classList.add('active');
        if(n === 3) updateUI(); // อัปเดตบาร์โค้ดทุกครั้งที่เปิดหน้า Preview
    }

    document.getElementById('excelFile').addEventListener('change', e => {
        const reader = new FileReader();
        reader.onload = ev => {
            const wb = XLSX.read(new Uint8Array(ev.target.result), {type: 'array'});
            db = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {range: 1, defval: ""});
            const custs = [...new Set(db.map(i => i.CUTOMER))].filter(Boolean).sort();
            const sel = document.getElementById('custSelect');
            sel.innerHTML = '<option value="">-- เลือกลูกค้า --</option>';
            custs.forEach(c => sel.add(new Option(c, c)));
        };
        reader.readAsArrayBuffer(e.target.files[0]);
    });

    function filterProducts() {
        const c = document.getElementById('custSelect').value;
        const sel = document.getElementById('prodSelect');
        sel.innerHTML = '<option value="">-- เลือกสินค้า --</option>';
        db.filter(i => i.CUTOMER === c).forEach(item => sel.add(new Option(item['ชื่อสินค้า'], db.indexOf(item))));
    }

    function updateUI() {
        const idx = document.getElementById('prodSelect').value;
        if(idx === "") return;
        activeItem = db[idx];

        document.getElementById('ph').innerText = document.getElementById('headerText').value;
        document.getElementById('p_cust').innerText = activeItem['CUTOMER'] || "-";
        document.getElementById('p_prod').innerText = activeItem['ชื่อสินค้า'] || "-";
        document.getElementById('p_id').innerText = activeItem['รหัสสินค้า'] || "-";
        document.getElementById('p_qty').innerText = (activeItem['Q.T.Y.'] || "0") + " " + (activeItem['หน่วย'] || "");
        document.getElementById('p_wk').innerText = document.getElementById('worker').value || "-";
        document.getElementById('p_mc').innerText = document.getElementById('machine').value || "-";
        document.getElementById('p_mfg').innerText = document.getElementById('mfgDate').value || "-";

        const val = String(activeItem['บาร์โค้ด'] || "0000000000000");
        
        // วาด QR Code
        const qrContainer = document.getElementById('p_qr');
        qrContainer.innerHTML = "";
        new QRCode(qrContainer, { text: val, width: 70, height: 70 });

        // วาด Barcode ด้วยดีเลย์เล็กน้อยเพื่อให้ Element พร้อมใช้งาน
        setTimeout(() => {
            JsBarcode("#p_barcode", val, {
                format: "CODE128",
                width: 2.5,
                height: 40,
                displayValue: true,
                fontSize: 14,
                margin: 0
            });
        }, 100);
    }

    async function connectBluetooth() {
        try {
            const device = await navigator.bluetooth.requestDevice({ acceptAllDevices: true });
            alert("เชื่อมต่อ " + device.name + " สำเร็จ");
        } catch (e) { alert("Bluetooth error: " + e.message); }
    }

    function doPrint() {
        if(!activeItem) return alert("เลือกข้อมูลก่อน");
        const qty = document.getElementById('printQty').value;
        const area = document.getElementById('printSection');
        area.innerHTML = "";
        const val = String(activeItem['บาร์โค้ด']);

        for(let i=0; i<qty; i++) {
            const page = document.createElement('div');
            page.className = 'print-page';
            page.innerHTML = `
                <div class="lbl-header">${document.getElementById('headerText').value}</div>
                <div style="flex-grow:1">
                    <div class="lbl-row"><span class="lbl-label">CUSTOMER:</span><span>${activeItem['CUTOMER']}</span></div>
                    <div class="lbl-row"><span class="lbl-label">PRODUCT:</span><span>${activeItem['ชื่อสินค้า']}</span></div>
                    <div class="lbl-row"><span class="lbl-label">รหัสสินค้า:</span><span>${activeItem['รหัสสินค้า']}</span></div>
                    <div class="lbl-row"><span class="lbl-label">Q.T.Y:</span><span>${activeItem['Q.T.Y.']} ${activeItem['หน่วย']}</span></div>
                    <div class="lbl-row"><span class="lbl-label">ผลิต /MG:</span><span>${document.getElementById('mfgDate').value}</span></div>
                </div>
                <div class="lbl-footer">
                    <div style="font-size: 10pt;">
                        <div><b>ช่าง:</b> ${document.getElementById('worker').value}</div>
                        <div><b>เครื่อง:</b> ${document.getElementById('machine').value}</div>
                    </div>
                    <div style="display:flex; justify-content:center;"><div id="print_qr_${i}"></div></div>
                </div>
                <div class="lbl-bc"><svg id="print_bc_${i}"></svg></div>
            `;
            area.appendChild(page);
            new QRCode(document.getElementById(`print_qr_${i}`), { text: val, width: 80, height: 80 });
            JsBarcode(`#print_bc_${i}`, val, { format: "CODE128", width: 3, height: 50, displayValue: true });
        }
        window.print();
    }
</script>
</body>
</html>
