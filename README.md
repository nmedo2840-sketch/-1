<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>ملخص مخازن القطاع</title>

<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

<script src="https://cdn.jsdelivr.net/npm/exceljs/dist/exceljs.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/file-saver/dist/FileSaver.min.js"></script>

<!-- Firebase (Compat Version - بدون مشاكل import) -->
<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-database-compat.js"></script>

<style>
body {
    background: #f4f6f9;
    font-family: 'Cairo', sans-serif;
    direction: rtl;
}

.table-container {
    width: 95%;
    margin: 30px auto;
    overflow-x: auto;
}

table {
    border-collapse: collapse;
    width: 100%;
}

th {
    background: #2F5597;
    color: white;
    text-align: center;
    padding: 10px;
}

td {
    height: 35px;
    border: 1px solid #ccc;
    text-align: center;
}

td:focus {
    outline: 2px solid #2F5597;
    background: #eef4ff;
}

.title {
    text-align: center;
    font-size: 22px;
    margin-top: 20px;
    font-weight: bold;
}

.delete-btn {
    background: red;
    color: white;
    border: none;
    padding: 5px 10px;
}
</style>
</head>

<body>

<div class="title">ملخص مخازن القطاع</div>

<div class="text-center mb-3">
    <button class="btn btn-success" onclick="addRow()">➕ إضافة صف</button>
    <button class="btn btn-primary" onclick="exportExcel()">📥 Excel</button>
    <button class="btn btn-warning" onclick="saveAll()">☁ حفظ Firebase</button>
</div>

<div class="table-container">
<table id="sheet">
<thead>
<tr>
    <th>حذف</th>
    <th>مسلسل</th>
    <th>كود المخزن</th>
    <th>اسم المخزن</th>
    <th>موقف الساب</th>
    <th>مسؤول الساب</th>
    <th>قيمة المخزون</th>
    <th>قيمة الرواكد</th>
    <th>تدوير رواكد</th>
    <th>مدة المشروع</th>
</tr>
</thead>

<tbody>
<tr>
    <td><button class="delete-btn" onclick="deleteRow(this)">❌</button></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
    <td contenteditable="true"></td>
</tr>
</tbody>
</table>
</div>

<script>

/* =========================
   🔥 Firebase Config
========================= */
const firebaseConfig = {
  apiKey: "AIzaSyCziOxTUx8lpNbBCJT9SQbebMhYdupw6Dg",
  authDomain: "market-app-4f1ef.firebaseapp.com",
  databaseURL: "https://market-app-4f1ef-default-rtdb.firebaseio.com",
  projectId: "market-app-4f1ef"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.database();

/* =========================
   ➕ إضافة صف
========================= */
function addRow() {
    let table = document.querySelector("#sheet tbody");
    let newRow = table.rows[0].cloneNode(true);

    newRow.querySelectorAll("td").forEach((cell, i) => {
        if (i !== 0) cell.innerText = "";
    });

    table.appendChild(newRow);
}

/* =========================
   ❌ حذف صف
========================= */
function deleteRow(btn) {
    btn.closest("tr").remove();
}

/* =========================
   📋 Paste من Excel
========================= */
document.addEventListener("paste", function (e) {
    let data = (e.clipboardData || window.clipboardData).getData("text");
    let rows = data.split("\n");
    let table = document.querySelector("#sheet tbody");

    rows.forEach(r => {
        if (!r.trim()) return;

        let cols = r.split("\t");
        let tr = document.createElement("tr");

        tr.innerHTML = `<td><button class="delete-btn" onclick="deleteRow(this)">❌</button></td>`;

        cols.forEach(c => {
            let td = document.createElement("td");
            td.contentEditable = true;
            td.innerText = c;
            tr.appendChild(td);
        });

        table.appendChild(tr);
    });

    e.preventDefault();
});

/* =========================
   📥 Excel Export
========================= */
async function exportExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("المخازن");

    let headers = ["مسلسل","كود","اسم","موقف","مسؤول","مخزون","رواكد","تدوير","مدة"];
    ws.addRow(headers);

    document.querySelectorAll("#sheet tbody tr").forEach(tr => {
        let row = [];
        tr.querySelectorAll("td").forEach((td, i) => {
            if (i !== 0) row.push(td.innerText);
        });
        ws.addRow(row);
    });

    const buffer = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), "المخازن.xlsx");
}

/* =========================
   ☁ حفظ Firebase
========================= */
function saveAll() {
    let data = [];

    document.querySelectorAll("#sheet tbody tr").forEach(tr => {
        let row = [];
        tr.querySelectorAll("td").forEach((td, i) => {
            if (i !== 0) row.push(td.innerText);
        });
        data.push(row);
    });

   db.ref("warehouse").push({
    rows: data,
    time: Date.now()
});
    .then(() => alert("تم الحفظ على Firebase ✅"))
    .catch(err => alert("خطأ في الحفظ ❌"));
}
window.onload = function () {
    loadData();
};

function loadData() {
    db.ref("warehouse").once("value", (snapshot) => {
        const data = snapshot.val();

        let tbody = document.querySelector("#sheet tbody");
        tbody.innerHTML = ""; // امسح الجدول

        if (!data) return;

        // لو البيانات محفوظة ك array (من set)
        if (Array.isArray(data)) {
            data.forEach(row => {
                addRowToTable(row);
            });
        } 
        // لو محفوظة ك objects (push)
        else {
            Object.values(data).forEach(item => {
                if (item.rows) {
                    addRowToTable(item.rows);
                }
            });
        }
    });
}

function addRowToTable(rowData) {
    let tr = document.createElement("tr");

    tr.innerHTML = `<td><button class="delete-btn" onclick="deleteRow(this)">❌</button></td>`;

    rowData.forEach(val => {
        let td = document.createElement("td");
        td.contentEditable = true;
        td.innerText = val;
        tr.appendChild(td);
    });

    document.querySelector("#sheet tbody").appendChild(tr);
}
</script>

</body>
</html>

</script>

</body>
</html>
