<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>ملخص مخازن القطاع</title>

<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

<script src="https://cdn.jsdelivr.net/npm/exceljs/dist/exceljs.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/file-saver/dist/FileSaver.min.js"></script>

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

/* ================= Firebase ================= */
const firebaseConfig = {
  apiKey: "AIzaSyCziOxTUx8lpNbBCJT9SQbebMhYdupw6Dg",
  authDomain: "market-app-4f1ef.firebaseapp.com",
  databaseURL: "https://market-app-4f1ef-default-rtdb.firebaseio.com",
  projectId: "market-app-4f1ef"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.database();

/* ================= إضافة صف ================= */
function addRow() {
    let table = document.querySelector("#sheet tbody");
    let newRow = table.rows[0].cloneNode(true);

    newRow.querySelectorAll("td").forEach((cell, i) => {
        if (i !== 0) cell.innerText = "";
    });

    table.appendChild(newRow);
}

/* ================= حذف صف ================= */
function deleteRow(btn) {
    btn.closest("tr").remove();
}

/* ================= Paste Excel ================= */
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

/* ================= Excel Export ================= */
async function exportExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ملخص مخازن القطاع ");

    const headers = [
        "مسلسل","كود المخزن","اسم المخزن","موقف الساب","مسؤول الساب",
        "قيمة المخزون","قيمة الرواكد الحالية","قيمة تدوير رواكد مشاريع اخرى","المدة المتبقية للمشروع"
    ];

    const headerRow = worksheet.addRow(headers);

    headerRow.eachCell(cell => {
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'2F5597' } };
        cell.font = { bold:true, color:{ argb:'FFFFFF' } };
        cell.alignment = { horizontal:'center' };
        cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });

    document.querySelectorAll("#sheet tbody tr").forEach(row => {
        let rowData = [];
        row.querySelectorAll("td").forEach((cell, index) => {
            if(index !== 0) rowData.push(cell.innerText);
        });

        let newRow = worksheet.addRow(rowData);

        newRow.eachCell(cell => {
            cell.alignment = { horizontal:'center' };

            if (newRow.number % 2 === 0) {
                cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'F2F2F2' } };
            }

            cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
        });
    });

    worksheet.columns = [
        { width:10 },{ width:15 },{ width:25 },{ width:15 },
        { width:20 },{ width:15 },{ width:20 },{ width:25 },{ width:20 }
    ];

    worksheet.autoFilter = {
        from: { row:1, column:1 },
        to: { row:worksheet.rowCount, column:worksheet.columnCount }
    };

    worksheet.views = [
        { state:'frozen', ySplit:1, rightToLeft:true }
    ];

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), "المخازن.xlsx");
}
/* ================= حفظ Firebase ================= */
function saveAll() {
    let data = [];

    document.querySelectorAll("#sheet tbody tr").forEach(tr => {
        let row = [];
        tr.querySelectorAll("td").forEach((td, i) => {
            if (i !== 0) row.push(td.innerText);
        });
        data.push(row);
    });

    db.ref("warehouse").set(data)
    .then(() => alert("تم الحفظ على Firebase ✅"))
    .catch(err => alert("خطأ ❌"));
}

/* ================= تحميل البيانات ================= */
window.onload = function () {
    loadData();
};

function loadData() {
    db.ref("warehouse").on("value", (snapshot) => {
        const data = snapshot.val();

        let tbody = document.querySelector("#sheet tbody");
        tbody.innerHTML = "";

        if (!data) {
            addEmptyRow();
            return;
        }

        // 🔥 مهم: نحولها لأي شكل (Array أو Object)
        let records = Array.isArray(data) ? data : Object.values(data);

        if (records.length === 0) {
            addEmptyRow();
            return;
        }

        records.forEach(row => {
            if (Array.isArray(row)) {
                addRowToTable(row);
            }
        });
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
    function addEmptyRow() {
    let tr = document.createElement("tr");

    tr.innerHTML = `
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
    `;

    document.querySelector("#sheet tbody").appendChild(tr);
}


</script>

</body>
</html>
