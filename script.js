/* ============================================================
   Excel Lookup Guide — interactive VLOOKUP demo
   Performs a real lookup in the browser, mimicking Excel's
   VLOOKUP behaviour and highlighting, exactly as explained
   in the tutorial above.
   ============================================================ */

(function () {
    "use strict";

    /* ---- The "spreadsheet" data (Employee table, A1:D8) ---- */
    var EMP = [
        { id: 101, name: "Priya Sharma", dept: "HR", salary: 45000 },
        { id: 102, name: "Rahul Verma", dept: "IT", salary: 62000 },
        { id: 103, name: "Sneha Iyer", dept: "Finance", salary: 58000 },
        { id: 104, name: "Arjun Patel", dept: "IT", salary: 71000 },
        { id: 105, name: "Meera Nair", dept: "Marketing", salary: 53000 },
        { id: 106, name: "Vikram Rao", dept: "Sales", salary: 49000 },
        { id: 107, name: "Divya Menon", dept: "Finance", salary: 66000 }
    ];

    /* Select value (0,1,2) -> which field to return, and the
       matching col_index_num inside the table_array A2:D8. */
    var RETURN_FIELDS = ["name", "dept", "salary"];
    var RETURN_LABELS = ["Name", "Department", "Salary"];
    var TABLE_COLS = [2, 3, 4];

    /* ---- Element handles ---- */
    var idInput = document.getElementById("demo-id");
    var colSelect = document.getElementById("demo-col");
    var body = document.getElementById("demo-grid-body");
    var status = document.getElementById("demo-status");
    var statusBar = document.getElementById("demo-statusbar");
    var fCol = document.getElementById("demo-f-col");

    function inr(n) {
        return "\u20B9" + n.toLocaleString("en-IN");
    }

    /* ---- Build the demo grid once ---- */
    function buildGrid() {
        var html = "";

        /* Row 1 — headers */
        html += '<tr><th class="rownum">1</th>';
        html += '<td class="th-data">Emp ID</td>';
        html += '<td class="th-data">Name</td>';
        html += '<td class="th-data">Department</td>';
        html += '<td class="th-data">Salary</td>';
        html += '<td></td>';
        html += '<td class="th-data">Employee ID</td>';
        html += '<td class="th-data" id="demo-g1">Department</td></tr>';

        /* Rows 2-8 — employee records */
        EMP.forEach(function (e, i) {
            var rowNum = i + 2;
            html += '<tr data-id="' + e.id + '"><th class="rownum">' + rowNum + '</th>';
            html += '<td class="num">' + e.id + '</td>';
            html += '<td>' + e.name + '</td>';
            html += '<td>' + e.dept + '</td>';
            html += '<td class="num">' + inr(e.salary) + '</td>';
            html += '<td></td>';
            html += rowNum === 2
                ? '<td class="num" id="demo-f2">105</td>'
                : '<td></td>';
            html += rowNum === 2
                ? '<td class="num" id="demo-g2">Marketing</td>'
                : '<td></td>';
            html += '</tr>';
        });

        body.innerHTML = html;
    }

    /* ---- Remove highlight classes added by a previous run ---- */
    function clearHighlights() {
        var tds = body.querySelectorAll("td.hl-amber, td.ring-blue");
        for (var i = 0; i < tds.length; i++) {
            tds[i].classList.remove("hl-amber");
            tds[i].classList.remove("ring-blue");
        }
        var ths = document.querySelectorAll("#demo-grid thead th");
        for (var j = 0; j < ths.length; j++) {
            ths[j].classList.remove("hl-green");
        }
    }

    /* ---- Recompute the "formula" ---- */
    function update() {
        var choice = parseInt(colSelect.value, 10);        /* 0, 1 or 2   */
        var field = RETURN_FIELDS[choice];
        var tableCol = TABLE_COLS[choice];                  /* 2, 3 or 4   */
        var raw = idInput.value.trim();
        var id = parseInt(raw, 10);

        /* Update the colour-coded formula pieces + labels */
        fCol.textContent = tableCol;
        document.getElementById("demo-g1").textContent = RETURN_LABELS[choice];

        var f2 = document.getElementById("demo-f2");
        var g2 = document.getElementById("demo-g2");
        f2.textContent = raw === "" ? "" : raw;

        clearHighlights();

        /* Highlight the returned column header (thead index === table col) */
        var headTh = document.querySelectorAll("#demo-grid thead th")[tableCol];
        if (headTh) {
            headTh.classList.add("hl-green");
        }

        /* The lookup itself — this IS the VLOOKUP */
        var row = null;
        for (var k = 0; k < EMP.length; k++) {
            if (EMP[k].id === id) {
                row = EMP[k];
                break;
            }
        }

        if (!row) {
            g2.textContent = "#N/A";
            g2.className = "num err-cell";
            status.classList.add("err");
            status.textContent = raw === ""
                ? "Type an Employee ID (101\u2013107) to run the lookup."
                : "\u274C #N/A \u2014 " + raw + " is not in the first column (A2:A8). VLOOKUP found nothing to return.";
            statusBar.textContent = "#N/A";
            return;
        }

        var value = field === "salary" ? inr(row.salary) : row[field];
        g2.textContent = value;
        g2.className = "num cell-formula";

        /* Highlight the matched row (amber) + matched ID cell (blue ring) */
        var tr = body.querySelector('tr[data-id="' + row.id + '"]');
        if (tr) {
            for (var c = 1; c <= 4; c++) {
                tr.cells[c].classList.add("hl-amber");
            }
            tr.cells[1].classList.add("ring-blue");
        }

        status.classList.remove("err");
        status.innerHTML = "\u2705 VLOOKUP found " + row.id + " \u2192 <strong>" + value + "</strong>";
        statusBar.textContent = "Ready";
    }

    buildGrid();
    idInput.addEventListener("input", update);
    colSelect.addEventListener("change", update);
    update();
})();