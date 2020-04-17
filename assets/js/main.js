let sqldb;
let _target = document.getElementById("upload-box");
let date_columns = ["End Project Date", "Testing and Commissioning date", "Project Start date", "Job Creation Date"];

/** Alerts **/
let _badfile = function () {
    alertify.alert('This file does not appear to be a valid Excel file.', function () { });
};

let _pending = function () {
    alertify.alert('Please wait until the current file is processed.', function () { });
};

let _large = function (len, cb) {
    alertify.confirm("This file is " + len + " bytes and may take a few moments.  Your browser may lock up during this process.  Shall we play?", cb);
};

let _failed = function (e) {
    console.log(e, e.stack);
    alertify.alert('We unfortunately dropped the ball here.  Please test the file using the <a href="/js-xlsx/">raw parser</a>.  If there are issues with the file processor, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can make things right.', function () { });
};

let _onsheet = function (json, cols) {
    /* add header row for table */
    if (!json) json = [];
    json.unshift(function (head) { let o = {}; for (i = 0; i != head.length; ++i) o[head[i]] = head[i]; return o; }(cols));
};

function getTable(stmt) {
    return new Promise((resolve, reject) => {
        if (!stmt) reject();

        if (stmt.indexOf(";") > -1) stmt = stmt.substr(0, stmt.indexOf(";"));
        sqldb.transaction(function (tx) {
            tx.executeSql(stmt, [], function (tx, result) {
                if (!result || result.length === 0) reject();

                if (result.rows.length != 0) {
                    let r = result.rows.item(0);
                    let cols = Object.keys(r);
                    let json = [];
                    for (let i = 0; i < result.rows.length; ++i) {
                        r = result.rows.item(i);
                        let o = {};
                        cols.forEach(function (x) { o[x] = r[x]; });
                        json.push(o);
                    }
                    resolve(json);
                }
                else reject();
            }, function (tx, e) {
                reject(e + "\n" + (e.message || "") + "\n" + (e.stack || ""));
            });
        });
    });
}

function getData(stmt) {
    return new Promise((resolve, reject) => {
        if (!stmt) reject();

        if (stmt.indexOf(";") > -1) stmt = stmt.substr(0, stmt.indexOf(";"));
        sqldb.transaction(function (tx) {
            tx.executeSql(stmt, [], function (tx, result) {
                if (!result || result.length === 0) reject();
                resolve(Object.values(result.rows.item(0))[0]);
            }, function (tx, e) {
                reject(e + "\n" + (e.message || "") + "\n" + (e.stack || ""));
            });
        });
    });
}

function prepstmt(s) {
    sqldb.transaction(function (tx) { tx.executeSql(s, []); });
}

function initDatabase(ws, sname) {
    /* Get sheet range */
    if (!ws || !ws['!ref']) return;
    let range = XLSX.utils.decode_range(ws['!ref']);
    if (!range || !range.s || !range.e || range.s > range.e) return;
    let R = range.s.r, C = range.s.c;

    /* Generate headers */
    let names = new Array(range.e.c - range.s.c + 1);
    for (C = range.s.c; C <= range.e.c; ++C) {
        let addr = XLSX.utils.encode_cell({ c: C, r: R });
        names[C - range.s.c] = ws[addr] ? ws[addr].v : XLSX.utils.encode_col(C);
    }
    sqldb.columns = names;

    /* De-duplicate headers */
    for (let i = 0; i < names.length; ++i) if (names.indexOf(names[i]) < i)
        for (let j = 0; j < names.length; ++j) {
            let _name = names[i] + "_" + (j + 1);
            if (names.indexOf(_name) > -1) continue;
            names[i] = _name;
        }

    /* Guess column types */
    let types = new Array(range.e.c - range.s.c + 1);
    for (C = range.s.c; C <= range.e.c; ++C) {
        let seen = {}, _type = "";

        for (R = range.s.r + 1; R <= range.e.r; ++R) {
            seen[(ws[XLSX.utils.encode_cell({ c: C, r: R })] || { t: "z" }).t] = true;
        }
        if (seen.s || seen.str) _type = "TEXT";
        else if (seen.n + seen.b + seen.d + seen.e > 1) _type = "TEXT";
        else switch (true) {
            case seen.b:
            case seen.n: _type = "REAL"; break;
            case seen.e: _type = "TEXT"; break;
            case seen.d: _type = "TEXT"; break;
        }

        if (date_columns.includes(names[C - range.s.c]))
            types[C - range.s.c] = "REAL";
        else
            types[C - range.s.c] = _type || "TEXT";
    }

    /* update list */
    let ss = ""
    names.forEach(function (n) { if (n) ss += "`" + n + "`<br />"; });

    /* create table */
    prepstmt("DROP TABLE IF EXISTS `" + sname + "`");
    prepstmt("CREATE TABLE `" + sname + "` (" + names.map(function (n, i) { return "`" + n + "` " + (types[i] || "TEXT"); }).join(", ") + ");");

    /* insert data */
    for (R = range.s.r + 1; R <= range.e.r; ++R) {
        let fields = [], values = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
            let cell = ws[XLSX.utils.encode_cell({ c: C, r: R })];
            if (!cell) continue;
            fields.push("`" + names[C - range.s.c] + "`");

            let val = cell.v;

            if (date_columns.includes(names[C - range.s.c]))
                val = Date.parse(cell.w) / 1000;

            switch (types[C - range.s.c]) {
                case 'REAL': if (cell.t == 'b' || typeof val == 'boolean') val = +val; break;
                default: val = '"' + val.toString().replace(/"/g, '""') + '"';
            }
            values.push(val);
        }

        prepstmt("INSERT INTO `" + sname + "` (" + fields.join(", ") + ") VALUES (" + values.join(",") + ");");
    }

}

function prepareDatabase(wb) {
    return new Promise((resolve, reject) => {
        smtp = "";
        if (wb.Sheets && wb.Sheets.Data && wb.Sheets.Metadata)
            smtp = "SELECT Format, Importance as Priority, Data.Code, css_class FROM Data JOIN Metadata ON Metadata.code = Data.code WHERE Importance < 3";
        else smtp = "SELECT * FROM `" + wb.SheetNames[0] + "` LIMIT 30";

        if (typeof openDatabase === 'undefined') {
            sqldiv.innerHTML = '<div class="error"><b>*** WebSQL not available.  Consider using Chrome or Safari ***</b></div>';
            return;
        }

        sqldb = openDatabase('sheetjs', '1.0', 'db', 3 * 1024 * 1024);
        wb.SheetNames.forEach(function (s) { initDatabase(wb.Sheets[s], s); });
        sqldb.title = wb.SheetNames[0];
        resolve();
    });
}

const COLORS = [
    '#e3e3e3',
    '#4acccd',
    '#fcc468',
    '#ef8157',
    '#468966',
    '#FFF0A5',
    '#FFB03B',
    '#B64926',
    '#8E2800'
];

DropSheet({
    drop: _target,
    on: {
        wb: function (wb) {
            prepareDatabase(wb).then(() => {
                _target.parentNode.remove();

                //Number of ongoing projects
                getData('select count(*) from "Job List" where "Job completed" is "No" and "Remarks" is not "Cancelled" and "Project Start date" <= ' + Math.floor(Date.now() / 1000)).then(v => $("#number-of-ongoing-project").text(v));

                //Number of projects delayed
                getData('select count(*) from "Job List" where "Job completed" = "No" and "End Project Date" IS NOT NULL and "End Project Date" < ' + Math.floor(Date.now() / 1000)).then(v => $("#number-of-delayed-project").text(v));

                //Numbers of projects cancelled
                getData('select count(*) from "Job List" where "Remarks" = "Cancelled"').then(v => $("#number-of-cancelled-project").text(v));

                //Number of projects completed within 6 months
                getData('select count(*) from "Job List" where "Job completed" = "Yes" and "End Project Date" - "Project Start date" < ' + 60 * 60 * 24 * 30 * 6).then(v => $("#number-of-completed-project").text(v));

                //Number of projects completed within x months
                $("#number-of-completed-project-date-range-sellector").change((e) => getData('select count(*) from "Job List" where "Job completed" = "Yes" and "End Project Date" - "Project Start date" < ' + 60 * 60 * 24 * 30 * e.currentTarget.value).then(v => $("#number-of-completed-project").text(v)));

                //Number of on-going projects - days calculated till date, from project start date. Rank of the project that is taking most days and months.
                //Days between job creation date and project start date.
                getTable('select * from "Job List" where "Job completed" is "No" and "Remarks" is not "Cancelled" and "Project Start date" <= ' + Math.floor(Date.now() / 1000)).then(v => {
                    let values = [];

                    v.forEach(e => {
                        let t = [];

                        t.push(e["No."]);
                        t.push(e["Job type"]);
                        t.push(e["Engineer Name"]);
                        t.push(Math.floor((Math.floor(Date.now() / 1000) - e["Project Start date"]) / 86400) + " days");
                        t.push(Math.floor((e["Project Start date"] - e["Job Creation Date"]) / 86400) + " days");

                        values.push(t);
                    });

                    $('#projects').DataTable({
                        data: values,
                        columns: [
                            { title: "No" },
                            { title: "Job type" },
                            { title: "Engineer Name" },
                            { title: "Running Time" },
                            { title: "Starting Time" }
                        ]
                    });

                });

                //Engineers who took lesser time to start a project (project start date - job creation date)
                getTable('select "Engineer Name", avg("Project Start date"-"Job Creation Date") as Count from "Job List" GROUP BY "Engineer Name" ORDER BY "Project Start date"-"Job Creation Date"').then(v => {
                    new Chart($("#quick-engineers"), {
                        type: 'pie',
                        data: {
                            labels: $.map(v, (n) => n["Engineer Name"]),
                            datasets: [{
                                borderWidth: 0,
                                pointRadius: 0,
                                pointHoverRadius: 0,
                                backgroundColor: COLORS,
                                data: $.map(v, (n) => n["Count"] / 86400)
                            }]
                        },

                        options: {
                            pieceLabel: {
                                render: 'percentage',
                                precision: 1
                            },
                            legend: {
                                position: "bottom"
                            }
                        }
                    });
                });

                // Number of projects per engineer
                getTable('select "Engineer Name", count(*) as Count from "Job List" GROUP BY "Engineer Name"').then(v => {
                    new Chart($("#number-of-projects-per-engineer"), {
                        type: 'pie',
                        data: {
                            labels: $.map(v, (n) => n["Engineer Name"]),
                            datasets: [{
                                borderWidth: 0,
                                pointRadius: 0,
                                pointHoverRadius: 0,
                                backgroundColor: COLORS,
                                data: $.map(v, (n) => n["Count"])
                            }]
                        },

                        options: {
                            pieceLabel: {
                                render: 'percentage',
                                precision: 1
                            },
                            legend: {
                                position: "bottom"
                            }
                        }
                    });
                });

                //Engineers ranking on days took to complete a project.
                getTable('select "Engineer Name", avg("End Project Date"-"Job Creation Date") as Rank from "Job List" where "End Project Date" IS NOT NULL GROUP BY "Engineer Name" ORDER BY "End Project Date"-"Job Creation Date"').then(v => console.log(v));

                // Total project value
                getData('select sum("Total contract value") from "Job List"').then(v => $("#total-project-value").val(v));

                // Individual engineer project value
                getTable('select "Engineer Name", sum("Total contract value") as Value from "Job List" GROUP BY "Engineer Name"').then(v => {
                    new Chart($("#individual-engineer-project-value"), {
                        type: 'pie',
                        data: {
                            labels: $.map(v, (n) => n["Engineer Name"]),
                            datasets: [{
                                borderWidth: 0,
                                pointRadius: 0,
                                pointHoverRadius: 0,
                                backgroundColor: COLORS,
                                data: $.map(v, (n) => n["Value"])
                            }]
                        },

                        options: {
                            pieceLabel: {
                                render: 'percentage',
                                precision: 1
                            },
                            legend: {
                                position: "bottom"
                            }
                        }
                    });
                });
            });
        }
    },
    errors: {
        badfile: _badfile,
        pending: _pending,
        failed: _failed,
        large: _large
    }
});
