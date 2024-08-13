<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SHACL Form Generator</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        .form-container {
            margin-bottom: 20px;
            padding: 20px;
            border: 1px solid #ccc;
            border-radius: 8px;
        }
        .form-container h2 {
            margin-top: 0;
        }
        .form-group {
            margin-bottom: 15px;
        }
        .form-group label {
            display: block;
            margin-bottom: 5px;
        }
        .form-group input, .form-group select {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
        }
        .buttons {
            margin-top: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        table, th, td {
            border: 1px solid black;
        }
        th, td {
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #f1f1f1;
            cursor: pointer;
        }
    </style>
    <!-- Firebase App (the core Firebase SDK) -->
    <script src="https://www.gstatic.com/firebasejs/8.6.1/firebase-app.js"></script>
    <!-- Add Firebase products that you want to use -->
    <script src="https://www.gstatic.com/firebasejs/8.6.1/firebase-firestore.js"></script>
    <!-- XLSX library -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js"></script>
</head>
<body>
    <!--
    <label>Update Cell Data</label> 
    <br/>
    <textarea id="w3review" name="w3review" rows="4" cols="50"></textarea>
    <br />
    <button onclick="updateCell()" type="button">Submit</button>

    <h2>Previous Submissions</h2>
    <table id="submissionsTable">
        <thead>
            <tr>
                <th>ID</th>
                <th>JSON</th>
            </tr>
        </thead>
        <tbody>
            Table rows will be inserted here dynamically
        </tbody>
    </table> -->

    <div id="forms"></div>

    <input type="file" id="fileInput" accept=".jsonld" />
    <input type="file" id="formFileInput" accept=".json" />

    <div class="buttons">
        <button id="submitButton">Submit Form</button>
        <button id="exportButton">Export Completed Form</button>
        <button id="exportExcelButton">Export Table to Excel</button>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/jsonld/1.8.1/jsonld.min.js"></script>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script>
        let currentFormData = {};
        let expandedShapes = [];

        Office.onReady(function (info) {
            if (info.host === Office.HostType.Excel) {
                Excel.run(function (context) {
                    var sheet = context.workbook.worksheets.getActiveWorksheet();
                    sheet.onSelectionChanged.add(handleSelectionChange);
                    return context.sync();
                }).catch(function (error) {
                    console.error(error);
                });
            }
        });

 	function handleSelectionChange(event) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var selectedRange = sheet.getRange(event.address);

        // Load the properties we need to check if it's a single cell selection
        selectedRange.load(["address", "rowCount", "columnCount", "values"]);

        return context.sync().then(function () {
            console.log("Selected Range Address:", selectedRange.address);
            console.log("Row Count:", selectedRange.rowCount, "Column Count:", selectedRange.columnCount);

            // Ensure that the selected range is a single cell
            if (selectedRange.rowCount === 1 && selectedRange.columnCount === 1) {
                var row = selectedRange.getRow();
		console.log("row", row)
		console.log("row+1", row + 1)
		console.log("address", "M" + selectedRange.address[7])
                var columnMRange = sheet.getRange("M" + (selectedRange.address[7])); // Get the corresponding value in Column M
		console.log("Column Range", columnMRange)
                columnMRange.load(["values"]);
		//console.log("values", columnMRange.values)
                return context.sync().then(function () {
		    console.log("test")
                    console.log("Column M Value:", columnMRange.values);
                    var shapeId = columnMRange.values[0][0];  // Value in column M
                    if (shapeId && selectedRange.values[0][0]) {
                        var jsonData = JSON.parse(selectedRange.values[0][0]);
                        console.log("Loaded JSON Data:", jsonData);
                        //populateForm(jsonData);
			if (shapeId.toLowerCase() === 'boiler') {
                        renderSpecificShape('tw:Boiler_defaultShape');
                    } else if (shapeId.toLowerCase() === 'motor') {
                        renderSpecificShape('tw:Motor_defaultShape');
                    }	
                    }
                });
            } else {
                console.error("Please select a single cell.");
            }
        });
    });
}
		
	function renderSpecificShape(shapeId) {
    const formsContainer = document.getElementById('forms');
    formsContainer.innerHTML = '';
	console.log(expandedShapes)
    const matchingShapes = expandedShapes.filter(node => {
        const shapeLabel = node['http://www.w3.org/2000/01/rdf-schema#label']?.[0]?.['@value'].toLowerCase();
    	return shapeLabel === shapeId.toLowerCase();

    });

    if (matchingShapes.length > 0) {
        const shape = matchingShapes[0];
        const formContainer = document.createElement('div');
        formContainer.classList.add('form-container');
        const shapeName = shape['http://www.w3.org/2000/01/rdf-schema#label']?.[0]?.['@value'] || shape['@id'];
        const title = document.createElement('h2');
        title.textContent = shapeName;
        formContainer.appendChild(title);

        const properties = shape['http://www.w3.org/ns/shacl#property'];
        if (properties) {
            properties.forEach(propertyRef => {
                const property = expandedShapes.find(node => node['@id'] === propertyRef['@id']);
                if (property) {
                    const path = property['http://www.w3.org/ns/shacl#path']?.[0]?.['@id'] || '';
                    const name = property['http://www.w3.org/ns/shacl#name']?.[0]?.['@value'] || path;
                    const formGroup = document.createElement('div');
                    formGroup.classList.add('form-group');

                    const label = document.createElement('label');
                    label.setAttribute('for', path);
                    label.textContent = name;
                    formGroup.appendChild(label);

                    const shIn = property['http://www.w3.org/ns/shacl#in'];
                    if (shIn && shIn[0]['@list']) {
                        const select = document.createElement('select');
                        select.setAttribute('id', path);
                        select.setAttribute('name', path);
                        shIn[0]['@list'].forEach(optionValue => {
                            const option = document.createElement('option');
                            option.setAttribute('value', optionValue['@value']);
                            option.textContent = optionValue['@value'];
                            select.appendChild(option);
                        });
                        formGroup.appendChild(select);
                    } else {
                        const class_name = property['http://www.w3.org/ns/shacl#class']?.[0]?.['@id'] || '';
                        const input = document.createElement('input');
                        input.setAttribute('type', 'text');
                        input.setAttribute('id', class_name);
                        input.setAttribute('name', class_name);
                        formGroup.appendChild(input);
                    }

                    formContainer.appendChild(formGroup);
                }
            });
        }

        formsContainer.appendChild(formContainer);
    } else {
        formsContainer.innerHTML = `<p>No matching form found for ${shapeId}</p>`;
    }
}
	
        function writeData(data) {
            Excel.run(function (context) {
                var sheet = context.workbook.worksheets.getActiveWorksheet();
                var cell = context.workbook.getSelectedRange();
                cell.values = [[JSON.stringify(data)]];
                return context.sync();
            }).catch(function (error) {
                console.error(error);
            });
        }

        function handleFileSelect(event) {
            const file = event.target.files[0];
            const reader = new FileReader();
            reader.onload = function(event) {
                const jsonldContent = JSON.parse(event.target.result);
                parseSHACL(jsonldContent);
            };
            reader.readAsText(file);
        }

        function updateCell() {
            currentFormData = JSON.parse(document.getElementById('w3review').value);
            populateForm(currentFormData);
        }

        function handleFormFileSelect(event) {
            const file = event.target.files[0];
            const reader = new FileReader();
            reader.onload = function(event) {
                currentFormData = JSON.parse(event.target.result);
                populateForm(currentFormData);
            };
            reader.readAsText(file);
        }

        function parseSHACL(jsonldContent) {
            jsonld.expand(jsonldContent, (err, expanded) => {
                if (err) {
                    console.error(err);
                    return;
                }
                expandedShapes = expanded.filter(node => node['@type'] && node['@type'].includes('http://www.w3.org/ns/shacl#NodeShape'));
            });
        }

        function populateForm(data) {
            for (const key in data) {
                const element = document.getElementById(key);
                if (element) {
                    element.value = data[key];
                    const event = new Event('change', { bubbles: true });
                    element.dispatchEvent(event);
                }
            }
        }

        function exportForm() {
            const forms = document.querySelectorAll('.form-container');
            let formData = {};

            forms.forEach(form => {
                const inputs = form.querySelectorAll('input, select');
                inputs.forEach(input => {
                    formData[input.name] = input.value;
                });
            });

            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(formData));
            const downloadAnchorNode = document.createElement('a');
            downloadAnchorNode.setAttribute("href", dataStr);
            downloadAnchorNode.setAttribute("download", "form_data.json");
            document.body.appendChild(downloadAnchorNode);
            downloadAnchorNode.click();
            downloadAnchorNode.remove();
        }

        async function submitForm() {
            const forms = document.querySelectorAll('.form-container');
            let formData = {};

            forms.forEach(form => {
                const inputs = form.querySelectorAll('input, select');
                inputs.forEach(input => {
                    formData[input.name] = input.value;
                });
            });

            try {
                writeData(formData);
                //addSubmissionToTable(docRef.id, formData);
            } catch (error) {
                console.error('Error adding document: ', error);
            }
        }

        async function loadSubmissions() {
            const tableBody = document.getElementById('submissionsTable').querySelector('tbody');
            tableBody.innerHTML = '';

            try {
                const querySnapshot = await db.collection('submissions').get();
                querySnapshot.forEach(doc => {
                    addSubmissionToTable(doc.id, doc.data());
                });
            } catch (error) {
                console.error('Error getting documents: ', error);
            }
        }

        function addSubmissionToTable(id, data) {
            const tableBody = document.getElementById('submissionsTable').querySelector('tbody');
            const row = document.createElement('tr');
            row.addEventListener('click', () => populateForm(data));
            const idCell = document.createElement('td');
            idCell.textContent = id;
            const jsonCell = document.createElement('td');
            jsonCell.textContent = JSON.stringify(data);
            row.appendChild(idCell);
            row.appendChild(jsonCell);
            tableBody.appendChild(row);
        }

        function exportTableToExcel() {
            const table = document.getElementById('submissionsTable');
            const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
            XLSX.writeFile(wb, 'submissions.xlsx');
        }
    </script>
</body>
</html>

