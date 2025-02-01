import React, { useState } from "react";
import { read, utils } from "xlsx";

function ExcelImportExport() {
    const [employees, setEmployees] = useState([]);
    const [sheets, setSheets] = useState([]);
    const [selectedSheet, setSelectedSheet] = useState("");
    const [columns, setColumns] = useState([]);
    const [file, setFile] = useState(null);

    const REQUIRED_COLUMNS = ["Name", "Amount", "Date", "Verified"];

    const handleFileSelection = (event) => {
        const files = event.target.files;
        if (files.length) {
            const selectedFile = files[0];
            if (selectedFile.size > 2 * 1024 * 1024) {
                alert("File size exceeds 2MB. Please upload a smaller file.");
                return;
            }
            setFile(selectedFile);
        }
    };

    const handleImport = () => {
        if (!file) {
            alert("Please select a file first.");
            return;
        }

        const reader = new FileReader();
        reader.onload = (event) => {
            const wb = read(event.target.result);
            const sheetNames = wb.SheetNames;
            if (sheetNames.length) {
                setSheets(sheetNames);
                setSelectedSheet(sheetNames[0]);
                updateTable(wb, sheetNames[0]);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    const formatDate = (value) => {
        if (!value) return "";
        if (typeof value === "number") {
            const date = new Date((value - 25569) * 86400 * 1000);
            return new Intl.DateTimeFormat("en-GB").format(date);
        }
        const parsedDate = new Date(value);
        if (!isNaN(parsedDate)) {
            return new Intl.DateTimeFormat("en-GB").format(parsedDate);
        }
        return value;
    };

    const updateTable = (workbook, sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const rows = utils.sheet_to_json(sheet, { header: 1 });

        if (rows.length < 2) {
            alert(`No data found in sheet: ${sheetName}`);
            setColumns([]);
            setEmployees([]);
            return;
        }

        const fileColumns = rows[0];
        const missingColumns = REQUIRED_COLUMNS.filter(col => !fileColumns.includes(col));
        if (missingColumns.length) {
            alert(`Missing required columns in sheet ${sheetName}: ${missingColumns.join(", ")}`);
            setEmployees([]);
            setColumns([]);
            return;
        }

        let validRows = [];
        let errors = [];

        rows.slice(1).forEach((row, rowIndex) => {
            const rowData = Object.fromEntries(fileColumns.map((col, i) => [col, row[i]]));
            const hasAllRequired = REQUIRED_COLUMNS.every(col => rowData[col] !== undefined && rowData[col] !== "");
            const isAmountValid = !isNaN(rowData.Amount) && Number(rowData.Amount) > 0;

            if (hasAllRequired && isAmountValid) {
                validRows.push({ ...rowData, Date: formatDate(rowData.Date) });
            } else {
                let errorDescription = [];
                if (!hasAllRequired) errorDescription.push("Missing required fields");
                if (!isAmountValid) errorDescription.push("Amount must be a number greater than zero");
                errors.push(`Sheet: ${sheetName}, Row: ${rowIndex + 2}, Errors: ${errorDescription.join(", ")}`);
            }
        });

        if (errors.length) {
            alert(`The following errors were found:\n${errors.join("\n")}`);
        }

        setColumns(fileColumns);
        setEmployees(validRows);
    };

    const handleSheetChange = (event) => {
        setSelectedSheet(event.target.value);
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const wb = read(e.target.result);
                updateTable(wb, event.target.value);
            };
            reader.readAsArrayBuffer(file);
        }
    };

    const handleDelete = (index) => {
        if (window.confirm("Are you sure you want to delete this row?")) {
            setEmployees(employees.filter((_, i) => i !== index));
        }
    };

    return (
        <div className="container">
            <div className="row">
                <input type="file" name="file" className="btn btn-primary" onChange={handleFileSelection} />
                <button className="btn btn-success ml-2" onClick={handleImport} disabled={!file}>Import Data</button>
            </div>
            {sheets.length > 0 && (
                <div className="row mt-3">
                    <label>Select Sheet: </label>
                    <select className="form-control w-25" value={selectedSheet} onChange={handleSheetChange}>
                        {sheets.map((sheet, index) => (
                            <option key={index} value={sheet}>{sheet}</option>
                        ))}
                    </select>
                </div>
            )}
            <div className="row mt-3">
                <table className="table">
                    <thead>
                        <tr>
                            {columns.map((col, index) => <th key={index}>{col}</th>)}
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {employees.length ? (
                            employees.map((emp, index) => (
                                <tr key={index}>
                                    {columns.map((col, colIndex) => (
                                        <td key={colIndex}>{col === "Date" ? formatDate(emp[col]) : emp[col]}</td>
                                    ))}
                                    <td>
                                        <button className="btn btn-danger btn-sm" onClick={() => handleDelete(index)}>Delete</button>
                                    </td>
                                </tr>
                            ))
                        ) : (
                            <tr>
                                <td colSpan={columns.length + 1} className="text-center">No Data</td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

export default ExcelImportExport;
