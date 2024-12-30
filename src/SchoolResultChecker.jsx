import React, { useState, useMemo, useEffect } from 'react';
import { useTable } from 'react-table';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import './SchoolResultChecker.css';
import schoollogo from "./schoollogo.jpg"
import { click } from '@testing-library/user-event/dist/click';

export default function SchoolResultChecker() {

    // useEffect(() => {
    //     // Condition 1: Warn on page reload
    //     const handleBeforeUnload = (event) => {
    //         event.preventDefault();
    //         event.returnValue = "If you reload, you may lose your data...";
    //     };

    //     window.addEventListener('beforeunload', handleBeforeUnload);

    //     // Condition 2: Check for slow internet connection
    //     const checkInternetSpeed = () => {
    //         const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    //         if (connection && connection.effectiveType) {
    //             // Adjust "slow" threshold based on your requirement
    //             const slowConnections = ['slow-2g', '2g', '3g'];
    //             if (slowConnections.includes(connection.effectiveType)) {
    //                 alert("Your internet is too slow; you may lose data if you continue.");
    //             }
    //         }
    //     };

    //     // Run once on mount
    //     checkInternetSpeed();

    //     // Cleanup event listener on unmount
    //     return () => {
    //         window.removeEventListener('beforeunload', handleBeforeUnload);
    //     };
    // }, []);

    const [students, setStudents] = useState([]);
    const [newStudent, setNewStudent] = useState({ name: '' });
    const [subjects, setSubjects] = useState(['विषय १']);
    const [newSubjects, setNewSubjects] = useState('');

    const calculateGrade = (total) => {
        if (total >= 91) return 'अ १';
        if (total >= 81) return 'अ २';
        if (total >= 71) return 'ब १';
        if (total >= 61) return 'ब २';
        if (total >= 51) return 'क १';
        if (total >= 41) return 'क २';
        if (total >= 31) return 'ड';
        if (total >= 21) return 'इ १';
        return 'इ २';
    };

    const calculateGradeDescription = (grade) => {
        if (grade == 'अ १') return 'अप्रतिम'
        if (grade == 'अ २') return 'खुप चांगला'
        if (grade == 'ब १') return 'चांगला'
        if (grade == 'ब २') return 'बरा'
        if (grade == 'क १') return 'सर्वसाधारण'
        if (grade == 'क २') return 'ठीक'
        if (grade == 'ड') return 'असमानधारक'
        if (grade == 'इ १') return 'सुधारणा आवश्यक'
        return 'सुधारणा आवश्यक';
    }

    const calculatePassOrFail = (grade) => {
        if (grade == 'इ २') return 'नापास'
        return 'पास'
    }

    const calculateStats = (student) => {
        const subjectStats = subjects.reduce((acc, subject) => {
            const theory = Math.min(Number(student[`${subject}Theory`]) || 0, 100);
            const practical = Math.min(Number(student[`${subject}Practical`]) || 0, 100);
            const total = theory + practical;
            const grade = calculateGrade(total);
            acc[subject] = { theory, practical, total, grade };
            return acc;
        }, {});

        const overallTotal = Object.values(subjectStats).reduce((sum, { total }) => sum + total, 0);
        const overallAverage = (overallTotal / (subjects.length * 100)) * 100; // Corrected calculation
        const overallGrade = calculateGrade(overallAverage);
        const overallGradeDescription = calculateGradeDescription(overallGrade);
        const PassOrFail = calculatePassOrFail(overallGrade);

        return {
            ...student,
            ...subjectStats,
            overallTotal: overallTotal.toFixed(2),
            overallAverage: overallAverage.toFixed(2),
            overallGrade,
            overallGradeDescription,
            PassOrFail,
        };
    };

    const data = useMemo(() => students.map(calculateStats), [students, subjects]);

    const columns = useMemo(
        () => [
            { Header: 'विद्यार्थ्यांचे नाव', accessor: 'name' },
            ...subjects.map(subject => ({
                Header: subject,
                columns: [
                    {
                        Header: 'आकारीक मुल्य',
                        accessor: `${subject}Theory`,
                    },
                    {
                        Header: 'संकलित मुल्य',
                        accessor: `${subject}Practical`,
                    },
                    {
                        Header: 'एकूण',
                        accessor: `${subject}.total`,
                        Cell: ({ value }) => value?.toFixed(2) || '',
                    },
                    {
                        Header: 'श्रेणी',
                        accessor: `${subject}.grade`,
                    },
                ],
            })),
            { Header: 'एकूण गुण', accessor: 'overallTotal' },
            { Header: 'शेकडा प्रमाण (%)', accessor: 'overallAverage' },
            { Header: 'श्रेणी', accessor: 'overallGrade' },
            { Header: 'श्रेणीवर्णन', accessor: 'overallGradeDescription' },
            { Header: 'शेरा', accessor: 'PassOrFail' },
            {
                Header: 'Actions',
                Cell: ({ row }) => (
                    <button onClick={() => deleteRow(row.original.id)} className="delete-row">
                        Delete
                    </button>
                ),
            },
        ],
        [subjects]
    );

    const {
        getTableProps,
        getTableBodyProps,
        headerGroups,
        rows,
        prepareRow,
    } = useTable({ columns, data });

    const handleInputChange = (e) => {
        setNewStudent({ ...newStudent, [e.target.name]: e.target.value });
    };

    const addStudent = () => {
        if (newStudent.name.trim()) {
            const studentData = {
                id: Date.now(),
                name: newStudent.name.trim(),
                ...subjects.reduce((acc, subject) => ({
                    ...acc,
                    [`${subject}Theory`]: '',
                    [`${subject}Practical`]: '',
                }), {}),
            };
            setStudents(prevStudents => [...prevStudents, studentData]);
            setNewStudent({ name: '' });
        }
    };

    const handleNewSubjectsChange = (e) => {
        setNewSubjects(e.target.value);
    };

    const addSubjects = () => {
        const subjectsToAdd = newSubjects.split(',').map(s => s.trim()).filter(s => s && !subjects.includes(s));
        setSubjects(prevSubjects => [...prevSubjects, ...subjectsToAdd]);
        setStudents(prevStudents => prevStudents.map(student => ({
            ...student,
            ...subjectsToAdd.reduce((acc, subject) => ({
                ...acc,
                [`${subject}Theory`]: '',
                [`${subject}Practical`]: '',
            }), {}),
        })));
        setNewSubjects('');
    };

    const removeSubjects = () => {
        const subjectsToRemove = newSubjects.split(',').map(s => s.trim());
        setSubjects(prevSubjects => prevSubjects.filter(subject => !subjectsToRemove.includes(subject)));
        setStudents(prevStudents => prevStudents.map(student => {
            const updatedStudent = { ...student };
            subjectsToRemove.forEach(subject => {
                delete updatedStudent[`${subject}Theory`];
                delete updatedStudent[`${subject}Practical`];
            });
            return updatedStudent;
        }));
        setNewSubjects('');
    };

    const handleCellEdit = (studentId, field, value) => {
        setStudents(prevStudents => prevStudents.map(student =>
            student.id === studentId ? { ...student, [field]: value } : student
        ));
    };

    const deleteRow = (id) => {
        setStudents(prevStudents => prevStudents.filter(student => student.id !== id));
    };

    const exportToExcel = async () => {
        const ExcelJS = await import('exceljs');
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Results');

        const schoolName = document.getElementById('school-name').innerText;
        const schoolClass = document.getElementById('school-class').innerText;
        const schoolSection = document.getElementById('school-section').innerText;
        const schoolPaperHeader = document.getElementById('school-paper-header').innerText;
        const schoolPaperYear = document.getElementById('school-year').innerText;

        // Add header row with merging
        worksheet.addRow([]);
        worksheet.getCell('A3').value = schoolName;
        worksheet.getCell('C3').value = schoolClass;
        worksheet.getCell('E3').value = schoolSection;

        worksheet.getCell('A3').font = { bold: true, size: 10 };
        worksheet.mergeCells('A3:B3');
        worksheet.getCell('A3').alignment = { horizontal: 'left', vertical: 'middle' };

        worksheet.getCell('C3').font = { bold: true, size: 10 };
        worksheet.mergeCells('C3:D3');
        worksheet.getCell('C3').alignment = { horizontal: 'left', vertical: 'middle' };

        worksheet.getCell('E3').font = { bold: true, size: 10 };
        worksheet.mergeCells('E3:F3');
        worksheet.getCell('E3').alignment = { horizontal: 'left', vertical: 'middle' };

        worksheet.addRow([]);

        const headerRow = ['विद्यार्थ्यांचे नाव'];
        subjects.forEach((subject) => {
            headerRow.push(subject, '', '', '');
        });

        headerRow.push('एकूण गुण', 'शेकडा प्रमाण (%)', 'श्रेणी', 'श्रेणीवर्णन', 'शेरा');
        worksheet.addRow(headerRow);

        let currentColumn = 3;
        subjects.forEach(() => {
            worksheet.mergeCells(worksheet.lastRow.number, currentColumn - 1, worksheet.lastRow.number, currentColumn + 2);
            worksheet.getCell(worksheet.lastRow.number, currentColumn - 1).alignment = { horizontal: 'center', vertical: 'middle' };
            currentColumn += 4;
        });

        // Adding sub-header row
        const subHeaderRow = [''];
        subjects.forEach(() => {
            subHeaderRow.push('आकारीक मुल्य', 'संकलित मुल्य', 'एकूण', 'श्रेणी');
        });
        subHeaderRow.push(' ', ' ', ' ');
        worksheet.addRow(subHeaderRow);

        // Apply "Rotate Text Up" style to specific cells in the sub-header row
        const subHeaderRowIndex = worksheet.lastRow.number; // Get the row index of the sub-header
        let currentColumnsubHeaderRow = 2; // Start after the first column (A)
        subjects.forEach(() => {
            ['आकारीक मुल्य', 'संकलित मुल्य', 'एकूण', 'श्रेणी'].forEach((_, index) => {
                const cell = worksheet.getCell(subHeaderRowIndex, currentColumnsubHeaderRow + index);
                cell.alignment = { textRotation: 90, horizontal: 'center', vertical: 'middle' }; // Rotate Text Up
            });
            currentColumnsubHeaderRow += 4; // Move to the next subject set
        });


        let subHeadercurrentColumn = 2;
        subjects.forEach(() => {
            worksheet.getColumn(subHeadercurrentColumn).width = 10;
            worksheet.getColumn(subHeadercurrentColumn + 1).width = 10;
            worksheet.getColumn(subHeadercurrentColumn + 2).width = 10;
            worksheet.getColumn(subHeadercurrentColumn + 3).width = 10;
            currentColumn += 4;
        });

        worksheet.getColumn(subHeadercurrentColumn).width = 10;
        worksheet.getColumn(subHeadercurrentColumn + 1).width = 10;
        worksheet.getColumn(subHeadercurrentColumn + 2).width = 10;
        worksheet.getColumn(subHeadercurrentColumn + 3).width = 10;

        worksheet.mergeCells('A5:A6');
        worksheet.getCell('A5').alignment = { horizontal: 'center', vertical: 'middle' };

        data.forEach((student) => {
            const row = [student.name];
            subjects.forEach((subject) => {
                const theory = student[`${subject}Theory`] || '';
                const practical = student[`${subject}Practical`] || '';
                const total = Number(theory) + Number(practical) || 'NaN';
                const grade = calculateGrade(total) || 'NaN';
                row.push(theory, practical, total, grade);
            });
            row.push(
                student.overallTotal || '',
                (student.overallTotal / subjects.length).toFixed(2) || '',
                student.overallGrade || '',
                student.overallGradeDescription || '',
                student.PassOrFail || ''
            );
            worksheet.addRow(row);
        });

        worksheet.columns.forEach((column) => {
            let maxWidth = 10;
            column.eachCell({ includeEmpty: true }, (cell) => {
                const cellText = cell.value ? cell.value.toString() : '';
                maxWidth = Math.max(maxWidth, cellText.length);
            });
            column.width = maxWidth + 2;
        });

        const lastRow = worksheet.lastRow.number;
        let lastColumn = 1;
        worksheet.eachRow((row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (cell.col > lastColumn) {
                    lastColumn = cell.col;
                }
            });
        });

        for (let row = 5; row <= lastRow; row++) {
            for (let col = 1; col <= lastColumn; col++) {
                const cell = worksheet.getCell(row, col);
                cell.border = {
                    top: { style: 'thin' },
                    right: { style: 'thick' },
                    bottom: { style: 'thin' },
                    left: { style: 'thick' },
                };
            }
        }

        const columnToLetter = (colIndex) => {
            let letter = '';
            while (colIndex > 0) {
                const remainder = (colIndex - 1) % 26;
                letter = String.fromCharCode(65 + remainder) + letter;
                colIndex = Math.floor((colIndex - 1) / 26);
            }
            return letter;
        };

        const lastColumnLetter = columnToLetter(lastColumn);
        const headerRange = `A1:${lastColumnLetter}2`;
        worksheet.mergeCells(headerRange);
        worksheet.getCell('A1').value = `${schoolPaperHeader} संकलित गुणपत्रक तक्ता ${schoolPaperYear}`;
        worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getCell('A1').font = { bold: true, size: 12 };
        worksheet.getCell('A5').border = {
            top: { style: 'thick' },
            right: { style: 'thick' },
            left: { style: 'thick' },
            bottom: { style: 'thick' }
        }


        // Apply thick borders to specific cells in the sub-header row
        const subHeaderRowIndexStyle = subHeaderRowIndex; // Get the row index of the sub-header
        let currentColumnsubHeader = 2; // Start after the first column (A)

        subjects.forEach(() => {
            const firstCell = worksheet.getCell(subHeaderRowIndexStyle, currentColumnsubHeader); // First cell
            const secondCell = worksheet.getCell(subHeaderRowIndexStyle, currentColumnsubHeader + 1); // Second cell
            const thirdCell = worksheet.getCell(subHeaderRowIndexStyle, currentColumnsubHeader + 2); // Third cell
            const fourthCell = worksheet.getCell(subHeaderRowIndexStyle, currentColumnsubHeader + 3); // Fourth cell

            // First cell (left and bottom thick borders)
            firstCell.border = {
                left: { style: 'thick' },
                bottom: { style: 'thick' },

                right: { style: 'thin' },
                top: { style: 'thick' },
            };

            // Second cell (bottom thick border)
            secondCell.border = {
                bottom: { style: 'thick' },

                left: { style: 'thin' },
                right: { style: 'thin' },
                top: { style: 'thick' },
            };

            // Third cell (bottom thick border)
            thirdCell.border = {
                bottom: { style: 'thick' },

                left: { style: 'thin' },
                right: { style: 'thin' },
                top: { style: 'thick' },
            };

            // Fourth cell (right and bottom thick borders)
            fourthCell.border = {
                right: { style: 'thick' },
                bottom: { style: 'thick' },

                left: { style: 'thin' },
                top: { style: 'thick' },
            };

            // Set column width to 10 px
            worksheet.getColumn(currentColumn).width = 10;
            worksheet.getColumn(currentColumn + 1).width = 10;
            worksheet.getColumn(currentColumn + 2).width = 10;
            worksheet.getColumn(currentColumn + 3).width = 10;

            currentColumn += 4; // Move to the next set of columns
            currentColumnsubHeader += 4; // Move to the next set of columns
        });

        const lastRowIndex = worksheet.lastRow.number; // Get the index of the last row
        worksheet.getRow(lastRowIndex).eachCell({ includeEmpty: true }, (cell) => {
            cell.border = {
                ...cell.border, // Retain any existing borders
                bottom: { style: 'thick' }, // Add a thick bottom border
                top: { style: 'thin' },
                left: { style: 'thick' },
                right: { style: 'thick' }
            };
        });

        worksheet.getRow(subHeaderRowIndex).eachCell({ includeEmpty: true }, (cell, colNumber) => {
            lastColumnIndex = Math.max(lastColumnIndex, colNumber); // Track the last column with data
        });
        // The row where you want to apply the thick border (A5)
        let lastColumnIndex = 1; // Start from the first column

        // Apply thick borders to the entire row from A5 to the last column with data
        for (let col = 1; col <= lastColumnIndex; col++) {
            const cell = worksheet.getCell(subHeaderRowIndex, col);
            cell.border = {
                top: { style: 'thick' },
                bottom: { style: 'thick' },
                left: { style: col === 1 ? 'thick' : undefined }, // Thick left border for the first cell
                right: { style: col === lastColumnIndex ? 'thick' : undefined }, // Thick right border for the last cell
            };
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'school_results.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };



    const exportToImage = async () => {
        const elements = ['resultInfo', 'school-paper-header', 'school-semester', 'resultsTable',];

        const tempContainer = document.createElement('div');
        tempContainer.style.position = 'absolute';
        tempContainer.style.top = '-9999px';

        elements.forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                const clonedElement = element.cloneNode(true);
                tempContainer.appendChild(clonedElement);
            }
        });

        document.body.appendChild(tempContainer);

        try {
            const canvas = await html2canvas(tempContainer);

            const link = document.createElement('a');
            link.download = 'school_results.png';
            link.href = canvas.toDataURL();
            link.click();
        } catch (err) {
            console.error('Error capturing elements:', err);
        } finally {
            document.body.removeChild(tempContainer);
        }
    };

    return (
        <main>
            <section className='header-hero'>
                <img src={schoollogo} alt="" />
                <div>
                    <h1>कै.आ.ह. आब्बा प्राथमिक विद्यालय सोलापूर</h1>
                    <h3>94/41, जोडभावी पेठ सोलापूर</h3>
                </div>
            </section>

            <div className="school-result-checker" >
                <div className='school-informations' id="resultInfo">
                    <div className='school-informations-section1'>
                        <h3 id='school-name'>शाळेचे नांव : <span className="editable-header" contenteditable="true" spellcheck="false">कै.आ.ह. आब्बा प्राथमिक विद्यालय सोलापूर</span></h3>
                        <h3 id='teacher-name'>वर्ग शिक्षकाचे नांव : <span className="editable-header" contenteditable="true" spellcheck="false">शिक्षकांचे नाव</span></h3>
                    </div>
                    <div className='school-informations-section2'>
                        <h3 id='school-year'>सन : <span className="editable-header" contenteditable="true" spellcheck="false">२०२३-२४</span></h3>
                        <h3 id='school-class'>वर्ग : <span className="editable-header" contenteditable="true" spellcheck="false">१ ली</span></h3>
                        <h3 id='school-section'>तुकडी : <span className="editable-header" contenteditable="true" spellcheck="false">अ</span></h3>
                    </div>
                </div>
                <h1 id='school-paper-header' className="editable-header" contenteditable="true" spellcheck="false">सातत्यपूर्ण सर्वंकष मूल्यमापन</h1>
                <h3 id='school-semester' className="editable-header semester-header" contenteditable="true" spellcheck="false">प्रथम सत्र / द्वितीय सत्र</h3>
                <div className="input-section">
                    <input
                        type="text"
                        name="name"
                        placeholder="विद्यार्थ्यांचे नाव"
                        value={newStudent.name}
                        onChange={handleInputChange}
                        spellcheck="false"
                    />
                    <button onClick={addStudent}>विद्यार्थी ॲड करा</button>
                    <input
                        type="text"
                        placeholder="विषय ॲड / रिमूव करा (comma-separated)"
                        value={newSubjects}
                        onChange={handleNewSubjectsChange}
                    />
                    <button onClick={addSubjects}>विषय ॲड करा</button>
                    <button onClick={removeSubjects}>विषय रिमूव करा</button>
                </div>
                <div className="table-container" id="resultsTable">
                    <table {...getTableProps()}>
                        <thead>
                            {headerGroups.map((headerGroup, i) => (
                                <tr {...headerGroup.getHeaderGroupProps()} key={i}>
                                    {headerGroup.headers.map((column, j) => (
                                        <th {...column.getHeaderProps()} key={j} colSpan={column.columns ? column.columns.length : 1}>
                                            {column.render('Header')}
                                        </th>
                                    ))}
                                </tr>
                            ))}
                        </thead>
                        <tbody {...getTableBodyProps()}>
                            {rows.map(row => {
                                prepareRow(row);
                                return (
                                    <tr {...row.getRowProps()}>
                                        {row.cells.map(cell => (
                                            <td {...cell.getCellProps()}>
                                                {cell.column.id === 'name' || cell.column.id.includes('Theory') || cell.column.id.includes('Practical') ? (
                                                    <input className={`datainput ${cell.column.id == 'name' ? "nameinput" : ""}`}
                                                        value={cell.value || ''}
                                                        onChange={(e) => handleCellEdit(row.original.id, cell.column.id, e.target.value)}
                                                        placeholder='Enter Marks here'
                                                        onBlur={(e) => {
                                                            if (cell.column.id.includes('Theory')) {
                                                                handleCellEdit(row.original.id, cell.column.id, Math.min(Number(e.target.value), 100).toString());
                                                            } else if (cell.column.id.includes('Practical')) {
                                                                handleCellEdit(row.original.id, cell.column.id, Math.min(Number(e.target.value), 100).toString());
                                                            }
                                                        }}
                                                    />
                                                ) : (
                                                    cell.render('Cell')
                                                )}
                                            </td>
                                        ))}
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            </div>
            <div className="export-section">
                <button onClick={exportToExcel}>Export to Excel</button>
                <button onClick={exportToImage}>Export to Image</button>
            </div>
            <footer>
                <h3>Developed by  <a href="https://vishwalatarati.in/">Vishwalatarati Digital Solutions Private Limited</a></h3>
            </footer>
        </main>
    );
}