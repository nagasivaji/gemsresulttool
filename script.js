// Horizontal drag 
const slider = document.getElementById('htmlTableArea');
let isDown = false;
let startX;
let scrollLeft;

slider.addEventListener('mousedown', (e) => {
    isDown = true;
    slider.classList.add('active');
    startX = e.pageX - slider.offsetLeft;
    scrollLeft = slider.scrollLeft;
});
slider.addEventListener('mouseleave', () => {
    isDown = false;
    slider.classList.remove('active');
});
slider.addEventListener('mouseup', () => {
    isDown = false;
    slider.classList.remove('active');
});
slider.addEventListener('mousemove', (e) => {
    if(!isDown) return;
    e.preventDefault();
    const x = e.pageX - slider.offsetLeft;
    const walk = (x - startX) * 3; //scroll-fast
    slider.scrollLeft = scrollLeft - walk;
    //console.log(walk);
    });

//Accessing Process Button
const processFileBtn = document.getElementById("processFileBtn");
// Accessing file input
const excelFile = document.getElementById('excelFile');


// Adding event Listener to handle click evenr
processFileBtn.addEventListener("click", () => {
    processExcelFile();
});




// VARABLES
// Variables for excel file 
var excelBook;
var excelSheets;
var excelSheetNo = 0;

// array to strore headings of the selected sheet
var headings = [];

// array to strore pass marks for different subjects
var passMarks = [];

// array to strore students data
var students = [];

// Assessments index array
var indexes = [];

// Map for calculating the toppers
var marksMap = new Map();


// Getting data in required sheet from excel file
function getSheetData() {
    // console.log(excelSheetNo);
    // console.log(excelSheets[excelSheetNo]);
    var sheetData = XLSX.utils.sheet_to_json(excelBook.Sheets[excelSheets[excelSheetNo]], { header: 1 });
    return sheetData;
}



function processExcelFile() {
    //console.log("Processing");

    // Hiding File input area
    document.getElementById("fileInputArea").style.display = "none";

    //showing backbtn on navigationbar
    document.getElementById("navBarBackBtn").style.display = "block";

    try {
        if (checkExcelOrNot()) {
            // If input is excel file
            //console.log("Correct File...");

            readExcelFile();
        } else {
            // If file is not an excel file
            //console.log("Wrong File...");
            const errorArea = document.getElementById("errorArea");
            errorArea.style.display = 'block';
            errorArea.innerHTML = "<p>Only Excel files are allowed</p>";
            setTimeout(() => {
                location.reload();
            }, 2000);
        }
    } catch (e) {
        // If file input in empty
        const errorArea = document.getElementById("errorArea");
        errorArea.style.display = 'block';
        errorArea.innerHTML = "<p>No file Choosen</p>";
        setTimeout(() => {
            location.reload();
        }, 2000);
    }

}



// file type checker
function checkExcelOrNot() {

    if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ].includes(excelFile.files[0].type)) {

        //excelFile.value = '';

        return false;
    }
    return true;
}


// reading excel file
function readExcelFile() {

    // Creating reader for excel file
    var reader = new FileReader();

    // Reading First Excel file from the input
    reader.readAsArrayBuffer(excelFile.files[0]);

    // ````````Anonymous function````````
    // after reading full file invoking an anonymous function for fetching data from excel file
    reader.onload = (event) => {

        // Storing reader reasult into a variable in array format
        var excelResult = new Uint8Array(reader.result);

        // Excel result book similar to excel file
        excelBook = XLSX.read(excelResult, { type: 'array' });

        // Getting excel sheets array from excel reasult book
        excelSheets = excelBook.SheetNames;

        // Show sheets dropdown
        showSheetsDropdown(excelSheets);

    }
}


// Display dropdown for selection os sheets
function showSheetsDropdown(sheets) {

    // Hiding File input area
    document.getElementById("fileInputArea").style.display = "none";

    // Creating dropdown for selection
    const sheetInputArea = document.getElementById("sheetInputArea");
    var temp = `<p>Select Sheet:</p><select id="selectSheet" class="form-select" aria-label="Default select example">`;
    for (var i = 0; i < sheets.length; i++) {
        temp += `<option value="${i}">${i + 1}:${sheets[i]}</option>`;
    }
    temp += `</select>`;
    // assigning dropdown to select dropdown area
    sheetInputArea.innerHTML = temp;

    // EVENT LISTENER FOR SELECT SHEET DROPDOWN
    // Accessing dropdown selection
    const selectSheet = document.getElementById("selectSheet");


    //getting required indexes for assessments columns
    searchIndex();
    //calling show Subjects
    showSubjects();

    // Adding event Listener to select sheet dropdown
    selectSheet.addEventListener("change", () => {
        excelSheetNo = selectSheet.value;
        // getting required indexes for assessments columns
        searchIndex();
        // Showing subjects list
        showSubjects();
    });
}



// Display inputs to senter passmarks for each subject
function showSubjects() {
    // getting data from required sheet in excel file
    var sheetData = getSheetData();


    // getting headings  and creating input fields for different subjects
    // creating new headings for different subjects
    headings = [];
    var temp = ``;

    for (var i = 0; i < sheetData[1].length; i++) {
        if (i < 3)
            headings.push(sheetData[0][i]);
        else {
            if (i >= indexes[0] && i < indexes[1]) {
                headings.push(sheetData[1][i]);
                temp += `
                        <div class="input-group mb-3">
                            <span class="input-group-text" id="inputGroup-sizing-default">MARKS</span>
                            <input type="number" class="form-control inputMarksForSubject" aria-label="Sizing example input" aria-describedby="inputGroup-sizing-default" min="0" max="100" placeholder="${sheetData[1][i]}">
                        </div>` ;
                i++;
            }

        }
    }

    temp += `<button class="btn btn-dark" id="claculateBtn">Calculate</button>`;
    // Appending inputs to subject area
    document.getElementById('subjectsArea').innerHTML = temp;

    //console.log(headings);

    //Calculate
    const claculateBtn = document.getElementById('claculateBtn');
    claculateBtn.addEventListener("click", () => {
        getPassMarks();
        getStudentsObjects();
        calculateRanks();
        prepareHTMLTable(students);
        showSearchArea();
    });
}



// geting indexes for assessments columns
function searchIndex() {
    const data = getSheetData();
    indexes = [];
    var flag = 0;
    var startIndex;
    var endIndex;

    //console.log(tempColumns1);
    //console.log(tempColumns2);

    for (var i = 0; i < data[0].length; i++) {
        if (data[0][i] !== undefined) {
            if (data[0][i].toLowerCase() === "assessments") {
                startIndex = i;
                flag = 1;
            } else {
                if (flag === 1) {
                    endIndex = i;
                    break;
                }
            }
        }
    }

    //console.log(startIndex);
    //console.log(endIndex);

    indexes.push(startIndex);
    indexes.push(endIndex);

}


// Getting pass marks from subjects input area
function getPassMarks() {
    // creating new array for passing marks
    passMarks = [];
    var temp = document.getElementsByClassName("inputMarksForSubject");
    for (var i = 0; i < temp.length; i++) {
        //console.log(Number(temp[i].value));
        passMarks.push(Number(temp[i].value));
    }
    //console.log(passMarks);

}


// Students class
class Student {
    studentId;
    studentName;
    studentEmail;
    studentMarks;
    objectiveStatus;
    subjectiveStatus;
    finalStatus;
    finalMarks;
    grade;
    partialStatus;
    rank;

    constructor(studentId, studentName, studentEmail, studentMarks, objectiveStatus, subjectiveStatus, finalStatus, finalMarks, grade, partialStatus, rank) {
        this.studentId = studentId;
        this.studentName = studentName;
        this.studentEmail = studentEmail;
        this.studentMarks = studentMarks;
        this.objectiveStatus = objectiveStatus;
        this.subjectiveStatus = subjectiveStatus;
        this.finalStatus = finalStatus;
        this.finalMarks = finalMarks;
        this.grade = grade;
        this.partialStatus = partialStatus;
        this.rank = rank;
    }
}

// Subjects class
class Subject {
    subjectName;
    subjectMarks;
    subjectPassMark;
    subjectAttempts;
    subjectStatus;

    constructor(subjectName, subjectMarks, subjectPassMark, subjectAttempts, subjectStatus) {
        this.subjectName = subjectName;
        this.subjectMarks = subjectMarks;
        this.subjectPassMark = subjectPassMark;
        this.subjectAttempts = subjectAttempts;
        this.subjectStatus = subjectStatus;
    }
}



function getStudentsObjects() {

    // creating new students array
    students = [];

    //creating new map for ranks
    marksMap = new Map();

    // getting data from required sheet in excel file
    var sheetData = getSheetData();
    //console.log(sheetData);

    // getting noOfRows and noOfColumns in excel file sheet
    const noOfRows = sheetData.length;
    const noOfCols = sheetData[0].length;

    // Iterating from second row since first row contains headings
    // For each students...
    for (var i = 2; i < noOfRows; i++) {
        // Creating required variables which will be used while creating objects
        var studentId, studentName, studentEmail;
        var subjectName, subjectMarks, subjectAttempts, subjectStatus;
        var studentMarks = [];

        // temp variable is indexing variable for headings array
        var temp1 = 3;
        // temp variable is indexing variable for headings array
        var temp2 = 0;

        for (var j = 0; j < noOfCols; j++) {
            if (j == 0)        // col:1 - studens id 
                studentId = sheetData[i][j];
            else if (j == 1)   // col:2  - student name
                studentName = sheetData[i][j];
            else if (j == 2)   // col:3  - student email
                studentEmail = sheetData[i][j];
            else            // col:3 to ...  - subjects
            {
                if (j < indexes[0] || j >= indexes[1])
                    continue;


                subjectName = headings[temp1++];     // subject name
                subjectMarks = sheetData[i][j];     // subject marks
                j++;
                subjectAttempts = sheetData[i][j];  // subject attemps

                // function to validate students subject status
                subjectStatus = validate(subjectMarks, subjectAttempts, passMarks[temp2]);

                // creating new subject object for the student
                var obj = new Subject(subjectName, subjectMarks, passMarks[temp2++], subjectAttempts, subjectStatus);

                // pushing the new subject into student marks
                studentMarks.push(obj);
            }
        }

        //console.log(studentMarks);

        // required varibale for students status
        var objectiveStatus = "NOT APPEARED", subjectiveStatus = "NOT APPEARED";

        // Iterating through  students sbjets array
        for (var k = 0; k < studentMarks.length; k++) {

            // getting subject object
            var subject = studentMarks[k];
            //console.log(subject);
            //console.log(subject.subjectName.toLowerCase("subjective"));

            // if its a  subjective type
            if (subject.subjectName.toLowerCase().includes("subjective")) {
                if (subject.subjectStatus == 'FAIL') {
                    subjectiveStatus = "FAIL";
                    break;
                }

                if (subject.subjectStatus == 'PENDING') {
                    subjectiveStatus = "PENDING";
                }

                if (subject.subjectStatus == 'PASS') {
                    subjectiveStatus = "PASS";
                }

            }
            else    // if its a  objective type
            {
                //console.log(subject);
                if (subject.subjectStatus == 'FAIL') {
                    objectiveStatus = "FAIL";
                    break;
                }

                if (subject.subjectStatus == 'PENDING') {
                    objectiveStatus = "PENDING";
                }

                if (subject.subjectStatus == 'PASS') {
                    objectiveStatus = "PASS";
                }

            }
        }


        // Conditions for finalstatus and sinal marks
        // required varibale for students status 
        var finalStatus, finalMarks;
        // final grade for student
        var grade = "NA"

        // Condtion for student to FAIL
        if (objectiveStatus == 'FAIL' || subjectiveStatus == 'FAIL') {
            finalStatus = 'FAILED';
            finalMarks = '--';
        }
        // Condtion for student to NOT APPEARED
        else if (objectiveStatus == 'NOT APPEARED' || subjectiveStatus == 'NOT APPEARED') {
            finalStatus = 'NOT APPEARED';
            finalMarks = '--';
        }
        // Condtion for student to PENDING
        else if (objectiveStatus == 'PENDING' || subjectiveStatus == 'PENDING') {
            finalStatus = 'PENDING';
            finalMarks = '--';
        }
        // Conditions for student to PASS
        else {

            finalStatus = 'PASS';
            var objMarks = 0, subMarks = 0, subjectCount = 0;

            // Iterating through  students sbjets array
            for (var k = 0; k < studentMarks.length; k++) {

                // getting subject object
                var subject = studentMarks[k];

                // Neglecting the subject if its has passmark zero
                if(subject.subjectPassMark == 0)
                    continue;

                // then its subjective
                if (subject.subjectName.toLowerCase().includes("subjective")) {
                    subMarks = subject.subjectMarks;
                }
                else // then its objective
                {
                    objMarks += subject.subjectMarks;
                    subjectCount++;
                }
            }

            //console.log(objMarks, subMarks);

            finalMarks = ((objMarks / subjectCount) + (subMarks)) / 2;
            finalMarks = Math.round(finalMarks);

            // Calculating Grade
            grade = getGrade(finalMarks);
        }



        //If student fails.. calculating partialStatus condition 
        var partialStatus = "NA";
        if(finalStatus === "FAILED"){
            partialStatus = getPartialStatus(studentMarks);
        }


        // creating new studemt object
        var obj = new Student(studentId, studentName, studentEmail, studentMarks, objectiveStatus, subjectiveStatus, finalStatus, finalMarks, grade, partialStatus, "--");

        // pushing the new student object to the list of students
        students.push(obj);

        // Updating map
        if(finalStatus === "PASS")
            marksMap.set(studentId, finalMarks);
        else
        marksMap.set(studentId, 0);
    }
    //console.log(students);
}


// Marks validation table
function validate(marks, attempts, criteria) {
    if (marks >= criteria && attempts <= 3) {
        return "PASS";
    }
    else if (marks < criteria && attempts < 3) {
        return "PENDING";
    }
    else if (marks < criteria && attempts == 3)
        return "FAIL";
    else
        return "NA";
}

//Grade validation function 
function getGrade(marks){
    if(marks >= 90)
        return "A+"
    else if(marks >= 80)
        return  "A";
    else if(marks >= 70)
        return "B";
    else 
        return "F"
}

// getting partialStatus
function getPartialStatus(studentMarks){
    // Iterating through  students sbjets array
    for (var k = 0; k < studentMarks.length; k++) {

        // getting subject object
        var subject = studentMarks[k];

        // Neglecting the subject if its has passmark zero
        if(subject.subjectPassMark === 0)
            continue;

        if(subject.subjectMarks < 60 )
            return "NO";

        // log
        //console.log(subject.subjectMarks);
    }
    //console.log("-----");
    return "YES";
}

// calculating student ranks
function calculateRanks(){
    var  map= new Map([...marksMap.entries()].sort((a,b) => b[1] - a[1]));
    //console.log(map);
    //console.log(students);
    var tempRank = 1;
    for (const [key, value] of map) {
        //console.log(key, value);
        for(var i=0;i<students.length;i++) {
            if(students[i].studentId === key && students[i].finalStatus === "PASS")
                students[i].rank = "Topper "+ tempRank++;
        }
    }
    //console.log(students);
}

// preparing html tbale
function prepareHTMLTable() {

    // Hiding sheetarea -  dropdown and  subjects input area - passmarks inputs
    document.getElementById("subjectsArea").style.display = "none";
    document.getElementById("sheetInputArea").style.display = "none";


    //creating HTML table based on the result of students data
    //Table start
    var tableStart = `<table class="table" id="htmlResultTable">`;


    //console.log(headings);
    //table Headings -> id, name, mail, subjects...
    // Openinig row
    var tableHeadings = `<tr>`;
    for (var i = 0; i < headings.length; i++) {
        if (i < 3)
            tableHeadings = tableHeadings + `<th>${headings[i]}</th>`;
        else{
            // tableHeadings = tableHeadings + `<th>${headings[i]} <br> ${passMarks[i - 3]}</th>`;
            tableHeadings = tableHeadings + `<th>${headings[i]}</th>`;
            tableHeadings = tableHeadings + `<th>Attemps</th>`;
        }
    }

    // Objective status
    tableHeadings = tableHeadings + `<th>Objective Status</th>`;
    // Subjective status
    tableHeadings = tableHeadings + `<th>Subjective Status</th>`;
    // Final status
    tableHeadings = tableHeadings + `<th>Final Status</th>`;
    // Finak Marks
    tableHeadings = tableHeadings + `<th>Final Marks</th>`;
    // Grade
    tableHeadings = tableHeadings + `<th>Grade</th>`;
    // Rank
    tableHeadings = tableHeadings + `<th>Rank</th>`;
    // Closing row
    tableHeadings = tableHeadings + `</tr>`;



    //console.log(students);
    // table rows 
    // All students data
    var tableRows = ``;
    for (var i = 0; i < students.length; i++) {
        // creating row tag based on status of the student
        var tempRow = ``;
        if (students[i].finalStatus === 'PASS')
            tempRow = tempRow + `<tr class="passRow">`;
        else if (students[i].finalStatus === 'FAILED' && students[i].partialStatus === 'YES')
            tempRow = tempRow + `<tr class="partialRow">`;
        else if (students[i].finalStatus === 'FAILED')
            tempRow = tempRow + `<tr class="failRow">`;
        else if (students[i].finalStatus === 'PENDING')
            tempRow = tempRow + `<tr class="pendingRow">`;
        else
            tempRow = tempRow + `<tr class="notAppearedRow">`;
        
    


        // creating col tags for student data
        // 1.Student ID
        tempRow = tempRow + `<td>${students[i].studentId}</td>`;
        // 2.Student Name
        tempRow = tempRow + `<td>${students[i].studentName}</td>`;
        // 3.Student Email
        tempRow = tempRow + `<td>${students[i].studentEmail}</td>`;

        // 4.Student marks and attempts for each Subject (Adding dynamically through looping subjects array from student object)
        for (var j = 0; j < students[i].studentMarks.length; j++) {
            tempRow = tempRow + `<td>
                                    ${students[i].studentMarks[j].subjectMarks}
                                </td>`;

            tempRow = tempRow + `<td>
                                    ${students[i].studentMarks[j].subjectAttempts}
                                </td>`;
        }

        // 5.Student Objective Status
        tempRow = tempRow + `<td>${students[i].objectiveStatus}</td>`;

        // 6.Student Subjective Status
        tempRow = tempRow + `<td>${students[i].subjectiveStatus}</td>`;

        // 7.Student Final Status
        tempRow = tempRow + `<td>${students[i].finalStatus}</td>`;

        // 8.Student Final Marks
        tempRow = tempRow + `<td>${students[i].finalMarks}</td>`;

        // 9.Student Final Marks
        tempRow = tempRow + `<td>${students[i].grade}</td>`;

        // 10.Student Final Marks
        tempRow = tempRow + `<td>${students[i].rank}</td>`;

        // closing row tag
        tempRow = tempRow + '</tr>';


        // Adding each temporary student row to table rows string 
        tableRows = tableRows + tempRow;
    }


    // table end
    var tableEnd = `</table>`;

    // download Button 
    var downloadButton = '<button class="btn btn-dark" id="downloadExcel">Get Excel</button>';

    // Adding table to HTML page
    document.getElementById("htmlTableArea").innerHTML = tableStart + tableHeadings + tableRows + tableEnd + downloadButton;
    document.getElementById("htmlTableArea").style.overflow = 'scroll';


    // Downloading excelfile using event listner
    document.getElementById("downloadExcel").addEventListener("click", function () {
        downloadExcelFile();
    });

}



// Show search Area
function showSearchArea() {
    // showing search input area
    document.getElementById("searchInputhArea").style.display = 'flex';

    // variable for type of search
    var searchField = 1;

    // Accessing search selection
    const searchSelection = document.getElementById("searchSelection");

    // Adding event handlers to select search type
    searchSelection.addEventListener("change", () => {
        searchField = searchSelection.value;
        //console.log(searchField);
    });

    // Accessing search value
    const searchInput = document.getElementById("searchInput");

    // search data
    var searchData = '';

    // Adding event handler to search value
    searchInput.addEventListener("input", () => {
        searchData = searchInput.value;
        sortStudentData(searchField, searchData);
    });

    // Accessing search btn
    const searchBtn = document.getElementById("searchBtn");

    // adding event handler to search btn
    searchBtn.addEventListener("click", () => {
        sortStudentData(searchField, searchData.toLowerCase());
    });


}

// Sorting students data to show in search result
function sortStudentData(field, data) {
    if (data === '')
        prepareHTMLSearchTable([]);

    var tempStudents = [];
    for (var i = 0; i < students.length; i++) {
        if (field == 1) {
            var tempData = students[i].studentId;
            if (tempData == data)
                tempStudents.push(students[i]);
        }
        else if (field == 2) {
            var tempData = students[i].studentName.toLowerCase();
            if (tempData.includes(data))
                tempStudents.push(students[i]);
        }
        else {
            var tempData = students[i].studentEmail.toLowerCase();
            if (tempData.includes(data))
                tempStudents.push(students[i]);
        }
    }

    //console.log(tempStudents);
    prepareHTMLSearchTable(tempStudents);
}


// preparing html tbale for search result
function prepareHTMLSearchTable(searchStudents) {
    //console.log(searchStudents);
    //creating HTML table based on the result of students data
    //Table start
    var tableDivHeading = '<p>Search Result:</p>';
    var tableStart = `<table class="table" id="htmlSearchTable">`;



    //table Headings -> id, name, mail, subjects...
    // Openinig row
    var tableHeadings = `<tr>`;
    for (var i = 0; i < headings.length; i++) {
        if (i < 3)
            tableHeadings = tableHeadings + `<th>${headings[i]}</th>`;
        else
            tableHeadings = tableHeadings + `<th">${headings[i]} <br> ${passMarks[i - 3]}</th>`;
    }

    // Objective status
    tableHeadings = tableHeadings + `<th>Objective Status</th>`;
    // Subjective status
    tableHeadings = tableHeadings + `<th>Subjective Status</th>`;
    // Final status
    tableHeadings = tableHeadings + `<th>Final Status</th>`;
    // Finak Marks
    tableHeadings = tableHeadings + `<th>Final Marks</th>`;
    // Closing row
    tableHeadings = tableHeadings + `</tr>`;



    //console.log(students);
    // table rows 
    // All students data
    var tableRows = ``;
    for (var i = 0; i < searchStudents.length; i++) {
        // creating row tag based on status of the student
        var tempRow = ``;
        if (searchStudents[i].finalStatus === 'PASS')
            tempRow = tempRow + `<tr class="passRow">`;
        else if (searchStudents[i].finalStatus === 'FAILED')
            tempRow = tempRow + `<tr class="failRow">`;
        else if (searchStudents[i].finalStatus === 'PENDING')
            tempRow = tempRow + `<tr class="pendingRow">`;
        else
            tempRow = tempRow + `<tr class="notAppearedRow">`;


        // creating col tags for student data
        // 1.Student ID
        tempRow = tempRow + `<td>${searchStudents[i].studentId}</td>`;
        // 2.Student Name
        tempRow = tempRow + `<td>${searchStudents[i].studentName}</td>`;
        // 3.Student Email
        tempRow = tempRow + `<td>${searchStudents[i].studentEmail}</td>`;

        // 4.Student marks and attempts for each Subject (Adding dynamically through looping subjects array from student object)
        for (var j = 0; j < searchStudents[i].studentMarks.length; j++) {
            tempRow = tempRow + `<td>
                                    ${searchStudents[i].studentMarks[j].subjectMarks}
                                    -
                                    ${searchStudents[i].studentMarks[j].subjectAttempts}
                                </td>`;
        }

        // 5.Student Objective Status
        tempRow = tempRow + `<td>${searchStudents[i].objectiveStatus}</td>`;

        // 6.Student Subjective Status
        tempRow = tempRow + `<td>${searchStudents[i].subjectiveStatus}</td>`;

        // 7.Student Final Status
        tempRow = tempRow + `<td>${searchStudents[i].finalStatus}</td>`;

        // 7.Student Final Marks
        tempRow = tempRow + `<td>${searchStudents[i].finalMarks}</td>`;

        // closing row tag
        tempRow = tempRow + '</tr>';


        // Adding each temporary student row to table rows string 
        tableRows = tableRows + tempRow;
    }


    // table end
    var tableEnd = `</table>`;

    if (searchStudents.length <= 0 || searchStudents.length == students.length) {
        // Adding table to HTML page
        document.getElementById("searchResultArea").style.display = "none";
    }
    else {
        // Adding table to HTML page
        document.getElementById("searchResultArea").style.display = "block";
        document.getElementById("searchResultArea").innerHTML = tableDivHeading + tableStart + tableHeadings + tableRows + tableEnd;
        document.getElementById("searchResultArea").style.overflow = 'scroll';
    }

}


// Downloading excelfile
function downloadExcelFile() {

    function html_table_to_excel(type) {
        var data = document.getElementById('htmlResultTable');
        console.log(data);

        var file = XLSX.utils.table_to_book(data, { sheet: "sheet1" });
        console.log(file);
        XLSX.write(file, { bookType: type, bookSST: true, type: 'base64' });
        XLSX.writeFile(file, 'Exco Result.' + type);

        setTimeout(() => {
            location.reload();
        }, 1000);
    }

    html_table_to_excel('xlsx');
}

/*

        

*/