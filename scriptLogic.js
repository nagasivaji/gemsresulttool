class Student{
    studentId;
    studentName;
    studentEmail;
    studentMarks;
    objectiveStatus;
    subjectiveStatus;

    constructor(studentId, studentName, studentEmail, studentMarks, objectiveStatus, subjectiveStatus) {
        this.studentId = studentId;
        this.studentName = studentName; 
        this.studentEmail = studentEmail;
        this.studentMarks =  studentMarks;
        this.objectiveStatus =  objectiveStatus;
        this.subjectiveStatus =  subjectiveStatus;
    }
}

class Subject{
    subjectName;
    subjectMarks;
    subjectAttempts;
    subjectStatus;

    constructor(subjectName, subjectMarks, subjectAttempts, subjectStatus) {
        this.subjectName = subjectName;
        this.subjectMarks = subjectMarks;
        this.subjectAttempts = subjectAttempts;
        this.subjectStatus = subjectStatus;
    }
}

// Excel all heading
var headings = [];
// Employee objects array
var students = [];







const excelFile = document.getElementById('excelFile');

excelFile.addEventListener('change', (event) => {

    if(!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(event.target.files[0].type))
    {
        document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Only .xlsx or .xls file format are allowed</div>';

        excelFile.value = '';

        return false;
    }

    var reader = new FileReader();

    reader.readAsArrayBuffer(event.target.files[0]);

    reader.onload = function(event){

        var data = new Uint8Array(reader.result);

        var work_book = XLSX.read(data, {type:'array'});

        var sheet_name = work_book.SheetNames;
        console.log(sheet_name.length)

        var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {header:1});

        if(sheet_data.length > 0)
        {
            const noOfRows = sheet_data.length;
            const noOfCols = sheet_data[0].length; 
            for(var i = 0; i < noOfRows; i++)
            {
                if(i== 0){
                    for(var j = 0; j < noOfCols; j++){
                        headings.push(sheet_data[i][j]);
                    }
                }
                else
                {
                    var studentId, studentName, studentEmail; 
                    var subjectName, subjectMarks, subjectAttempts, subjectStatus;
                    var studentMarks = [];

                    for(var j = 0; j < noOfCols; j++){
                        if(j==0)
                            studentId = sheet_data[i][j];
                        else if(j==1)
                            studentName = sheet_data[i][j];
                        else if(j==2)
                            studentEmail = sheet_data[i][j];
                        else {
                            subjectName = headings[j];
                            subjectMarks = sheet_data[i][j];
                            j++;
                            subjectAttempts = sheet_data[i][j];
                            subjectStatus = validate(subjectMarks, subjectAttempts, 70);
                            var obj = new Subject(subjectName, subjectMarks, subjectAttempts, subjectStatus);
                            studentMarks.push(obj);
                        }
                    }

                    var  objectiveStatus= "PASS", subjectiveStatus="PASS";
                    for(var k=0;k<studentMarks.length;k++){
                        var subject = studentMarks[k];
                        //console.log(studentName);
                        //console.log(subject);
                        if(k == studentMarks.length-1){
                            //console.log(subject.subjectName, k, "IN SUB");
                            if(subject.subjectStatus == 'FAIL'){
                                subjectiveStatus = "FAIL";
                                break;
                            }
                            
                            if(subject.subjectStatus == 'PENDING'){
                                subjectiveStatus = "PENDING";
                            }

                            if(subject.subjectStatus == 'NA'){
                                subjectiveStatus = "NOT APPEARED";
                            }

                        }else{
                            //console.log(subject.subjectName, k, "IN OBJ");
                            if(subject.subjectStatus == 'FAIL'){
                                objectiveStatus = "FAIL";
                                break;
                            }
                            
                            if(subject.subjectStatus == 'PENDING'){
                                objectiveStatus = "PENDING";
                            }

                            if(subject.subjectStatus == 'NA'){
                                objectiveStatus = "NOT APPEARED";
                            }
                            
                        }
                    }

                    var obj = new Student(studentId, studentName, studentEmail, studentMarks, objectiveStatus, subjectiveStatus);
                    students.push(obj);
                }
                
            }
            console.log(students);
        }
        excelFile.value = '';
    }
});

function validate(marks, attempts, criteria){
    if(marks >= criteria && attempts <= 3){
        return "PASS";
    }
    else if(marks < criteria && attempts<3){
        return "PENDING";
    }
    else if(marks < criteria && attempts == 3)
        return "FAIL";
    else
        return "NA";
}