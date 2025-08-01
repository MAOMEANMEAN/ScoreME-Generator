#include "ExcelUtil.hpp"
#include "MenuUtils.hpp"
#include "GradeUtil.hpp"
#include "Student.hpp"
#include <xlnt/xlnt.hpp>
#include <iostream>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <ctime>
#include <filesystem>

using namespace std;

// Main Excel operations
void ExcelUtils::writeExcel(const std::string& filename, const std::vector<Student>& students) {
    try {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.active_sheet();
        ws.title("Student Grades");

        // Write headers
        auto headers = getExcelHeaders();
        for (size_t i = 0; i < headers.size(); ++i) {
            ws.cell(xlnt::cell_reference(i + 1, 1)).value(headers[i]);
        }

        // Format header row
        formatExcelHeader(ws);

        // Write student data
        for (size_t i = 0; i < students.size(); ++i) {
            writeStudentToExcel(ws, students[i], i + 2);
        }

        wb.save(filename);
        cout << "Excel file '" << filename << "' created successfully!" << endl;
    }
    catch (const exception& e) {
        cerr << "Error writing Excel file: " << e.what() << endl;
    }
}

std::vector<Student> ExcelUtils::readExcelToVector(const std::string& filename) {
    std::vector<Student> students;
    
    try {
        if (!fileExists(filename)) {
            cerr << "File '" << filename << "' does not exist!" << endl;
            return students;
        }

        xlnt::workbook wb;
        wb.load(filename);
        xlnt::worksheet ws = wb.active_sheet();

        // Skip header row and read data
        auto rows = ws.rows();
        auto row_iter = rows.begin();
        if (row_iter != rows.end()) ++row_iter; // Skip header

        int rowNum = 2;
        for (; row_iter != rows.end(); ++row_iter, ++rowNum) {
            try {
                Student student = readStudentFromExcel(ws, rowNum);
                students.push_back(student);
            }
            catch (const exception& e) {
                cerr << "Error reading row " << rowNum << ": " << e.what() << endl;
                continue;
            }
        }
    }
    catch (const exception& e) {
        cerr << "Error reading Excel file: " << e.what() << endl;
    }

    return students;
}

void ExcelUtils::readExcel(const std::string& filename) {
    auto students = readExcelToVector(filename);
    
    if (students.empty()) {
        cout << "No student data found in the file." << endl;
        return;
    }

    cout << "Successfully read " << students.size() << " students from " << filename << endl;
    MenuUtils::displayTable(students);
}

// Enhanced Excel operations
void ExcelUtils::writeExcelWithTimestamp(const std::string& baseFilename, const std::vector<Student>& students) {
    string timestampFilename = generateTimestampFilename(baseFilename);
    writeExcel(timestampFilename, students);
}

void ExcelUtils::createBackup(const std::string& sourceFilename, const std::vector<Student>& students) {
    // Create backup directory if it doesn't exist
    std::filesystem::create_directories("data/backups");
    
    string backupFilename = "data/backups/backup_" + generateTimestampFilename(sourceFilename);
    writeExcel(backupFilename, students);
    
    cout << "Backup created: " << backupFilename << endl;
}

void ExcelUtils::exportGradeReport(const std::string& filename, const std::vector<Student>& students) {
    try {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.active_sheet();
        ws.title("Grade Report");

        // Add report title
        ws.cell("A1").value("GRADE REPORT - " + getCurrentTimestamp());
        ws.merge_cells("A1:H1");
        
        // Add summary statistics
        int totalStudents = students.size();
        int passingStudents = 0;
        double totalAverage = 0.0;
        
        for (const auto& student : students) {
            if (student.hasPassingGrade()) {
                passingStudents++;
            }
            totalAverage += student.getAverageScore();
        }
        
        double classAverage = totalStudents > 0 ? totalAverage / totalStudents : 0.0;
        double passRate = totalStudents > 0 ? (static_cast<double>(passingStudents) / totalStudents) * 100.0 : 0.0;
        
        ws.cell("A3").value("Total Students: " + to_string(totalStudents));
        ws.cell("A4").value("Passing Students: " + to_string(passingStudents));
        ws.cell("A5").value("Pass Rate: " + to_string(static_cast<int>(passRate * 100) / 100.0) + "%");
        ws.cell("A6").value("Class Average: " + to_string(static_cast<int>(classAverage * 100) / 100.0));

        // Write headers starting from row 8
        auto headers = getExcelHeaders();
        for (size_t i = 0; i < headers.size(); ++i) {
            ws.cell(xlnt::cell_reference(i + 1, 8)).value(headers[i]);
        }

        // Write student data
        for (size_t i = 0; i < students.size(); ++i) {
            writeStudentToExcel(ws, students[i], i + 9);
        }

        wb.save(filename);
        cout << "Grade report exported to: " << filename << endl;
    }
    catch (const exception& e) {
        cerr << "Error creating grade report: " << e.what() << endl;
    }
}

// Import operations
bool ExcelUtils::importStudentData(const std::string& filename, std::vector<Student>& students) {
    try {
        if (!fileExists(filename)) {
            cerr << "File '" << filename << "' does not exist!" << endl;
            return false;
        }

        auto importedStudents = readExcelToVector(filename);
        if (importedStudents.empty()) {
            cerr << "No valid student data found in the file." << endl;
            return false;
        }

        // Add imported students to existing vector
        students.insert(students.end(), importedStudents.begin(), importedStudents.end());
        
        cout << "Successfully imported " << importedStudents.size() << " students." << endl;
        return true;
    }
    catch (const exception& e) {
        cerr << "Error importing student data: " << e.what() << endl;
        return false;
    }
}

bool ExcelUtils::validateExcelFormat(const std::string& filename) {
    try {
        if (!fileExists(filename)) {
            return false;
        }

        xlnt::workbook wb;
        wb.load(filename);
        xlnt::worksheet ws = wb.active_sheet();

        // Check if file has expected headers
        auto expectedHeaders = getExcelHeaders();
        for (size_t i = 0; i < expectedHeaders.size(); ++i) {
            try {
                string cellValue = ws.cell(xlnt::cell_reference(i + 1, 1)).to_string();
                if (cellValue != expectedHeaders[i]) {
                    return false;
                }
            }
            catch (...) {
                return false;
            }
        }

        return true;
    }
    catch (...) {
        return false;
    }
}

// Utility methods
std::string ExcelUtils::generateTimestampFilename(const std::string& baseFilename) {
    string timestamp = getCurrentTimestamp();
    
    // Replace spaces and colons with underscores for filename safety
    for (char& c : timestamp) {
        if (c == ' ' || c == ':') {
            c = '_';
        }
    }
    
    size_t dotPos = baseFilename.find_last_of('.');
    if (dotPos != string::npos) {
        return baseFilename.substr(0, dotPos) + "_" + timestamp + baseFilename.substr(dotPos);
    }
    return baseFilename + "_" + timestamp;
}

std::string ExcelUtils::getCurrentTimestamp() {
    auto now = time(0);
    auto tm = *localtime(&now);
    
    ostringstream oss;
    oss << put_time(&tm, "%Y-%m-%d %H:%M:%S");
    return oss.str();
}

bool ExcelUtils::fileExists(const std::string& filename) {
    ifstream file(filename);
    return file.good();
}

std::vector<std::string> ExcelUtils::getExcelHeaders() {
    vector<string> headers = {
        "Student ID", "Name", "Age", "Gender", "Date of Birth", "Email"
    };
    
    // Add subject headers
    auto subjects = GradeUtil::getSubjectNames();
    headers.insert(headers.end(), subjects.begin(), subjects.end());
    
    // Add calculated fields
    headers.insert(headers.end(), {
        "Average Score", "Letter Grade", "GPA", "Remark", "Last Updated"
    });
    
    return headers;
}

// Helper methods for Excel formatting
void ExcelUtils::formatExcelHeader(xlnt::worksheet& ws) {
    auto headers = getExcelHeaders();
    for (size_t i = 0; i < headers.size(); ++i) {
        auto cell = ws.cell(xlnt::cell_reference(i + 1, 1));
        cell.font(xlnt::font().bold(true));
    }
}

void ExcelUtils::writeStudentToExcel(xlnt::worksheet& ws, const Student& student, int row) {
    int col = 1;
    
    // Basic information
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getStudentId());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getName());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getAge());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getGender());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getDateOfBirth());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getEmail());
    
    // Subject scores
    auto scores = student.getSubjectScores();
    for (const auto& score : scores) {
        ws.cell(xlnt::cell_reference(col++, row)).value(score);
    }
    
    // Calculated fields
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getAverageScore());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getLetterGrade());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getGpa());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getRemark());
    ws.cell(xlnt::cell_reference(col++, row)).value(student.getFormattedTimestamp());
}

Student ExcelUtils::readStudentFromExcel(xlnt::worksheet& ws, int row) {
    int col = 1;
    
    try {
        // Read basic information
        string studentId = ws.cell(xlnt::cell_reference(col++, row)).to_string();
        string name = ws.cell(xlnt::cell_reference(col++, row)).to_string();
        int age = ws.cell(xlnt::cell_reference(col++, row)).value<int>();
        string gender = ws.cell(xlnt::cell_reference(col++, row)).to_string();
        string dateOfBirth = ws.cell(xlnt::cell_reference(col++, row)).to_string();
        string email = ws.cell(xlnt::cell_reference(col++, row)).to_string();
        
        // Read subject scores
        vector<double> scores;
        auto subjects = GradeUtil::getSubjectNames();
        for (size_t i = 0; i < subjects.size(); ++i) {
            double score = ws.cell(xlnt::cell_reference(col++, row)).value<double>();
            scores.push_back(score);
        }
        
        // Create student object
        Student student(studentId, name, age, gender, dateOfBirth, email, scores);
        return student;
    }
    catch (const exception& e) {
        throw runtime_error("Error reading student data from row " + to_string(row) + ": " + e.what());
    }
}