#include "Admin.hpp"
#include "MenuUtils.hpp"
#include "ExcelUtil.hpp"
#include "GradeUtil.hpp"
#include <iostream>
#include <algorithm>
#include <fstream>

// Static member definitions
const std::string Admin::DEFAULT_ADMIN_USERNAME = "admin";
const std::string Admin::DEFAULT_ADMIN_PASSWORD = "admin123";

// Constructors
Admin::Admin() : Person(DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_PASSWORD, "Administrator") {}

Admin::Admin(const std::string& username, const std::string& password, const std::string& name)
    : Person(username, password, name) {}

// Override virtual methods
bool Admin::login() {
    MenuUtils::printHeader("ADMIN LOGIN");
    
    std::string inputUsername = MenuUtils::getStringInput("Username: ");
    std::string inputPassword = MenuUtils::getHiddenInput("Password: ");
    
    if (validateCredentials(inputUsername, inputPassword)) {
        MenuUtils::printSuccess("Login successful! Welcome, Administrator!");
        return true;
    } else {
        MenuUtils::printError("Invalid admin credentials!");
        return false;
    }
}

void Admin::showMenu() {
    int choice;
    std::vector<Student> students = Student::createSampleData(); // Load sample data
    
    do {
        MenuUtils::clearScreen();
        MenuUtils::printAdminMenu();
        choice = MenuUtils::getMenuChoice(6);
        
        switch (choice) {
            case 1:
                manageStudents(students);
                break;
            case 2:
                importExcelData(students, "data/students.xlsx");
                MenuUtils::pauseScreen();
                break;
            case 3:
                exportData(students, "data/grade_report.xlsx");
                MenuUtils::pauseScreen();
                break;
            case 4:
                backupData(students);
                MenuUtils::pauseScreen();
                break;
            case 5:
                MenuUtils::printInfo("Signing out from admin dashboard...");
                return;
            case 6:
                MenuUtils::printInfo("Returning to main menu...");
                return;
        }
        
        if (choice != 5 && choice != 6) {
            if (!MenuUtils::askContinue()) {
                break;
            }
        }
    } while (choice != 5 && choice != 6);
}

std::string Admin::getRole() const {
    return "Administrator";
}

// Admin-specific methods
void Admin::manageStudents(std::vector<Student>& students) {
    int choice;
    do {
        MenuUtils::clearScreen();
        MenuUtils::printHeader("STUDENT MANAGEMENT");
        
        std::vector<std::string> studentMenu = {
            "View All Students",
            "Add New Student",
            "Edit Student Info",
            "Delete Student",
            "Search Student",
            "Show Failing Students",
            "Sort Students by Score",
            "Back to Admin Dashboard"
        };
        
        MenuUtils::printMenu(studentMenu);
        choice = MenuUtils::getMenuChoice(8);
        
        switch (choice) {
            case 1:
                viewAllStudents(students);
                break;
            case 2:
                addNewStudent(students);
                break;
            case 3:
                editStudentInfo(students);
                break;
            case 4:
                deleteStudent(students);
                break;
            case 5:
                searchStudent(students);
                break;
            case 6:
                showFailingStudents(students);
                break;
            case 7:
                sortStudentsByScore(students);
                break;
            case 8:
                return;
        }
        
        if (choice != 8) {
            MenuUtils::pauseScreen();
        }
    } while (choice != 8);
}

void Admin::viewAllStudents(const std::vector<Student>& students) {
    MenuUtils::printHeader("ALL STUDENTS");
    if (students.empty()) {
        MenuUtils::printWarning("No students found!");
        return;
    }
    MenuUtils::displayTable(students);
}

void Admin::addNewStudent(std::vector<Student>& students) {
    MenuUtils::printHeader("ADD NEW STUDENT");
    
    std::string studentId = MenuUtils::getStringInput("Student ID: ");
    if (!isValidStudentId(studentId, students)) {
        MenuUtils::printError("Student ID already exists!");
        return;
    }
    
    std::string name = MenuUtils::getStringInput("Student Name: ");
    if (isStudentExists(name, students)) {
        MenuUtils::printError("Student with this name already exists!");
        return;
    }
    
    int age = MenuUtils::getIntInput("Age: ");
    std::string gender = MenuUtils::getStringInput("Gender: ");
    std::string dob = MenuUtils::getStringInput("Date of Birth (YYYY-MM-DD): ");
    std::string email = MenuUtils::getStringInput("Email: ");
    
    std::vector<double> scores(7);
    std::vector<std::string> subjects = GradeUtil::getSubjectNames();
    
    MenuUtils::printInfo("Enter scores for all subjects:");
    for (size_t i = 0; i < subjects.size(); ++i) {
        scores[i] = MenuUtils::getDoubleInput(subjects[i] + " score: ");
        if (!GradeUtil::isValidScore(scores[i])) {
            MenuUtils::printError("Invalid score! Score must be between 0-100.");
            i--; // Retry current subject
        }
    }
    
    students.emplace_back(studentId, name, age, gender, dob, email, scores);
    MenuUtils::printSuccess("Student added successfully!");
    
    // Save updated data to Excel
    try {
        ExcelUtils::writeExcel("data/students.xlsx", students);
        MenuUtils::printInfo("Data saved to Excel file.");
    }
    catch (const std::exception& e) {
        MenuUtils::printWarning("Student added but failed to save to Excel: " + std::string(e.what()));
    }
}

void Admin::editStudentInfo(std::vector<Student>& students) {
    MenuUtils::printHeader("EDIT STUDENT INFO");
    
    std::string searchId = MenuUtils::getStringInput("Enter Student ID to edit: ");
    Student* student = findStudentById(students, searchId);
    
    if (!student) {
        MenuUtils::printError("Student not found!");
        return;
    }
    
    MenuUtils::printInfo("Current student info:");
    MenuUtils::displayStudentDetails(*student);
    
    std::vector<std::string> editMenu = {
        "Edit Name",
        "Edit Age",
        "Edit Gender",
        "Edit Date of Birth",
        "Edit Email",
        "Edit Scores",
        "Cancel"
    };
    
    MenuUtils::printMenu(editMenu);
    int choice = MenuUtils::getMenuChoice(7);
    
    switch (choice) {
        case 1: {
            std::string newName = MenuUtils::getStringInput("New name: ");
            student->setName(newName);
            break;
        }
        case 2: {
            int newAge = MenuUtils::getIntInput("New age: ");
            student->setAge(newAge);
            break;
        }
        case 3: {
            std::string newGender = MenuUtils::getStringInput("New gender: ");
            student->setGender(newGender);
            break;
        }
        case 4: {
            std::string newDob = MenuUtils::getStringInput("New date of birth (YYYY-MM-DD): ");
            student->setDateOfBirth(newDob);
            break;
        }
        case 5: {
            std::string newEmail = MenuUtils::getStringInput("New email: ");
            student->setEmail(newEmail);
            break;
        }
        case 6: {
            std::vector<double> newScores(7);
            std::vector<std::string> subjects = GradeUtil::getSubjectNames();
            
            for (size_t i = 0; i < subjects.size(); ++i) {
                newScores[i] = MenuUtils::getDoubleInput(subjects[i] + " new score: ");
                if (!GradeUtil::isValidScore(newScores[i])) {
                    MenuUtils::printError("Invalid score! Score must be between 0-100.");
                    i--; // Retry current subject
                }
            }
            student->setSubjectScores(newScores);
            break;
        }
        case 7:
            return;
    }
    
    if (choice != 7) {
        MenuUtils::printSuccess("Student information updated successfully!");
        
        // Save updated data to Excel
        try {
            ExcelUtils::writeExcel("data/students.xlsx", students);
            MenuUtils::printInfo("Data saved to Excel file.");
        }
        catch (const std::exception& e) {
            MenuUtils::printWarning("Student updated but failed to save to Excel: " + std::string(e.what()));
        }
    }
}

void Admin::deleteStudent(std::vector<Student>& students) {
    MenuUtils::printHeader("DELETE STUDENT");
    
    std::string searchId = MenuUtils::getStringInput("Enter Student ID to delete: ");
    auto it = std::find_if(students.begin(), students.end(),
        [&searchId](const Student& s) { return s.getStudentId() == searchId; });
    
    if (it != students.end()) {
        MenuUtils::printInfo("Student found:");
        MenuUtils::displayStudentDetails(*it);
        
        std::string confirm = MenuUtils::getStringInput("Are you sure you want to delete this student? (yes/no): ");
        if (confirm == "yes" || confirm == "y" || confirm == "Y") {
            students.erase(it);
            MenuUtils::printSuccess("Student deleted successfully!");
            
            // Save updated data to Excel
            try {
                ExcelUtils::writeExcel("data/students.xlsx", students);
                MenuUtils::printInfo("Data saved to Excel file.");
            }
            catch (const std::exception& e) {
                MenuUtils::printWarning("Student deleted but failed to save to Excel: " + std::string(e.what()));
            }
        } else {
            MenuUtils::printInfo("Deletion cancelled.");
        }
    } else {
        MenuUtils::printError("Student not found!");
    }
}

void Admin::searchStudent(const std::vector<Student>& students) {
    MenuUtils::printHeader("SEARCH STUDENT");
    
    std::string searchTerm = MenuUtils::getStringInput("Enter Student ID or Name: ");
    bool found = false;
    
    for (const auto& student : students) {
        if (student.getStudentId() == searchTerm || student.getName() == searchTerm) {
            MenuUtils::displayStudentDetails(student);
            found = true;
            break;
        }
    }
    
    if (!found) {
        MenuUtils::printError("Student not found!");
    }
}

void Admin::showFailingStudents(const std::vector<Student>& students) {
    MenuUtils::printHeader("FAILING STUDENTS");
    
    std::vector<Student> failingStudents;
    for (const auto& student : students) {
        if (!student.hasPassingGrade()) {
            failingStudents.push_back(student);
        }
    }
    
    if (failingStudents.empty()) {
        MenuUtils::printSuccess("No failing students found!");
    } else {
        MenuUtils::displayFailingStudents(failingStudents);
    }
}

void Admin::sortStudentsByScore(std::vector<Student>& students) {
    MenuUtils::printHeader("SORT STUDENTS BY SCORE");
    
    std::string order = MenuUtils::getStringInput("Sort order (asc/desc): ");
    
    if (order == "asc") {
        std::sort(students.begin(), students.end(),
            [](const Student& a, const Student& b) {
                return a.getAverageScore() < b.getAverageScore();
            });
    } else {
        std::sort(students.begin(), students.end(),
            [](const Student& a, const Student& b) {
                return a.getAverageScore() > b.getAverageScore();
            });
    }
    
    MenuUtils::printSuccess("Students sorted successfully!");
    MenuUtils::displayTable(students);
    
    // Save sorted data to Excel
    try {
        ExcelUtils::writeExcel("data/students.xlsx", students);
        MenuUtils::printInfo("Sorted data saved to Excel file.");
    }
    catch (const std::exception& e) {
        MenuUtils::printWarning("Students sorted but failed to save to Excel: " + std::string(e.what()));
    }
}

// Data management methods
void Admin::importExcelData(std::vector<Student>& students, const std::string& filename) {
    MenuUtils::printHeader("IMPORT EXCEL DATA");
    
    if (ExcelUtils::importStudentData(filename, students)) {
        MenuUtils::printSuccess("Data imported successfully from " + filename + "!");
        MenuUtils::printInfo("Total students now: " + std::to_string(students.size()));
    } else {
        MenuUtils::printError("Failed to import data from " + filename);
    }
}

void Admin::exportData(const std::vector<Student>& students, const std::string& filename) {
    MenuUtils::printHeader("EXPORT DATA");
    
    try {
        ExcelUtils::exportGradeReport(filename, students);
        MenuUtils::printSuccess("Grade report exported successfully to " + filename + "!");
    }
    catch (const std::exception& e) {
        MenuUtils::printError("Failed to export data: " + std::string(e.what()));
    }
}

void Admin::backupData(const std::vector<Student>& students) {
    MenuUtils::printHeader("BACKUP DATA");
    
    try {
        ExcelUtils::createBackup("students.xlsx", students);
        MenuUtils::printSuccess("Data backup created successfully!");
    }
    catch (const std::exception& e) {
        MenuUtils::printError("Failed to create backup: " + std::string(e.what()));
    }
}

// Utility methods
bool Admin::isValidStudentId(const std::string& id, const std::vector<Student>& students) {
    return std::find_if(students.begin(), students.end(),
        [&id](const Student& s) { return s.getStudentId() == id; }) == students.end();
}

bool Admin::isStudentExists(const std::string& name, const std::vector<Student>& students) {
    return std::find_if(students.begin(), students.end(),
        [&name](const Student& s) { return s.getName() == name; }) != students.end();
}

Student* Admin::findStudentById(std::vector<Student>& students, const std::string& id) {
    auto it = std::find_if(students.begin(), students.end(),
        [&id](const Student& s) { return s.getStudentId() == id; });
    return (it != students.end()) ? &(*it) : nullptr;
}

Student* Admin::findStudentByName(std::vector<Student>& students, const std::string& name) {
    auto it = std::find_if(students.begin(), students.end(),
        [&name](const Student& s) { return s.getName() == name; });
    return (it != students.end()) ? &(*it) : nullptr;
}