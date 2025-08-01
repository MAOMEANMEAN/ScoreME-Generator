#include "Student.hpp"
#include "GradeUtil.hpp"
#include "MenuUtils.hpp"
#include "ExcelUtil.hpp"
#include <iostream>
#include <algorithm>
#include <ctime>
#include <iomanip>
#include <sstream>

// Constructors
Student::Student(const std::string& name, int age) : age(age), averageScore(0.0), gpa(0.0) {
    this->name = name;
    subjectScores.resize(7, 0.0);
    updateTimestamp();
}

Student::Student(const std::string& studentId, const std::string& name, int age,
    const std::string& gender, const std::string& dateOfBirth,
    const std::string& email, const std::vector<double>& scores) : studentId(studentId), age(age), gender(gender), dateOfBirth(dateOfBirth),email(email), subjectScores(scores) {
    this->name = name;
    updateAllGrades();
    updateTimestamp();
}

Student::Student(const std::string& username, const std::string& password,
    const std::string& studentId, const std::string& name, int age,
    const std::string& gender, const std::string& dateOfBirth,
    const std::string& email, const std::vector<double>& scores) : Person(username, password, name), studentId(studentId), age(age),gender(gender), dateOfBirth(dateOfBirth), email(email), subjectScores(scores) {
    updateAllGrades();
    updateTimestamp();
}

// Getters
std::string Student::getStudentId() const { return studentId; }
int Student::getAge() const { return age; }
std::string Student::getGender() const { return gender; }
std::string Student::getDateOfBirth() const { return dateOfBirth; }
std::string Student::getEmail() const { return email; }
std::vector<double> Student::getSubjectScores() const { return subjectScores; }
double Student::getAverageScore() const { return averageScore; }
std::string Student::getLetterGrade() const { return letterGrade; }
double Student::getGpa() const { return gpa; }
std::string Student::getRemark() const { return remark; }
std::time_t Student::getLastUpdated() const { return lastUpdated; }

// Setters
void Student::setStudentId(const std::string& id) {
    studentId = id;
    updateTimestamp();
}

void Student::setAge(int age) {
    this->age = age;
    updateTimestamp();
}

void Student::setGender(const std::string& gender) {
    this->gender = gender;
    updateTimestamp();
}

void Student::setDateOfBirth(const std::string& dob) {
    dateOfBirth = dob;
    updateTimestamp();
}

void Student::setEmail(const std::string& email) {
    this->email = email;
    updateTimestamp();
}

void Student::setSubjectScores(const std::vector<double>& scores) {
    subjectScores = scores;
    updateAllGrades();
    updateTimestamp();
}

// Calculation methods
void Student::calculateAverageScore() {
    averageScore = GradeUtil::calculateAverage(subjectScores);
}

void Student::assignLetterGrade() {
    letterGrade = GradeUtil::assignLetterGrade(averageScore);
}

void Student::calculateGpa() {
    gpa = GradeUtil::calculateGpa(averageScore);
}

void Student::assignRemark() {
    remark = GradeUtil::assignRemark(averageScore);
}

void Student::updateAllGrades() {
    calculateAverageScore();
    assignLetterGrade();
    calculateGpa();
    assignRemark();
}

// Override virtual methods
bool Student::login() {
    MenuUtils::printHeader("STUDENT LOGIN");
    
    std::string inputUsername = MenuUtils::getStringInput("Username: ");
    std::string inputPassword = MenuUtils::getHiddenInput("Password: ");
    
    if (validateCredentials(inputUsername, inputPassword)) {
        MenuUtils::printSuccess("Login successful! Welcome, " + name + "!");
        return true;
    } else {
        MenuUtils::printError("Invalid credentials!");
        return false;
    }
}

void Student::showMenu() {
    int choice;
    do {
        MenuUtils::clearScreen();
        MenuUtils::printStudentMenu();
        choice = MenuUtils::getMenuChoice(3);
        
        switch (choice) {
            case 1: {
                MenuUtils::printHeader("SEARCH YOUR DATA");
                std::string searchTerm = MenuUtils::getStringInput("Enter your Student ID or Name: ");
                
                if (searchTerm == studentId || searchTerm == name) {
                    MenuUtils::displayStudentDetails(*this);
                } else {
                    MenuUtils::printError("No matching record found!");
                }
                
                MenuUtils::pauseScreen();
                break;
            }
            case 2: {
                MenuUtils::printHeader("YOUR GRADE REPORT");
                MenuUtils::displayStudentDetails(*this);
                MenuUtils::pauseScreen();
                break;
            }
            case 3:
                MenuUtils::printInfo("Signing out...");
                break;
        }
    } while (choice != 3 && MenuUtils::askContinue());
}

std::string Student::getRole() const {
    return "Student";
}

// Utility methods
void Student::updateTimestamp() {
    lastUpdated = std::time(nullptr);
}

std::string Student::getFormattedTimestamp() const {
    std::tm* tm = std::localtime(&lastUpdated);
    std::ostringstream oss;
    oss << std::put_time(tm, "%Y-%m-%d %H:%M:%S");
    return oss.str();
}

bool Student::hasPassingGrade() const {
    return GradeUtil::isPassingGrade(averageScore);
}

// Static method for sample data
std::vector<Student> Student::createSampleData() {
    std::vector<Student> students;
    
        students.emplace_back("STU001", "John Smith", 20, "Male", "2003-05-15", "john.smith@email.com",
                            std::vector<double>{85.5, 78.0, 92.3, 88.7, 76.5, 90.1, 82.8});
        students.emplace_back("STU002", "Emily Johnson", 19, "Female", "2004-08-22", "emily.johnson@email.com",
                            std::vector<double>{92.1, 89.5, 87.3, 91.2, 88.9, 85.7, 90.4});
        students.emplace_back("STU003", "Michael Brown", 21, "Male", "2002-12-10", "michael.brown@email.com",
                            std::vector<double>{76.8, 82.3, 79.5, 85.2, 81.7, 78.9, 80.1});
        students.emplace_back("STU004", "Sarah Davis", 20, "Female", "2003-03-18", "sarah.davis@email.com",
                            std::vector<double>{88.9, 91.2, 86.5, 89.8, 87.3, 90.7, 88.1});
        students.emplace_back("STU005", "David Wilson", 19, "Male", "2004-07-25", "david.wilson@email.com",
                            std::vector<double>{65.2, 58.9, 62.1, 59.8, 61.5, 63.7, 60.3});
        students.emplace_back("STU006", "Lisa Miller", 20, "Female", "2003-11-30", "lisa.miller@email.com",
                            std::vector<double>{94.5, 96.2, 93.8, 95.1, 97.3, 92.9, 94.7});
        students.emplace_back("STU007", "James Taylor", 21, "Male", "2002-09-14", "james.taylor@email.com",
                            std::vector<double>{78.3, 81.5, 77.9, 82.1, 79.7, 80.4, 78.8});
        students.emplace_back("STU008", "Jennifer Anderson", 19, "Female", "2004-04-08", "jennifer.anderson@email.com",
                            std::vector<double>{91.7, 88.3, 90.5, 87.9, 89.1, 92.4, 90.8});
        students.emplace_back("STU009", "Robert Thomas", 20, "Male", "2003-01-20", "robert.thomas@email.com",
                            std::vector<double>{55.8, 52.3, 58.1, 54.9, 56.7, 53.5, 55.2});
        students.emplace_back("STU010", "Jessica Martinez", 19, "Female", "2004-06-12", "jessica.martinez@email.com",
                            std::vector<double>{87.4, 85.9, 89.2, 86.7, 88.5, 87.1, 86.8});
    
    // Update all grades for sample data
    for (auto& student : students) {
        student.updateAllGrades();
    }
    
    return students;
}