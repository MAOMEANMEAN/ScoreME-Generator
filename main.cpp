#include <iostream>
#include <vector>
#include <memory>
#include <string>
#include "Student.hpp"
#include "Admin.hpp"
#include "Person.hpp"
#include "ExcelUtil.hpp"
#include "MenuUtils.hpp"

using namespace std;

class ScoreMEApp {
private:
    vector<Student> registeredStudents;
    Admin admin;
    
public:
    ScoreMEApp() {
        // Initialize with sample student data for login testing
        initializeStudentAccounts();
        
        // Create sample Excel files with data
        createSampleExcelFiles();
    }
    
    void initializeStudentAccounts() {
        // Create some student accounts with login credentials
        registeredStudents = Student::createSampleData();
        
        // Set login credentials for first few students for testing
        if (registeredStudents.size() >= 3) {
            registeredStudents[0].setUsername("john.smith");
            registeredStudents[0].setPassword("pass123");
            
            registeredStudents[1].setUsername("emily.johnson");
            registeredStudents[1].setPassword("pass456");
            
            registeredStudents[2].setUsername("michael.brown");
            registeredStudents[2].setPassword("pass789");
        }
    }
    
    void createSampleExcelFiles() {
        try {
            // Create sample Excel files with student data
            ExcelUtils::writeExcel("data/students.xlsx", registeredStudents);
            ExcelUtils::writeExcel("data/persons.xlsx", registeredStudents);
            
            cout << "Sample Excel files created successfully!" << endl;
        }
        catch (const exception& e) {
            cerr << "Error creating sample Excel files: " << e.what() << endl;
        }
    }
    
    void run() {
        MenuUtils::printWelcome();
        
        int choice;
        do {
            MenuUtils::clearScreen();
            MenuUtils::printMainMenu();
            choice = MenuUtils::getMenuChoice(3);
            
            switch (choice) {
                case 1:
                    handleAdminLogin();
                    break;
                case 2:
                    handleStudentLogin();
                    break;
                case 3:
                    MenuUtils::printInfo("Thank you for using ScoreME Generator! Goodbye! 👋");
                    break;
            }
            
            if (choice != 3) {
                MenuUtils::pauseScreen();
            }
            
        } while (choice != 3);
    }
    
private:
    void handleAdminLogin() {
        MenuUtils::clearScreen();
        
        if (admin.login()) {
            // FIXED: Use showMenuWithData() method instead of showMenu()
            admin.showMenuWithData(registeredStudents);
        } else {
            MenuUtils::printError("Admin login failed!");
        }
    }
    
    void handleStudentLogin() {
        MenuUtils::clearScreen();
        MenuUtils::printHeader("STUDENT LOGIN");
        
        string username = MenuUtils::getStringInput("Username: ");
        string password = MenuUtils::getHiddenInput("Password: ");
        
        // Find student with matching credentials
        Student* loggedInStudent = nullptr;
        for (auto& student : registeredStudents) {
            if (student.validateCredentials(username, password)) {
                loggedInStudent = &student;
                break;
            }
        }
        
        if (loggedInStudent) {
            MenuUtils::printSuccess("Login successful! Welcome, " + loggedInStudent->getName() + "!");
            MenuUtils::pauseScreen();
            loggedInStudent->showMenu();
        } else {
            MenuUtils::printError("Invalid student credentials!");
        }
    }
};

// Function to create sample data when called with command line argument
void createSampleDataFiles() {
    try {
        auto students = Student::createSampleData();
        ExcelUtils::writeExcel("data/students.xlsx", students);
        ExcelUtils::writeExcel("data/persons.xlsx", students);
        cout << "Sample Excel files created successfully in data/ directory!" << endl;
    }
    catch (const exception& e) {
        cerr << "Error creating sample files: " << e.what() << endl;
    }
}

int main(int argc, char* argv[]) {
    
    try {
        // Check if we're being called to create sample data
        if (argc > 1 && string(argv[1]) == "--create-sample-data") {
            createSampleDataFiles();
            return 0;
        }
        
        ScoreMEApp app;
        app.run();
    } catch (const exception& e) {
        MenuUtils::printError("An error occurred: " + string(e.what()));
        return 1;
    }
    
    return 0;
}