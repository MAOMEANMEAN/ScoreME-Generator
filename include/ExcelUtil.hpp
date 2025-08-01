#pragma once
#include <string>
#include <vector>
#include "Student.hpp"

// Forward declaration to avoid circular dependency
namespace xlnt {
    class worksheet;
}

class ExcelUtils {
public:
    // Main Excel operations
    static void writeExcel(const std::string& filename, const std::vector<Student>& students);
    static std::vector<Student> readExcelToVector(const std::string& filename);
    static void readExcel(const std::string& filename);
    
    // Enhanced Excel operations
    static void writeExcelWithTimestamp(const std::string& baseFilename, const std::vector<Student>& students);
    static void createBackup(const std::string& sourceFilename, const std::vector<Student>& students);
    static void exportGradeReport(const std::string& filename, const std::vector<Student>& students);
    
    // Import operations
    static bool importStudentData(const std::string& filename, std::vector<Student>& students);
    static bool validateExcelFormat(const std::string& filename);
    
    // Utility methods
    static std::string generateTimestampFilename(const std::string& baseFilename);
    static std::string getCurrentTimestamp();
    static bool fileExists(const std::string& filename);
    
    // Excel column headers
    static std::vector<std::string> getExcelHeaders();
    
private:
    // Helper methods for Excel formatting
    static void formatExcelHeader(xlnt::worksheet& ws);
    static void writeStudentToExcel(xlnt::worksheet& ws, const Student& student, int row);
    static Student readStudentFromExcel(xlnt::worksheet& ws, int row);
};