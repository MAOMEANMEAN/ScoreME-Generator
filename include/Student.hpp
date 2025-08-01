#pragma once
#include <string>
#include <vector>
#include <ctime>
#include "Person.hpp"

class Student : public Person {
private:
    std::string studentId;
    int age;
    std::string gender;
    std::string dateOfBirth;
    std::string email;
    std::vector<double> subjectScores; // 7 subjects
    double averageScore;
    std::string letterGrade;
    double gpa;
    std::string remark;
    std::time_t lastUpdated;

public:
    // Constructors
    Student() = default;
    Student(const std::string& name, int age);
    Student(const std::string& studentId, const std::string& name, int age,
            const std::string& gender, const std::string& dateOfBirth,
            const std::string& email, const std::vector<double>& scores);
    Student(const std::string& username, const std::string& password,
            const std::string& studentId, const std::string& name, int age,
            const std::string& gender, const std::string& dateOfBirth,
            const std::string& email, const std::vector<double>& scores);

    // Getters
    std::string getStudentId() const;
    int getAge() const;
    std::string getGender() const;
    std::string getDateOfBirth() const;
    std::string getEmail() const;
    std::vector<double> getSubjectScores() const;
    double getAverageScore() const;
    std::string getLetterGrade() const;
    double getGpa() const;
    std::string getRemark() const;
    std::time_t getLastUpdated() const;

    // Setters
    void setStudentId(const std::string& id);
    void setAge(int age);
    void setGender(const std::string& gender);
    void setDateOfBirth(const std::string& dob);
    void setEmail(const std::string& email);
    void setSubjectScores(const std::vector<double>& scores);

    // Calculation methods
    void calculateAverageScore();
    void assignLetterGrade();
    void calculateGpa();
    void assignRemark();
    void updateAllGrades();

    // Override virtual methods from Person
    bool login() override;
    void showMenu() override;
    std::string getRole() const override;

    // Utility methods
    void updateTimestamp();
    std::string getFormattedTimestamp() const;
    bool hasPassingGrade() const;
    
    // Static method for creating sample data
    static std::vector<Student> createSampleData();
};