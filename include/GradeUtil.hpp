#pragma once
#include <string>
#include <vector>

class GradeUtil {
public:
    // Grade calculation methods
    static double calculateAverage(const std::vector<double>& scores);
    static std::string assignLetterGrade(double average);
    static double calculateGpa(double average);
    static std::string assignRemark(double average);
    
    // Grade validation
    static bool isValidScore(double score);
    static bool isPassingGrade(double average);
    
    // Subject names
    static std::vector<std::string> getSubjectNames();
    
    // Grade ranges and constants
    static const double MIN_SCORE;
    static const double MAX_SCORE;
    static const double PASSING_THRESHOLD;
    
private:
    // Grade boundaries
    static const double GRADE_A_THRESHOLD;
    static const double GRADE_B_THRESHOLD;
    static const double GRADE_C_THRESHOLD;
    static const double GRADE_D_THRESHOLD;
    static const double GRADE_E_THRESHOLD;
};