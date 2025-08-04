#include "GradeUtil.hpp"
#include <numeric>
#include <algorithm>

// Static member definitions - UPDATED PASSING THRESHOLD
const double GradeUtil::MIN_SCORE = 0.0;
const double GradeUtil::MAX_SCORE = 100.0;
const double GradeUtil::PASSING_THRESHOLD = 50.0;  // CHANGED FROM 60.0 TO 50.0
const double GradeUtil::GRADE_A_THRESHOLD = 90.0;
const double GradeUtil::GRADE_B_THRESHOLD = 80.0;
const double GradeUtil::GRADE_C_THRESHOLD = 70.0;
const double GradeUtil::GRADE_D_THRESHOLD = 60.0;
const double GradeUtil::GRADE_E_THRESHOLD = 50.0;  // E starts at 50 (passing)

double GradeUtil::calculateAverage(const std::vector<double>& scores) {
    if (scores.empty()) return 0.0;
    
    double sum = std::accumulate(scores.begin(), scores.end(), 0.0);
    return sum / scores.size();
}

std::string GradeUtil::assignLetterGrade(double average) {
    if (average >= GRADE_A_THRESHOLD) return "A";
    else if (average >= GRADE_B_THRESHOLD) return "B";
    else if (average >= GRADE_C_THRESHOLD) return "C";
    else if (average >= GRADE_D_THRESHOLD) return "D";
    else if (average >= GRADE_E_THRESHOLD) return "E";  // 50-59 = E (Pass)
    else return "F";  // Below 50 = F (Fail)
}

double GradeUtil::calculateGpa(double average) {
    if (average >= GRADE_A_THRESHOLD) return 4.0;
    else if (average >= GRADE_B_THRESHOLD) return 3.0;
    else if (average >= GRADE_C_THRESHOLD) return 2.0;
    else if (average >= GRADE_D_THRESHOLD) return 1.0;
    else if (average >= GRADE_E_THRESHOLD) return 0.5;  // E grade gets 0.5 GPA
    else return 0.0;  // F grade gets 0.0 GPA
}

// UPDATED: Only F grade (below 50) is considered fail
std::string GradeUtil::assignRemark(double average) {
    return (average >= PASSING_THRESHOLD) ? "Pass" : "Fail";  // 50+ is Pass, below 50 is Fail
}

bool GradeUtil::isValidScore(double score) {
    return score >= MIN_SCORE && score <= MAX_SCORE;
}

// UPDATED: Passing grade is now 50+
bool GradeUtil::isPassingGrade(double average) {
    return average >= PASSING_THRESHOLD;  // 50+ is passing
}

std::vector<std::string> GradeUtil::getSubjectNames() {
    return {
        "Mathematics",
        "Physics", 
        "Chemistry",
        "Biology",
        "English",
        "History",
        "Computer Science"
    };
}