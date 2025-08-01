#pragma once
#include <string>
#include <iostream>

class Person {
protected:
    std::string username;
    std::string password;
    std::string name;
    
public:
    Person() = default;
    Person(const std::string& username, const std::string& password, const std::string& name);
    virtual ~Person() = default;
    
    // Getters
    std::string getUsername() const;
    std::string getPassword() const;
    std::string getName() const;
    
    // Setters
    void setUsername(const std::string& username);
    void setPassword(const std::string& password);
    void setName(const std::string& name);
    
    // Virtual methods for polymorphism
    virtual bool login() = 0;
    virtual void showMenu() = 0;
    virtual std::string getRole() const = 0;
    
    // Utility methods
    std::string getHiddenPassword() const;
    bool validateCredentials(const std::string& inputUsername, const std::string& inputPassword) const;
};