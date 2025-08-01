#include "Person.hpp"
#include "MenuUtils.hpp"
#include <iostream>

Person::Person(const std::string& username, const std::string& password, const std::string& name)
    : username(username), password(password), name(name) {}

std::string Person::getUsername() const {
    return username;
}

std::string Person::getPassword() const {
    return password;
}

std::string Person::getName() const {
    return name;
}

void Person::setUsername(const std::string& username) {
    this->username = username;
}

void Person::setPassword(const std::string& password) {
    this->password = password;
}

void Person::setName(const std::string& name) {
    this->name = name;
}

std::string Person::getHiddenPassword() const {
    return std::string(password.length(), '*');
}

bool Person::validateCredentials(const std::string& inputUsername, const std::string& inputPassword) const {
    return (username == inputUsername && password == inputPassword);
}