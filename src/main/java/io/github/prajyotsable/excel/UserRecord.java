package io.github.prajyotsable.excel;

import java.time.LocalDate;

public class UserRecord {

    private final Long id;
    private final String firstName;
    private final String lastName;
    private final String email;
    private final Integer age;
    private final LocalDate createdDate;

    public UserRecord(Long id, String firstName, String lastName, String email, Integer age, LocalDate createdDate) {
        this.id = id;
        this.firstName = firstName;
        this.lastName = lastName;
        this.email = email;
        this.age = age;
        this.createdDate = createdDate;
    }

    public Long getId() {
        return id;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public String getEmail() {
        return email;
    }

    public Integer getAge() {
        return age;
    }

    public LocalDate getCreatedDate() {
        return createdDate;
    }

}