package io.github.prajyotsable;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * Spring Boot entry point.
 * Run this class to start the embedded Tomcat on http://localhost:8080
 * then open http://localhost:8080/index.html in a browser.
 */
@SpringBootApplication
public class SpringApp {

    public static void main(String[] args) {
        SpringApplication.run(SpringApp.class, args);
    }
}
