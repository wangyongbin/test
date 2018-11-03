package com.poi.testpoi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@ComponentScan(value = {"com.poi.testpoi"})
public class TestpoiApplication {

	public static void main(String[] args) {
		SpringApplication.run(TestpoiApplication.class, args);
	}
}
