package org.com.test;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurerAdapter;

@SpringBootApplication
public class SpringBootContactApplication extends WebMvcConfigurerAdapter {

	public static void main(String[] args) 
	{
		SpringApplication.run(SpringBootContactApplication.class, args);
	}
}