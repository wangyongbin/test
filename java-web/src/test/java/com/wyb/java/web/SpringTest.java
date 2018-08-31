package com.wyb.java.web;

import org.junit.Test;

public class SpringTest {

    @Test
    public void testSpringGetSystemEnv() {
        System.out.println(System.getenv("test.test"));
    }
}
