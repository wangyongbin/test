package com.poi.testpoi;

import com.poi.testpoi.service.BiDeptService;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

@RunWith(SpringRunner.class)
@SpringBootTest
public class TestpoiApplicationTests {

	@Autowired
	private BiDeptService biDeptService;

	@Test
	public void contextLoads() {
	}



	@Test
	public void testBatchImport() throws Exception {

        biDeptService.batchImport("import_dept.xlsx",null);
	}



}
