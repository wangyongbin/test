package com.poi.testpoi;

import com.poi.testpoi.pojo.BiLocalComplete;
import com.poi.testpoi.service.*;
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

	@Autowired
	private BiGroupCompleteService biGroupCompleteService;

	@Autowired
	private BiLocalCompleteService biLocalCompleteService;

	@Autowired
	private BiInnerCompleteService biInnerCompleteService;

	@Autowired
	private BiAllService biAllService;

	@Test
	public void contextLoads() {
	}


	@Test
	public void testImportDept() throws Exception {
        biDeptService.batchImport("import_dept.xlsx",null);
	}

	@Test
	public void testImportGroup() throws Exception {
		biGroupCompleteService.batchImport("import_group.xls",null);
	}

	@Test
	public void testImportLocal() throws Exception {
		biLocalCompleteService.batchImport("import_local.xls",null);
	}

	@Test
	public void testImportInner() throws Exception {
		biInnerCompleteService.batchImport("import_inner.xls",null);
	}

	@Test
	public void testImportAll() throws Exception {
		biAllService.batchImport("import_all.xlsx",null);
	}


}
