package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiDeptCostDetail;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiDeptCostDetailRepository extends JpaRepository<BiDeptCostDetail,String> {
}
