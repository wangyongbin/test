package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiDeptCost;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiDeptCostRepository extends JpaRepository<BiDeptCost,Integer> {
}
