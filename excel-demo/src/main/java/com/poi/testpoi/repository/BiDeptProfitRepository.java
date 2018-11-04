package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiDeptProfit;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiDeptProfitRepository extends JpaRepository<BiDeptProfit,Integer> {
}
