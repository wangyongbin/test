package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiDeptComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiDeptCompleteRepository extends JpaRepository<BiDeptComplete,Integer> {
}
