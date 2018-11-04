package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiIndustryComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiIndustryCompleteRepository extends JpaRepository<BiIndustryComplete,Integer> {
}
