package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiInnerComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiInnerCompleteRepository extends JpaRepository<BiInnerComplete,Integer> {
}
