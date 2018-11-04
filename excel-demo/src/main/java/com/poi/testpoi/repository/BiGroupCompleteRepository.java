package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiGroupComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiGroupCompleteRepository extends JpaRepository<BiGroupComplete,Integer> {
}
