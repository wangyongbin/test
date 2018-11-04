package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiLocalComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiLocalCompleteRepository extends JpaRepository<BiLocalComplete,Integer> {
}
