package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiOwnerComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiOwnerCompleteRepository extends JpaRepository<BiOwnerComplete,Integer> {
}
