package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiDomainComplete;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiDomainCompleteRepository extends JpaRepository<BiDomainComplete,Integer> {
}
