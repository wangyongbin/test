package com.poi.testpoi.repository;

import com.poi.testpoi.pojo.BiDeptCompleteDetail;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface BiDeptCompleteDetailRepository extends JpaRepository<BiDeptCompleteDetail,String> {
}
