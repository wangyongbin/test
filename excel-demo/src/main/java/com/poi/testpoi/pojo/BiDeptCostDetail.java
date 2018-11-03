package com.poi.testpoi.pojo;

import javax.persistence.*;

/**
 * 部门成本执行情况
 */
@Entity
@Table(name = "bi_dept_cost_detail")
public class BiDeptCostDetail {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    @Column(name = "id")
    private Integer id;

    //所属年份
    @Column(name = "count_year", nullable = true, length = 45)
    private String countYear;

    //部门名称
    @Column(name = "dept_name", nullable = true, length = 45)
    private String deptName;

    //经济指标
    @Column(name = "type_name", nullable = true, length = 45)
    private String typeName;

    //金额单位
    @Column(name = "value_unit", nullable = true, length = 45)
    private String valueUnit;

    //月份
    @Column(name = "count_month", nullable = true, length = 45)
    private String countMonth;

    //年度目标值
    @Column(name = "year_target")
    private Double yearTarget;

    //本月发生值
    @Column(name = "exec_val")
    private Double execVal;

    //本月累计发生值
    @Column(name = "exec_sum_val")
    private Double execSumVal;

    //本月发生率
    @Column(name = "exec_sum_rate")
    private Double execSumRate;



    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getCountYear() {
        return countYear;
    }

    public void setCountYear(String countYear) {
        this.countYear = countYear;
    }

    public String getDeptName() {
        return deptName;
    }

    public void setDeptName(String deptName) {
        this.deptName = deptName;
    }

    public String getTypeName() {
        return typeName;
    }

    public void setTypeName(String typeName) {
        this.typeName = typeName;
    }

    public String getValueUnit() {
        return valueUnit;
    }

    public void setValueUnit(String valueUnit) {
        this.valueUnit = valueUnit;
    }

    public String getCountMonth() {
        return countMonth;
    }

    public void setCountMonth(String countMonth) {
        this.countMonth = countMonth;
    }

    public Double getYearTarget() {
        return yearTarget;
    }

    public void setYearTarget(Double yearTarget) {
        this.yearTarget = yearTarget;
    }

    public Double getExecVal() {
        return execVal;
    }

    public void setExecVal(Double execVal) {
        this.execVal = execVal;
    }

    public Double getExecSumVal() {
        return execSumVal;
    }

    public void setExecSumVal(Double execSumVal) {
        this.execSumVal = execSumVal;
    }

    public Double getExecSumRate() {
        return execSumRate;
    }

    public void setExecSumRate(Double execSumRate) {
        this.execSumRate = execSumRate;
    }
}
