package com.poi.testpoi.pojo;

import javax.persistence.*;

/**
 * 部门收入、到款、新签、利润完成及成本控制情况
 */
@Entity
@Table(name = "bi_domain_complete")
public class BiDomainComplete {

    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    @Column(name = "id")
    private Integer id;

    //所属年份
    @Column(name = "count_year", nullable = true, length = 45)
    private String countYear;

    //经济指标
    @Column(name = "type_name", nullable = true, length = 45)
    private String typeName;

    //金额单位
    @Column(name = "value_unit", nullable = true, length = 45)
    private String valueUnit;

    //年度目标值
    @Column(name = "year_target")
    private Double yearTarget;

    //本月完成值
    @Column(name = "complete_val")
    private Double completeVal;

    //本月累计完成值
    @Column(name = "complete_sum_val")
    private Double completeSumVal;

    //完成率
    @Column(name = "complete_rate")
    private Double completeRate;

    //去年同期
    @Column(name = "last_year_val")
    private Double lastYearVal;

    //同比增减
    @Column(name = "last_year_increase_rate")
    private Double lastYearIncreaseRate;


    //====
    //开始月份
    @Column(name = "start_month", nullable = true, length = 45)
    private String startMonth;

    //结束月份
    @Column(name = "end_month", nullable = true, length = 45)
    private String endMonth;

    //领域名称
    @Column(name = "domain_name", nullable = true, length = 45)
    private String deptName;


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


    public Double getYearTarget() {
        return yearTarget;
    }

    public void setYearTarget(Double yearTarget) {
        this.yearTarget = yearTarget;
    }

    public Double getCompleteVal() {
        return completeVal;
    }

    public void setCompleteVal(Double completeVal) {
        this.completeVal = completeVal;
    }

    public Double getCompleteSumVal() {
        return completeSumVal;
    }

    public void setCompleteSumVal(Double completeSumVal) {
        this.completeSumVal = completeSumVal;
    }

    public Double getCompleteRate() {
        return completeRate;
    }

    public void setCompleteRate(Double completeRate) {
        this.completeRate = completeRate;
    }

    public Double getLastYearVal() {
        return lastYearVal;
    }

    public void setLastYearVal(Double lastYearVal) {
        this.lastYearVal = lastYearVal;
    }

    public Double getLastYearIncreaseRate() {
        return lastYearIncreaseRate;
    }

    public void setLastYearIncreaseRate(Double lastYearIncreaseRate) {
        this.lastYearIncreaseRate = lastYearIncreaseRate;
    }


    public String getStartMonth() {
        return startMonth;
    }

    public void setStartMonth(String startMonth) {
        this.startMonth = startMonth;
    }

    public String getEndMonth() {
        return endMonth;
    }

    public void setEndMonth(String endMonth) {
        this.endMonth = endMonth;
    }

    public String getDeptName() {
        return deptName;
    }

    public void setDeptName(String deptName) {
        this.deptName = deptName;
    }

}
