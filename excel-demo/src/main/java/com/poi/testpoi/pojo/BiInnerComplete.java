package com.poi.testpoi.pojo;

import javax.persistence.*;

/**
 * 内部指标总体完成情况（含公司）
 */
@Entity
@Table(name = "bi_inner_complete")
public class BiInnerComplete {

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

    //月份
    @Column(name = "count_month", nullable = true, length = 45)
    private String countMonth;

    //年度目标值
    @Column(name = "inner_target")
    private Double innerTarget;

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

    public String getCountMonth() {
        return countMonth;
    }

    public void setCountMonth(String countMonth) {
        this.countMonth = countMonth;
    }

    public Double getInnerTarget() {
        return innerTarget;
    }

    public void setInnerTarget(Double innerTarget) {
        this.innerTarget = innerTarget;
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
}
