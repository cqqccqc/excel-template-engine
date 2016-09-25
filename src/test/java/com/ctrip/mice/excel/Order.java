package com.ctrip.mice.excel;

/**
 * Created by chenqi on 16/9/24.
 * For test orderTest.xlsx
 */
public class Order {

    private Double totalAmount = 100.21D;

    private String moreRequirement = "more requirement";

    private String remark = "remark remark";

    private Long orderNo = 20012345L;

    public Double getTotalAmount() {
        return totalAmount;
    }

    public void setTotalAmount(Double totalAmount) {
        this.totalAmount = totalAmount;
    }

    public String getMoreRequirement() {
        return moreRequirement;
    }

    public void setMoreRequirement(String moreRequirement) {
        this.moreRequirement = moreRequirement;
    }

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }

    public Long getOrderNo() {
        return orderNo;
    }

    public void setOrderNo(Long orderNo) {
        this.orderNo = orderNo;
    }
}
