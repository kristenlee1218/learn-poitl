package com.learn.example6;

import com.deepoove.poi.data.NumbericRenderData;
/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
public class ExperienceData {
    private String company;
    private String department;
    private String time;
    private String position;
    private NumbericRenderData responsibility;

    public String getCompany() {
        return company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getTime() {
        return time;
    }

    public void setTime(String time) {
        this.time = time;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public NumbericRenderData getResponsibility() {
        return responsibility;
    }

    public void setResponsibility(NumbericRenderData responsibility) {
        this.responsibility = responsibility;
    }

}
