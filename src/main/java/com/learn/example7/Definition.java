package com.learn.example7;

import com.deepoove.poi.data.TextRenderData;

import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
public class Definition {
    private String name;
    List<ExampleProperty> properties;
    List<TextRenderData> codes;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<ExampleProperty> getProperties() {
        return properties;
    }

    public void setProperties(List<ExampleProperty> properties) {
        this.properties = properties;
    }

    public List<TextRenderData> getCodes() {
        return codes;
    }

    public void setCodes(List<TextRenderData> codes) {
        this.codes = codes;
    }
}