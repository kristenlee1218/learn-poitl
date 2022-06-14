package com.learn.example7;

import com.deepoove.poi.data.TextRenderData;

import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
public class ExampleProperty {
    private String name;
    private boolean required;
    private List<TextRenderData> schema;
    private String description;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public List<TextRenderData> getSchema() {
        return schema;
    }

    public void setSchema(List<TextRenderData> schema) {
        this.schema = schema;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }
}