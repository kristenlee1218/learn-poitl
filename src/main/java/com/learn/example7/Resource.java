package com.learn.example7;

import java.util.List;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
public class Resource {
    private String name;
    private String description;
    private List<Endpoint> endpoints;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public List<Endpoint> getEndpoints() {
        return endpoints;
    }

    public void setEndpoints(List<Endpoint> endpoints) {
        this.endpoints = endpoints;
    }
}