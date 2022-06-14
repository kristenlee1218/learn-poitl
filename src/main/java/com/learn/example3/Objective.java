package com.learn.example3;

public class Objective {
    private String desc;
    private String progress;

    public Objective(String desc, String progress) {
        this.desc = desc;
        this.progress = progress;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public String getProgress() {
        return progress;
    }

    public void setProgress(String progress) {
        this.progress = progress;
    }

}