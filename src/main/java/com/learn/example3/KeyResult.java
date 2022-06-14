package com.learn.example3;

/**
 * @author ：Kristen
 * @date ：2022/6/14
 * @description :
 */
public class KeyResult {
    private String desc;
    private String progress;

    public KeyResult(String desc, String progress) {
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