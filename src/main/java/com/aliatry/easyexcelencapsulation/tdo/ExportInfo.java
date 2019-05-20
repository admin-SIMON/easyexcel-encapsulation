package com.aliatry.easyexcelencapsulation.tdo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * 导出 Excel 时使用的映射实体类,Excel 模型
 */
public class ExportInfo extends BaseRowModel {
    @ExcelProperty(value = "NAME", index = 0)
    private String name;

    @ExcelProperty(value = "AGE", index = 1)
    private String age;

    @ExcelProperty(value = "EMAIL", index = 2)
    private String email;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }
}