package com.aliatry.easyexcelencapsulation.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel监听器
 */
public class ExcelListener extends AnalysisEventListener {

    // 自定义用于暂时存储data
    // 可以通过实例获取该值
    private List<Object> data = new ArrayList<>();

    /**
     * 通过 AnalysisContext 对象还可以获取当前 sheet,当前行等数据
     */
    @Override
    public void invoke(Object object, AnalysisContext context) {
        // 数据存储到list,供批量处理,或后续自己业务逻辑处理
        data.add(object);
        // 根据自己业务做处理
        doSomething(object);
    }

    private void doSomething(Object object) {
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 解析结束销毁不用的资源
//        data.clear();
    }

    public List<Object> getData() {
        return data;
    }

    public void setData(List<Object> data) {
        this.data = data;
    }
}