package com.ybqdren.github.policy;

import com.deepoove.poi.data.style.BorderStyle;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * @Author WenZhao <withzhaowen@126com>
 * @GiteHub https://github.com/ybqdren
 * @Date 2021/4/1 10:50
 * @Description
 **/

public class CustomTableRenderPolicy extends AbstractRenderPolicy<Object> {
	@Override
	public void doRender(RenderContext<Object> context) throws Exception {
		// 获取一个行游标
		XWPFRun run = context.getRun();

		// 获取一个BodyContainer
		BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);

		// 定义表格的行列
		int row = 10,col = 8;

		// 将run插入表格（bodyContainer对象）中，获取到一个XWPFTable表格对象
		// 通过bodyContainer.insertNewTable在当前标签位置插入表格
		XWPFTable table = bodyContainer.insertNewTable(run,row,col);

		// 设置表格的宽度 此方法已过时 需要使用setWidth方法进行代替
		TableTools.widthTable(table,15.63f,col);

		// 设置表格的边框和样式
		TableTools.borderTable(table, BorderStyle.DEFAULT);


		// 1) 调用XWPFTable API操作表格
		// 2) 调用TableRenderPolicy.Helper.renderRow方法快速方便的渲染一行数据
		// 3）调用TableTools类方法操作表格，比如合并单元格
		// 设置表格合并参数
		// 水平方向合并表格 从 第0列合并到第7列
		TableTools.mergeCellsHorizonal(table,0,0,7);
		// 从垂直方向合并表格 从 第0行合并到第9行
		TableTools.mergeCellsVertically(table,0,0,9);
	}
}
