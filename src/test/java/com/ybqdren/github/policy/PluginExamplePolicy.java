package com.ybqdren.github.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.data.Numberings;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.render.WhereDelegate;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

/**
 * @Author ZhaoWen <withzhaowen@126com>
 * @GitHub https://github.com/ybqdren
 * @Date 2021/4/1 11:19
 * @Description
 **/

// 使用匿名类创建新插件 Do Anything Anywhere

public class PluginExamplePolicy {
	public static void main(String[] args) throws IOException {
// where绑定policy
		Configure config = Configure.builder().bind("sea", new AbstractRenderPolicy<String>() {
			@Override
			public void doRender(RenderContext<String> context) throws Exception {
				// anywhere
				XWPFRun where = context.getWhere();
				// anything
				String thing = context.getThing();
				// do 文本
				where.setText(thing, 0);
			}
		}).bind("sea_img", new AbstractRenderPolicy<String>() {
			@Override
			public void doRender(RenderContext<String> context) throws Exception {
				// anywhere delegate
				WhereDelegate where = context.getWhereDelegate();
				// any thing
				String thing = context.getThing();
				// do 图片
				FileInputStream stream = null;
				try {
					stream = new FileInputStream(thing);
					where.addPicture(stream, XWPFDocument.PICTURE_TYPE_JPEG, 400, 450);
				} finally {
					IOUtils.closeQuietly(stream);
				}
				// clear
				clearPlaceholder(context, false);
			}
		}).bind("sea_feature", new AbstractRenderPolicy<List<String>>() {
			@Override
			public void doRender(RenderContext<List<String>> context) throws Exception {
				// anywhere delegate
				WhereDelegate where = context.getWhereDelegate();
				// anything
				List<String> thing = context.getThing();
				// do 列表
				where.renderNumbering(Numberings.of(thing.toArray(new String[] {})).create());
				// clear
				clearPlaceholder(context, true);
			}
		}).build();

		// 初始化where的数据
		HashMap<String, Object> arg = new HashMap<String, Object>();
		arg.put("sea", "Hello, world!");
//		arg.put("sea_img", "sea.jpg");
		arg.put("sea_feature", Arrays.asList("面朝大海春暖花开", "今朝有酒今朝醉"));
		arg.put("sea_location", Arrays.asList("日落：日落山花红四海", "花海：你想要的都在这里"));

		// 一行代码
		XWPFTemplate.compile("sea.docx", config).render(arg).writeToFile("out_sea.docx");
	}
}
