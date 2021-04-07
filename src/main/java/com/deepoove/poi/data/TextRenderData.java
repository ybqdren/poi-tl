/*
 * Copyright 2014-2020 Sayi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.deepoove.poi.data;

import com.deepoove.poi.data.style.Style;

/**
 * Basic text template
 * 
 * @author Sayi
 *
 * TextRenderData对象：
 * 包含两个属性：
 * 		style {@link Style}
 * 		text  {@link String}
 * 四个构造方法：
 * 		TextRenderData
 * 		TextRenderData(String text)
 * 		TextRenderData(String color,String text)
 * 		TextRenderData(String text,String style)
 * 剩下为getter/setter方法和toString方法重写形式
 *
 */
public class TextRenderData implements RenderData {

    private static final long serialVersionUID = 1L;

    // 文字的样式
    protected Style style;

    /**
     * \n means line break
	 * 文本内容 以\n结束
     */
    protected String text;

    public TextRenderData() {
    }

    public TextRenderData(String text) {
        this.text = text;
    }

    // 构造方法 传入颜色和文本内容
	// 将颜色放入style的颜色属性中
	// 此处style使用到了bulder类似的构造，可以深入进去看看
    public TextRenderData(String color, String text) {
        this.style = Style.builder().buildColor(color).build();
        this.text = text;
    }

    public TextRenderData(String text, Style style) {
        this.style = style;
        this.text = text;
    }

    public Style getStyle() {
        return style;
    }

    public void setStyle(Style style) {
        this.style = style;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    @Override
    public String toString() {
        return text;
    }

}
