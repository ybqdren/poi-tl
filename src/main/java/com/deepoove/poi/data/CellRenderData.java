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

import java.util.ArrayList;
import java.util.List;

import com.deepoove.poi.data.style.CellStyle;

/**
 * @author Sayi
 *
 * CellRenderData Class 一个与表格单元格有关的类
 * 两个属性：
 * 		paragraphs {@link List<ParagraphRenderData>}
 * 		cellStyle {@link CellStyle}
 *
 * 	一个额外方法：
 * 		addParagraph(ParagraphRenderData para)
 */
public class CellRenderData implements RenderData {

    private static final long serialVersionUID = 1L;
    private List<ParagraphRenderData> paragraphs = new ArrayList<>();
    private CellStyle cellStyle;

    public List<ParagraphRenderData> getParagraphs() {
        return paragraphs;
    }

    public void setParagraphs(List<ParagraphRenderData> paragraphs) {
        this.paragraphs = paragraphs;
    }

	/**
	 * 将{@link ParagraphRenderData} ParagraphRenderData 对象放入List<ParagraphRenderData>
	 *  在poi-tl中，构造表格最小单位是段落，也即ParagraphRenderData
	 * @param para
	 * @return
	 */
	public CellRenderData addParagraph(ParagraphRenderData para) {
        this.paragraphs.add(para);
        return this;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

}
