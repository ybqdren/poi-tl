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

import com.deepoove.poi.data.style.RowStyle;

/**
 * RowRenderData Class 对于表格中行的封装
 * 拥有3个属性，可用属性2：
 * 		cells {@link List<CellRenderData>}
 * 		rowStyle {@link RowStyle}
 * 	两个额外方法：
 * 		addCells(CellRenderData cell)
 * 		obtainColSize()
 *
 */
public class RowRenderData implements RenderData {

    private static final long serialVersionUID = 1L;
    private List<CellRenderData> cells = new ArrayList<>();
    private RowStyle rowStyle;

    public List<CellRenderData> getCells() {
        return cells;
    }

    public void setCells(List<CellRenderData> cells) {
        this.cells = cells;
    }

    public RowStyle getRowStyle() {
        return rowStyle;
    }

    public void setRowStyle(RowStyle rowStyle) {
        this.rowStyle = rowStyle;
    }

	/**
	 * 传入一个CellRenderData对象{@link CellRenderData}
	 * @param cell
	 * @return
	 */
	public RowRenderData addCell(CellRenderData cell) {
		// 将传入的单个CellRenderData对象放入 List<CellRenderData> 对象中
        cells.add(cell);
        return this;
    }

	/**
	 * 获得List<CellRenderData> 类型cells对象中存放CellRenderData的数量
	 * @return
	 */
	public int obtainColSize() {
        if (null == cells || cells.isEmpty()) return 0;
        return cells.size();
    }

}
