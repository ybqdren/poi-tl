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

package com.deepoove.poi.xwpf;

import java.util.List;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import com.deepoove.poi.util.ParagraphUtils;
import com.deepoove.poi.util.ReflectionUtils;

/**
 * {@link IBody} operation
 * 
 * @author Sayi
 *
 */
public interface BodyContainer extends ParentContext {

    /**
     * get the position of paragraph in bodyElements
     * 
     * @param ctp paragraph
     * @return the position of paragraph
     */

	/**
	 * 获取CTP在段落中的下标位置
	 *
	 * 传入参数为一个 org.openxmlformats.schemas.wordprocessingml.x2006.main.ctp接口实现类
	 * CTP是一POI对于操作ooxml规范下xml文件标签属性封装的接口
	 *
	 * @param ctp
	 * @return
	 */
	default int getPosOfParagraphCTP(CTP ctp) {
		// org.apache.poi.xwpf.usermodel.IBodyElement -> poi中对于ooxml规范下xml文件中所有标签的封装类
		// 通过 IBody getBody() 方法可以获取到文档xml化后所有的标签信息
        IBodyElement current;

        // getTarget().getBodyElements 是实现IBody接口的一个方法 作用是按照文本的顺序返回段落表格
		// 返回对象为 IBodyElement类型的List集合 也就是此处的bodyElements
        List<IBodyElement> bodyElements = getTarget().getBodyElements();

		// 对于获取到的List<IBodyElement> 对象进行遍历
        for (int i = 0; i < bodyElements.size(); i++) {
        	// 将i个IBodyElement对象从bodyElements中获取出来
            current = bodyElements.get(i);

			// 将刚才获取到的IBodyElements对象的类型 进行判断
			/**
			 * 在BodyElementType中内置了三种类型
			 *     CONTENTCONTROL （翻译软件为内容控制 具体不可查 需要后续再观察）
			 *     PARAGRAPH（段落）
			 *     TABLE（表格）
			 */
            if (current.getElementType() == BodyElementType.PARAGRAPH) {
            	// 如果是段落 继续进if判断
				// 判断当前段落中的CTP信息 返回下标i
                if (((XWPFParagraph) current).getCTP().equals(ctp)) {
                    return i;
                }
            }
        }
        return -1;
    }

    /**
     * get the position of paragraph in bodyElements
     * 
     * @param paragraph
     * @return the position of paragraph
     */
    default int getPosOfParagraph(XWPFParagraph paragraph) {
        return getPosOfParagraphCTP(paragraph.getCTP());
    }

    /**
     * get all bodyElements
     * 
     * @return
     */
    @SuppressWarnings("unchecked")
    default List<IBodyElement> getBodyElements() {
        return (List<IBodyElement>) ReflectionUtils.getValue("bodyElements", getTarget());
    }

    /**
     * remove body element from bodyElements
     * 
     * @param pos the position of bodyElement
     */
    void removeBodyElement(int pos);

    /**
     * insert paragraph at position of the cursor
     * 
     * @param insertPostionCursor
     * @return the inserted paragraph
     */
    default XWPFParagraph insertNewParagraph(XmlCursor insertPostionCursor) {
        return getTarget().insertNewParagraph(insertPostionCursor);
    }

    /**
     * insert paragraph at position of run
     * 
     * @param run
     * @return the inserted paragraph
     */
    default XWPFParagraph insertNewParagraph(XWPFRun run) {
        XmlCursor cursor = ((XWPFParagraph) run.getParent()).getCTP().newCursor();
        return insertNewParagraph(cursor);
    }

    /**
     * get the position of paragraph in paragraphs
     * 
     * @param paragraph
     * @return the position of paragraph
     */
    default int getParaPos(XWPFParagraph paragraph) {
        List<XWPFParagraph> paragraphs = getTarget().getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            if (paragraphs.get(i) == paragraph) {
                return i;
            }
        }
        return -1;
    }

    /**
     * set paragraph at position
     * 
     * @param paragraph
     * @param pos
     */
    void setParagraph(XWPFParagraph paragraph, int pos);

    /**
     * container itself
     * 
     * @return
     */
    IBody getTarget();

    /**
     * insert table at position of the cursor
     * 
     * @param insertPostionCursor
     * @return the inserted table
     */
    default XWPFTable insertNewTbl(XmlCursor insertPostionCursor) {
        return getTarget().insertNewTbl(insertPostionCursor);
    }

    /**
     * get the position of table in tables
     * 
     * @param table
     * @return the position of table
     */
    default int getTablePos(XWPFTable table) {
        List<XWPFTable> tables = getTarget().getTables();
        for (int i = 0; i < tables.size(); i++) {
            if (tables.get(i) == table) {
                return i;
            }
        }
        return -1;
    }

    /**
     * set table
     * 
     * @param tablePos
     * @param table
     */
    void setTable(int tablePos, XWPFTable table);

    /**
     * update body elements
     * 
     * @param bodyElement
     * @param copy
     */
    default void updateBodyElements(IBodyElement bodyElement, IBodyElement copy) {
        int pos = -1;
        List<IBodyElement> bodyElements = getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            if (bodyElements.get(i) == bodyElement) {
                pos = i;
            }
        }
        if (-1 != pos) bodyElements.set(pos, copy);
    }

    /**
     * insert table at position of the run
     * 
     * @param run
     * @param row
     * @param col
     * @return
     */
    XWPFTable insertNewTable(XWPFRun run, int row, int col);

    /**
     * clear run
     * 
     * @param run
     */
    default void clearPlaceholder(XWPFRun run) {
        IRunBody parent = run.getParent();
        run.setText("", 0);
        if (parent instanceof XWPFParagraph) {
            String paragraphText = ParagraphUtils.trimLine((XWPFParagraph) parent);
            boolean havePictures = ParagraphUtils.havePictures((XWPFParagraph) parent);
            if ("".equals(paragraphText) && !havePictures) {
                int pos = getPosOfParagraph((XWPFParagraph) parent);
                removeBodyElement(pos);
            }
        }
    }

    XWPFSection closelySectPr(IBodyElement element);

}
