package com.alibaba.excel.analysis.v07;

import com.alibaba.excel.constant.ExcelXmlConstants;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventRegisterCenter;
import com.alibaba.excel.event.OneRowAnalysisFinishEvent;
import com.alibaba.excel.util.DateUtils;
import com.alibaba.excel.util.PositionUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Arrays;
import java.util.Date;

import static com.alibaba.excel.constant.ExcelXmlConstants.*;

public class KayWallXlsxRowHandler extends DefaultHandler {
    /**
     * 共享字符串表
     */
    private SharedStringsTable sst;
    /**
     * 上一次的内容
     */
    private String lastContents;
    /**
     * 字符串标识
     */
    private boolean nextIsString;
    /**
     * 行记录数组
     */
	private String[] curRowContent = new String[20];
    /**
     * 当前行号
     */
    private int curRow = 0;
    /**
     * 当前列
     */
    private int curCol = 0;
    /**
     * T元素标识
     */
    private boolean isTElement;
    /**
     * 单元格数据类型，默认为字符串类型
     */
    private CellDataType nextDataType = CellDataType.SSTINDEX;

    private final DataFormatter formatter = new DataFormatter();

    private short formatIndex;

    private String formatString;

	/**
	 * 定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
	 */
	private String preRef = null, ref = null;

	/**
	 * 定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
	 */
	private String maxRef = null;

    /**
     * 单元格
     */
    private StylesTable stylesTable;

	private AnalysisContext analysisContext;

	private AnalysisEventRegisterCenter registerCenter;

	private String currentCellIndex;

	private String currentCellValue;

	public KayWallXlsxRowHandler(AnalysisEventRegisterCenter registerCenter, SharedStringsTable sst,
			AnalysisContext analysisContext, StylesTable stylesTable) {
		this.registerCenter = registerCenter;
		this.sst = sst;
		this.analysisContext = analysisContext;
		this.stylesTable = stylesTable;
	}

	/**
     * 单元格中的数据可能的数据类型
     */
    enum CellDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
		setTotalRowCount(qName, attributes);

		startCell(qName, attributes);

		// 置空
		startCellValue(qName);
    }

	/**
	 * 记录行总数
	 * @param name
	 * @param attributes
	 */
	private void setTotalRowCount(String name, Attributes attributes) {
		if (DIMENSION.equals(name)) {
			String d = attributes.getValue(DIMENSION_REF);
			String totalStr = d.substring(d.indexOf(":") + 1, d.length());
			String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
			analysisContext.setTotalCount(Integer.parseInt(c));
		}
	}

	/**
	 * 读单元格组件前，清空 currentCellValue
	 * @param name
	 */
	private void startCellValue(String name) {
		if (name.equals(CELL_VALUE_TAG) || name.equals(CELL_VALUE_TAG_1)) {
			// initialize current cell value
			currentCellValue = "";
		}
	}

	/**
	 * 获取单元格数据类型（读取组件之前处理）
	 * @param qName
	 * @param attributes
	 */
	private void startCell(String qName, Attributes attributes) {
		// c => 单元格
		if ("c".equals(qName)) {
			// 位置索引
			currentCellIndex = attributes.getValue(ExcelXmlConstants.POSITION);
			int nextRow = PositionUtils.getRow(currentCellIndex);
			if (nextRow > curRow) {
				curRow = nextRow;
			}
			analysisContext.setCurrentRowNum(curRow);
			curCol = PositionUtils.getCol(currentCellIndex);

			// 设定单元格类型
			this.setNextDataType(attributes);

			// Figure out if the value is an index in the SST
			String cellType = attributes.getValue("t");
			nextDataType = CellDataType.SSTINDEX;
		}
	}

	/**
     * 单元格数据类型获取：nextDataType formatIndex formatString
     * @param attributes
     */
    public void setNextDataType(Attributes attributes) {
        nextDataType = CellDataType.NUMBER;
        formatIndex = -1;
        formatString = null;
        String cellType = attributes.getValue("t");
        String cellStyleStr = attributes.getValue("s");

        if ("b".equals(cellType)) {
            nextDataType = CellDataType.BOOL;
        } else if ("e".equals(cellType)) {
            nextDataType = CellDataType.ERROR;
        } else if ("inlineStr".equals(cellType)) {
            nextDataType = CellDataType.INLINESTR;
        } else if ("s".equals(cellType)) {
            nextDataType = CellDataType.SSTINDEX;
        } else if ("str".equals(cellType)) {
            nextDataType = CellDataType.FORMULA;
        }

        if (cellStyleStr != null) {
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();

			if ("m/d/yy" == formatString) {
				nextDataType = CellDataType.DATE;
				formatString = "yyyy-MM-dd hh:mm:ss.SSS";
			}

			if (formatString == null) {
				nextDataType = CellDataType.NULL;
				formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
			}
        }

		System.out.println("单元格信息" + currentCellIndex + ": nextDataType = " + nextDataType + ";formatIndex=" + formatIndex
				+ ";formatString = " + formatString);
    }


    /**
     * 根据单元格数据类型解析单元格数据（时间，日期处理）
     * @param value 单元格的值（这时候是一串数字）
     * @param thisStr 一个空字符串
     * @return
     */
    public String getDataValue(String value, String thisStr) {
        switch (nextDataType) {
            // 这几个的顺序不能随便交换，交换了很可能会导致数据错误
            case BOOL:
                char first = value.charAt(0);
                thisStr = first == '0' ? "FALSE" : "TRUE";
                break;
            case ERROR:
                thisStr = "\"ERROR:" + value + '"';
                break;
            case FORMULA:
                thisStr = '"' + value + '"';
                break;
            case INLINESTR:
                XSSFRichTextString rtsi = new XSSFRichTextString(value);
                thisStr = rtsi.toString();
                rtsi = null;
                break;
            case SSTINDEX:
                try {
					int idx = Integer.parseInt(value);
                    XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));
                    thisStr = rtss.toString();
                    rtss = null;
                } catch (NumberFormatException ex) {
                    thisStr = value;
                }
                break;
            case NUMBER:
                System.out.println("formatString:" + formatString);
                if (formatString != null) {
					Double d = Double.parseDouble(value);
					if(d.compareTo(0d) == 0){
						thisStr = "00:00:00";
						break;
					}
					Date date = HSSFDateUtil.getJavaDate(d, false);
					// 判断时间序列号
					if(value.contains(".")){
						// 分两种情况 年月日 时分秒 和 时分秒
						if(d.compareTo(new Double(1)) > 0){
							thisStr = DateUtils.format(date, DateUtils.DATE_FORMAT_19);
						}else {
							thisStr = DateUtils.format(date, DateUtils.DATE_FORMAT_8_TIME);
						}
					}else {
						// 不包含小数点，只有年月日
						thisStr = DateUtils.format(date, DateUtils.DATE_FORMAT_10);
					}
                } else {
                    thisStr = value;
                }
//                thisStr = thisStr.replace("_", "").trim();
                break;
            case DATE:
                thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                // 对日期字符串作特殊处理
                thisStr = thisStr.replace(" ", "T");
                break;
            default:
                thisStr = " ";
                break;
        }

        return thisStr;
    }


    /**
     * 检查 formatString 格式是否为时间/日期
     * @return
     */
    boolean checkFormatStringIfDate(){
        if(formatString.contains("h")){
            return true;
        }
        return false;
    }



    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
		endRow(qName);
		endCellValue(qName);

	}

	private void endCellValue(String qName) {
		// ensure size
		if (curCol >= curRowContent.length) {
			curRowContent = Arrays.copyOf(curRowContent, (int)(curCol * 1.5));
		}

		if ("v".equals(qName)) {
			currentCellValue = getDataValue(currentCellValue, "");
			curRowContent[curCol] = currentCellValue;
			curCol++;
		}else if ("t".equals(qName)) {
			curRowContent[curCol] = currentCellValue;
		}
	}


	@Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        // 得到单元格内容的值
		currentCellValue += new String(ch, start, length);
    }

	private void endRow(String name) {
		if (name.equals(ROW_TAG)) {
			registerCenter.notifyListeners(new OneRowAnalysisFinishEvent(curRowContent,curCol));
			curRowContent = new String[20];
		}
	}

}
