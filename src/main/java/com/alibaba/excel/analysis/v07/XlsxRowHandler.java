package com.alibaba.excel.analysis.v07;

import com.alibaba.excel.annotation.FieldType;
import com.alibaba.excel.constant.ExcelXmlConstants;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventRegisterCenter;
import com.alibaba.excel.event.OneRowAnalysisFinishEvent;
import com.alibaba.excel.util.DateUtils;
import com.alibaba.excel.util.PositionUtils;
import com.alibaba.excel.util.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

import static com.alibaba.excel.constant.ExcelXmlConstants.*;

/**
 *
 * @author jipengfei
 */
public class XlsxRowHandler extends DefaultHandler {

	private String currentCellIndex;

	private FieldType currentCellType;

	private int curRow;

	private int curCol;

	private String[] curRowContent = new String[20];

	private String currentCellValue;

	private SharedStringsTable sst;

	private AnalysisContext analysisContext;

	private AnalysisEventRegisterCenter registerCenter;

	/**
	 * 单元格格式
	 */
	private StylesTable stylesTable;
	private short formatIndex;
	private String formatString;

	/**
	 * 日期类型
	 */
	List<Integer> dateTypeList = Arrays
			.asList(14, 15, 16, 17, 22, 30, 31, 57, 58, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189,
					190, 191, 192, 193, 194, 195, 196, 197, 198, 199);

	/**
	 * 时间类型
	 */
	List<Integer> timeTypeList = Arrays
			.asList(18, 19, 20, 21, 32, 33, 45, 46, 47, 55, 56, 176, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186);

	/**
	 * 科学计数正则表达式
	 */
	private static Pattern pattern = Pattern.compile("[+-]?[\\d]+([.][\\d]*)?([Ee][+-]?[\\d]+)?");
	private static Pattern formatStringPattern = Pattern.compile("0\\.[0]*E\\+00");

	public XlsxRowHandler(AnalysisEventRegisterCenter registerCenter, SharedStringsTable sst,
			AnalysisContext analysisContext) {
		this.registerCenter = registerCenter;
		this.analysisContext = analysisContext;
		this.sst = sst;

	}

	public XlsxRowHandler(AnalysisEventRegisterCenter registerCenter, SharedStringsTable sst,
			AnalysisContext analysisContext, StylesTable stylesTable) {
		this.sst = sst;
		this.analysisContext = analysisContext;
		this.registerCenter = registerCenter;
		this.stylesTable = stylesTable;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

		setTotalRowCount(name, attributes);

		startCell(name, attributes);

		startCellValue(name);

	}

	private void startCellValue(String name) {
		if (name.equals(CELL_VALUE_TAG) || name.equals(CELL_VALUE_TAG_1)) {
			// initialize current cell value
			currentCellValue = "";
		}
	}

	private void startCell(String name, Attributes attributes) {
		if (ExcelXmlConstants.CELL_TAG.equals(name)) {
			// 位置索引
			currentCellIndex = attributes.getValue(ExcelXmlConstants.POSITION);
			int nextRow = PositionUtils.getRow(currentCellIndex);
			if (nextRow > curRow) {
				curRow = nextRow;
			}
			analysisContext.setCurrentRowNum(curRow);
			curCol = PositionUtils.getCol(currentCellIndex);

			// 获取单元格数据类型  currentCellType
			/**
			 *  s : 字符串
			 *  null : 数字，日期
			 */
			currentCellType = FieldType.EMPTY;
			formatIndex = -1;
			formatString = null;
			String cellType = attributes.getValue("t");
			// 获取单元格样式
			String cellStyle = attributes.getValue("s");
			if(cellType == null){
				if(cellStyle != null){
					int styleIndex = Integer.parseInt(cellStyle);
					XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
					formatIndex = style.getDataFormat();
					formatString = style.getDataFormatString();
					// 判断日期
//					currentCellType = getFieldTypeByFormatIndex();
					if(checkFormatIndexIfDate()){
						currentCellType = FieldType.DATE;
					}

					// 科学计数类型
					if(checkENum()){
						currentCellType = FieldType.ENumber;
					}
				}
			}

			// 字符串
			if (cellType != null && cellType.equals("s")) {
				currentCellType = FieldType.STRING;
			}

//			System.out.println("单元格信息" + currentCellIndex + ": cellType=" + cellType + "; cellStyle=" + cellStyle
//					+ "; currentCellType= " + currentCellType + "; formatIndex=" + formatIndex + "; formatString= "
//					+ formatString);
		}
	}

	private FieldType getFieldTypeByFormatIndex(){
		Integer format = Short.toUnsignedInt(formatIndex);
		if(dateTypeList.contains(format)){
			return  FieldType.DATE;
		}else if(timeTypeList.contains(format)){
			return  FieldType.TIME;
		}else {
			return  FieldType.EMPTY;
		}
	}

	private boolean checkFormatIndexIfDate(){
		Integer format = Short.toUnsignedInt(formatIndex);
		if(dateTypeList.contains(format) || timeTypeList.contains(format)){
			if(format.equals(176) || format.equals(177) || format.equals(178) || format.equals(179) || format.equals(180)){
				if(formatString.contains("mm") || formatString.contains("m") || formatString.equals("yy/m/d") || formatString.equals("m/d")){
					return true;
				}
			}else {
				return true;
			}
		}
		return false;
	}

	private void endCellValue(String name) throws SAXException {
		// ensure size
		if (curCol >= curRowContent.length) {
			curRowContent = Arrays.copyOf(curRowContent, (int)(curCol * 1.5));
		}
		if (CELL_VALUE_TAG.equals(name)) {
			getCellDataValue();
			curRowContent[curCol] = currentCellValue;
		} else if (CELL_VALUE_TAG_1.equals(name)) {
			curRowContent[curCol] = currentCellValue;
		}
	}

	private void getCellDataValue() {
		switch (currentCellType) {
			case STRING:
				int idx = Integer.parseInt(currentCellValue);
				currentCellValue = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				currentCellType = FieldType.EMPTY;
				break;
			case DATE:
				// 转化时间
				Double d = Double.parseDouble(currentCellValue);
				if(d.compareTo(0d) == 0){
					currentCellValue = "00:00:00";
					break;
				}
				Date date = HSSFDateUtil.getJavaDate(d, false);
				// 判断时间序列号
				if(currentCellValue.contains(".")){
					// 分两种情况 年月日 时分秒  /  时分秒
					if(d.compareTo(new Double(1)) > 0){
						currentCellValue = DateUtils.format(date, DateUtils.DATE_FORMAT_19);
					}else {
						currentCellValue = DateUtils.format(date, DateUtils.DATE_FORMAT_8_TIME);
					}
				}else {
					// 不包含小数点，只有年月日
					currentCellValue = DateUtils.format(date, DateUtils.DATE_FORMAT_10);
				}
				break;
			case ENumber:
				// 科学计数
				if(!StringUtils.isEmpty(currentCellValue)){
					BigDecimal decimal = new BigDecimal(currentCellValue);
					currentCellValue = decimal.stripTrailingZeros().toPlainString();
				}
				break;
			default:
				break;
		}
	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		endRow(name);
		endCellValue(name);
	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		currentCellValue += new String(ch, start, length);
	}


	private void setTotalRowCount(String name, Attributes attributes) {
		if (DIMENSION.equals(name)) {
			String d = attributes.getValue(DIMENSION_REF);
			String totalStr = d.substring(d.indexOf(":") + 1, d.length());
			String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
			analysisContext.setTotalCount(Integer.parseInt(c));
		}

	}

	private void endRow(String name) {
		if (name.equals(ROW_TAG)) {
			registerCenter.notifyListeners(new OneRowAnalysisFinishEvent(curRowContent,curCol));
			curRowContent = new String[20];
		}
	}


	/**
	 * 判断输入字符串是否为科学计数法
	 * @param input
	 * @return
	 */
	private static boolean isENum(String input) {
		return pattern.matcher(input).matches();
	}

	/**
	 * 通过 formatString 判断单元格数据是否为科学计数法
	 * @return
	 */
	private boolean checkENum(){
		if(!StringUtils.isEmpty(formatString)){
			return formatStringPattern.matcher(formatString).matches();
		}
		return false;
	}
}

