package com.alibaba.excel.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.regex.Pattern;

/**
 *  日期工具类
 * @author zhuangjiaju
 * @date 2019年05月06日 16:08
 */
public class DateUtils {
	public static final String DATE_FORMAT_6 = "HHmmss";
	public static final String DATE_FORMAT_8_TIME = "HH:mm:ss";
	public static final String DATE_FORMAT_8 = "yyyyMMdd";
	public static final String DATE_FORMAT_10 = "yyyy-MM-dd";
	public static final String DATE_FORMAT_10_FORWARD_SLASH = "yyyy/MM/dd";
	public static final String DATE_FORMAT_14 = "yyyyMMddHHmmss";
	public static final String DATE_FORMAT_19 = "yyyy-MM-dd HH:mm:ss";
	public static final String DATE_FORMAT_19_FORWARD_SLASH = "yyyy/MM/dd HH:mm:ss";

	/***
	 * 判断是否一个 double
	 */
	private static final Pattern DOUBLE_PATTERN = Pattern.compile("^-?([1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*|0?\\.0+|0)$");

	private static final Logger LOGGER = LoggerFactory.getLogger(DateUtils.class);

	private DateUtils() {
	}


	/**
	 * 将任何String类型的日期转换为日期，不抛出异常,如果遇到不兼容的格式 建议直接加进来
	 *
	 * <p>
	 * 目前支持如下格式：
	 * <li>yyyy-MM-dd HH:mm:ss
	 * <li>yyyy/MM/dd HH:mm:ss
	 * <li>yyyyMMddHHmmss
	 * <li>yyyy-MM-dd
	 * <li>yyyy/MM/dd
	 * <li>yyyyMMdd
	 *
	 * @param dateString       必须是24h的日期格式 如果不是 需要修改此方法
	 * @param use1904windowing
	 * @return 如果异常 则输出warn且返回null
	 * @author zhuangjiaju
	 */
	public static Date parseDate(String dateString, boolean use1904windowing) {
		return parseDate(dateString, null, use1904windowing);
	}

	/**
	 * 将一个String转换根据指定格式转换成日期
	 *
	 * @param dateString
	 * @param dateFormat
	 * @param use1904windowing
	 * @return 不会抛出异常，但是如果转换错误，返回对象可能为null 请自行判断
	 * @author zhuangjiaju
	 */
	public static Date parseDate(String dateString, String dateFormat, boolean use1904windowing) {
		if (StringUtils.isEmpty(dateString)) {
			return null;
		}
		String dateStringTrim = dateString.trim();
		if (StringUtils.isEmpty(dateFormat)) {
			if (isValidExcelDate(dateStringTrim)) {
				return HSSFDateUtil.getJavaDate(Double.parseDouble(dateStringTrim), use1904windowing);
			}
			dateFormat = switchDateFormat(dateStringTrim);
		}
		if (StringUtils.isEmpty(dateFormat)) {
			return null;
		}
		try {
			return new SimpleDateFormat(dateFormat).parse(dateString);
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期字符串为:{}，转换格式为:{}", dateString, dateFormat);
		}
		return null;
	}


	/**
	 * 将任何String类型的日期转换为日期，不抛出异常,如果遇到不兼容的格式 建议直接加进来
	 *
	 * <p>
	 * 目前支持如下格式：
	 * <li>yyyy-MM-dd
	 * <li>yyyy/MM/dd
	 * <li>yyyyMMdd
	 *
	 * @param dateString       必须是24h的日期格式 如果不是 需要修改此方法
	 * @param use1904windowing
	 * @return 如果异常 则输出warn且返回null
	 * @author zhuangjiaju
	 */
	public static LocalDate parseLocalDate(String dateString, boolean use1904windowing) {
		return parseLocalDate(dateString, null, use1904windowing);
	}

	/**
	 * 将一个String转换根据指定格式转换成日期
	 *
	 * @param dateString
	 * @param dateFormat
	 * @param use1904windowing
	 * @return 不会抛出异常，但是如果转换错误，返回对象可能为null 请自行判断
	 * @author zhuangjiaju
	 */
	public static LocalDate parseLocalDate(String dateString, String dateFormat, boolean use1904windowing) {
		if (StringUtils.isEmpty(dateString)) {
			return null;
		}
		String dateStringTrim = dateString.trim();
		if (StringUtils.isEmpty(dateFormat)) {
			if (isValidExcelDate(dateStringTrim)) {
				return HSSFDateUtil.getJavaDate(Double.parseDouble(dateStringTrim), use1904windowing).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
			}
			dateFormat = switchDateFormat(dateStringTrim);
		}
		if (StringUtils.isEmpty(dateFormat)) {
			return null;
		}
		try {
			return LocalDate.parse(dateString, DateTimeFormatter.ofPattern(dateFormat));
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期字符串为:{}，转换格式为:{}", dateString, dateFormat);
		}
		return null;
	}

	/**
	 * 将任何String类型的日期转换为日期，不抛出异常,如果遇到不兼容的格式 建议直接加进来
	 *
	 * <p>
	 * 目前支持如下格式：
	 * <li>HHmmss
	 * <li>HH:mm:ss
	 *
	 * @param dateString       必须是24h的日期格式 如果不是 需要修改此方法
	 * @param use1904windowing
	 * @return 如果异常 则输出warn且返回null
	 * @author zhuangjiaju
	 */
	public static LocalTime parseLocalTime(String dateString, boolean use1904windowing) {
		return parseLocalTime(dateString, null, use1904windowing);
	}

	/**
	 * 将一个String转换根据指定格式转换成日期
	 *
	 * @param dateString
	 * @param dateFormat
	 * @param use1904windowing
	 * @return 不会抛出异常，但是如果转换错误，返回对象可能为null 请自行判断
	 * @author zhuangjiaju
	 */
	public static LocalTime parseLocalTime(String dateString, String dateFormat, boolean use1904windowing) {
		if (StringUtils.isEmpty(dateString)) {
			return null;
		}
		String dateStringTrim = dateString.trim();
		if (StringUtils.isEmpty(dateFormat)) {
			if (isValidExcelDate(dateStringTrim)) {
				return HSSFDateUtil.getJavaDate(Double.parseDouble(dateStringTrim), use1904windowing).toInstant().atZone(ZoneId.systemDefault()).toLocalTime();
			}
			dateFormat = switchDateFormat(dateStringTrim);
		}
		if (StringUtils.isEmpty(dateFormat)) {
			return null;
		}
		try {
			return LocalTime.parse(dateString, DateTimeFormatter.ofPattern(dateFormat));
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期字符串为:{}，转换格式为:{}", dateString, dateFormat);
		}
		return null;
	}


	/**
	 * 将任何String类型的日期转换为日期，不抛出异常,如果遇到不兼容的格式 建议直接加进来
	 *
	 * <p>
	 * 目前支持如下格式：
	 * <li>yyyy-MM-dd HH:mm:ss
	 * <li>yyyy/MM/dd HH:mm:ss
	 * <li>yyyyMMddHHmmss
	 * <li>yyyy-MM-dd
	 * <li>yyyy/MM/dd
	 * <li>yyyyMMdd
	 *
	 * @param dateString       必须是24h的日期格式 如果不是 需要修改此方法
	 * @param use1904windowing
	 * @return 如果异常 则输出warn且返回null
	 * @author zhuangjiaju
	 */
	public static LocalDateTime parseLocalDateTime(String dateString, boolean use1904windowing) {
		return parseLocalDateTime(dateString, null, use1904windowing);
	}

	/**
	 * 将一个String转换根据指定格式转换成日期
	 *
	 * @param dateString
	 * @param dateFormat
	 * @param use1904windowing
	 * @return 不会抛出异常，但是如果转换错误，返回对象可能为null 请自行判断
	 * @author zhuangjiaju
	 */
	public static LocalDateTime parseLocalDateTime(String dateString, String dateFormat, boolean use1904windowing) {
		if (StringUtils.isEmpty(dateString)) {
			return null;
		}
		String dateStringTrim = dateString.trim();
		if (StringUtils.isEmpty(dateFormat)) {
			if (isValidExcelDate(dateStringTrim)) {
				return HSSFDateUtil.getJavaDate(Double.parseDouble(dateStringTrim), use1904windowing).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
			}
			dateFormat = switchDateFormat(dateStringTrim);
		}
		if (StringUtils.isEmpty(dateFormat)) {
			return null;
		}
		try {
			return LocalDateTime.parse(dateString, DateTimeFormatter.ofPattern(dateFormat));
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期字符串为:{}，转换格式为:{}", dateString, dateFormat);
		}
		return null;
	}


	/**
	 * 选择日期转换器
	 *
	 * @param dateStringTrim
	 * @return
	 */
	private static String switchDateFormat(String dateStringTrim) {
		int length = dateStringTrim.length();
		switch (length) {
			case 19:
				if (dateStringTrim.contains("-")) {
					return DATE_FORMAT_19;
				} else {
					return DATE_FORMAT_19_FORWARD_SLASH;
				}
			case 14:
				return DATE_FORMAT_14;
			case 10:
				if (dateStringTrim.contains("-")) {
					return DATE_FORMAT_10;
				} else {
					return DATE_FORMAT_10_FORWARD_SLASH;
				}
			case 8:
				if (dateStringTrim.contains(":")) {
					return DATE_FORMAT_8;
				} else {
					return DATE_FORMAT_8_TIME;
				}
			case 6:
				return DATE_FORMAT_6;
			default:
				return null;
		}
	}

	/**
	 * 判断是否是excel的日期
	 *
	 * @param dateStringTrim
	 * @return
	 */
	private static boolean isValidExcelDate(String dateStringTrim) {
		if (!DOUBLE_PATTERN.matcher(dateStringTrim).matches()) {
			return false;
		}
		return DateUtil.isValidExcelDate(Double.parseDouble(dateStringTrim));
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param date
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(Date date) {
		return format(date, null);
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param date
	 * @param dateFormat
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(Date date, String dateFormat) {
		if (date == null) {
			return "";
		}
		if (StringUtils.isEmpty(dateFormat)) {
			dateFormat = DATE_FORMAT_19;
		}
		try {
			return new SimpleDateFormat(dateFormat).format(date);
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期为:{}，转换格式为:{}", date, dateFormat);
		}
		return "";
	}


	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param localDate
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(LocalDate localDate) {
		return format(localDate, null);
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param localDate
	 * @param dateFormat
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(LocalDate localDate, String dateFormat) {
		if (localDate == null) {
			return "";
		}
		if (StringUtils.isEmpty(dateFormat)) {
			dateFormat = DATE_FORMAT_10;
		}
		try {
			return localDate.format(DateTimeFormatter.ofPattern(dateFormat));
		} catch (Exception e) {
			e.printStackTrace();
			LOGGER.warn("日期格式转换异常，日期为:{}，转换格式为:{}", localDate, dateFormat);
		}
		return "";
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param localTime
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(LocalTime localTime) {
		return format(localTime, null);
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param localTime
	 * @param dateFormat
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(LocalTime localTime, String dateFormat) {
		if (localTime == null) {
			return "";
		}
		if (StringUtils.isEmpty(dateFormat)) {
			dateFormat = DATE_FORMAT_8_TIME;
		}
		try {
			return localTime.format(DateTimeFormatter.ofPattern(dateFormat));
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期为:{}，转换格式为:{}", localTime, dateFormat);
		}
		return "";
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param localDateTime
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(LocalDateTime localDateTime) {
		return format(localDateTime, null);
	}

	/**
	 * 根据日期返回对应的日期格式
	 *
	 * @param localDateTime
	 * @param dateFormat
	 * @return 如果遇到空 则返回空字符串
	 * @author zhuangjiaju
	 */
	public static String format(LocalDateTime localDateTime, String dateFormat) {
		if (localDateTime == null) {
			return "";
		}
		if (StringUtils.isEmpty(dateFormat)) {
			dateFormat = DATE_FORMAT_19;
		}
		try {
			return localDateTime.format(DateTimeFormatter.ofPattern(dateFormat));
		} catch (Exception e) {
			LOGGER.warn("日期格式转换异常，日期为:{}，转换格式为:{}", localDateTime, dateFormat);
		}
		return "";
	}
}
