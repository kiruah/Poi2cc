package com.kiruah.poi2cc;

import java.time.format.DateTimeFormatter;

/**
 * デフォルト定数
 *
 * @author Kiruah
 */
public class Poi2ccConstants {

	/** 空文字列 */
	public static final String EMPTY_STRING = "";

	/** アルファベット文字数 */
	public static final int ALPHABET_NUMBER = 'Z' - 'A' + 1;

	public static final DateTimeFormatter DEFAULT_DATE_TIME_FORMAT = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");

}
