package com.kiruah.poi2cc.storage;

import java.io.Serializable;

import com.kiruah.poi2cc.Poi2ccConstants;
import com.kiruah.poi2cc.Poi2ccRuntimeException;
import com.kiruah.poi2cc.Poi2ccUtil;

/**
 * セル参照先アドレス情報
 *
 * @author Kiruah
 */
public class Address implements Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = 6872916140404518127L;

	/** 列名(アルファベット表記) */
	protected String column = "A";

	/** 列番号 */
	protected int columnNumber = 0;

	/** 行番号 */
	protected int rowNumber = 0;

	/**
	 * デフォルトコンストラクタ
	 */
	public Address() {

	}

	/**
	 * コンストラクタ
	 *
	 * @param position 位置
	 * @param row trueの場合行を示し、falseの場合列を示す
	 */
	public Address(int position, boolean row) {

		if (row == true) {
			rowNumber = position;
			columnNumber = -1;
		} else {
			rowNumber = -1;
			columnNumber = position;
			this.column = convertToAlphabet(columnNumber);
		}
	}

	/**
	 * 文字列表記により構築を行うコンストラクタ
	 *
	 * @param address アドレス(A1形式)
	 */
	public Address(String address) {

		if (address == null) {
			throw new Poi2ccRuntimeException("アドレスにnullが指定されました");
		}

		String column = Poi2ccConstants.EMPTY_STRING;
		int length = address.length();

		if (length < 1) {
			throw new Poi2ccRuntimeException("アドレスが指定されていません");
		}

		int i = 0;

		for (; i < length; i++) {
			char c = address.charAt(i);

			if (Poi2ccUtil.isAsciiAlpha(c) == true) {
				column = column + Character.toString(c);
			} else {
				break;
			}
		}

		this.column = column;
		this.columnNumber = convertToNumber(column);
		this.rowNumber = Integer.parseInt(address.substring(i)) - 1;
	}

	/**
	 * 文字列表記から指定した相対位置で生成するコンストラクタ
	 *
	 * @param address 文字列表記によるアドレス(A1表記)
	 * @param relativeColumnNumber 相対列番号
	 * @param relativeRowNumber 相対行番号
	 */
	public Address(String address, int relativeColumnNumber, int relativeRowNumber) {

		this(address);

		setColumnNumber(columnNumber + relativeColumnNumber);
		setRowNumber(rowNumber + relativeRowNumber);
	}

	/**
	 * 数値指定による生成コンストラクタ
	 *
	 * @param columnNumber 列番号
	 * @param rowNumber 行番号
	 */
	public Address(int columnNumber, int rowNumber) {

		setAddress(columnNumber, rowNumber);
	}

	/**
	 * アドレス位置を変更します。
	 *
	 * @param columnNumber 列番号
	 * @param rowNumber 行番号
	 */
	public void setAddress(int columnNumber, int rowNumber) {

		if (columnNumber < 0) {
			columnNumber = 0;
		}
		if (rowNumber < 0) {
			rowNumber = 0;
		}

		setColumnNumber(columnNumber);
		setRowNumber(rowNumber);
	}

	public Address changeRow(int rowNumber) {

		Address address = new Address();

		address.setColumnNumber(this.columnNumber);
		address.setRowNumber(rowNumber);

		return address;
	}

	public Address changeColumn(int columnNumber) {

		Address address = new Address();

		address.setColumnNumber(columnNumber);
		address.setRowNumber(this.rowNumber);

		return address;
	}

	public Address changeColumn(String column) {

		Address address = new Address();

		int columnNumber = convertToNumber(column);
		address.setColumnNumber(columnNumber);
		address.setRowNumber(this.rowNumber);

		return address;
	}

	/**
	 * 数値をアルファベット表記に変換します。
	 *
	 * @param number 数値
	 * @return アルファベット表記
	 */
	public static String convertToAlphabet(int number) {

		if (number < 0) {
			return Poi2ccConstants.EMPTY_STRING;
		}

		int sum = number;
		char c = 0;
		String column = Poi2ccConstants.EMPTY_STRING;

		while (sum >= Poi2ccConstants.ALPHABET_NUMBER) {
			c = (char) ((sum % Poi2ccConstants.ALPHABET_NUMBER) + 'A');
			column = Character.toString(c) + column;
			sum = sum / Poi2ccConstants.ALPHABET_NUMBER - 1;
		}

		c = (char) ((sum % Poi2ccConstants.ALPHABET_NUMBER) + 'A');
		column = Character.toString(c) + column;

		return column;
	}

	/**
	 * アルファベット表記を数値に変換します。
	 *
	 * @param alphabet アルファベット表記
	 * @return 数値
	 */
	public static int convertToNumber(String alphabet) {

		char[] charArray = alphabet.toCharArray();
		int length = charArray.length - 1;
		int column = 0;
		int[] power = new int[] {1, Poi2ccConstants.ALPHABET_NUMBER, Poi2ccConstants.ALPHABET_NUMBER * Poi2ccConstants.ALPHABET_NUMBER};

		for (int i = 0; i <= length; i++) {
			int c = Character.toUpperCase(charArray[i]) - 'A' + 1;

			column = column + c * power[length - i];
		}

		return column - 1;
	}

	/**
	 * オブジェクトをクローンします。
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 *
	 * @return クローンされたアドレス情報
	 * @see java.lang.Object#clone()
	 */
	public Address clone() {

		return Poi2ccUtil.cloneObject(this);
	}

	/**
	 * 現在のアドレスから相対移動したアドレス情報を生成して返却します。
	 *
	 * @param columnNumber 相対移動列数
	 * @param rowNumber 相対移動行数
	 * @return 相対移動結果のアドレス
	 */
	public Address step(int columnNumber, int rowNumber) {

		Address address = new Address();

		try {
			address.rowNumber = this.rowNumber + rowNumber;
			address.columnNumber = this.columnNumber + columnNumber;
		} catch (Exception e) {
			this.rowNumber = -1;
			this.columnNumber = -1;
		}

		return address;
	}

	/**
	 * 現在のアドレスから相対移動したアドレス情報を生成して返却します。
	 *
	 * @param rowNumber 相対移動行数
	 * @return 相対移動結果のアドレス
	 */
	public Address stepRow(int rowNumber) {

		return step(0, rowNumber);
	}

	/**
	 * 現在のアドレスから相対移動したアドレス情報を生成して返却します。
	 *
	 * @param columnNumber 相対移動列数
	 * @return 相対移動結果のアドレス
	 */
	public Address stepColumn(int columnNumber) {

		return step(columnNumber, 0);
	}

	/**
	 * アドレス情報を文字列表現で返却します。
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 *
	 * @return インスタンス情報の文字列結果
	 * @see java.lang.Object#toString()
	 */
	public String toString() {

		if (columnNumber < 0) {
			return Integer.toString(rowNumber + 1);
		} else if (rowNumber < 0) {
			return column;
		}

		column = convertToAlphabet(columnNumber);

		return column + (rowNumber + 1);
	}

	/**
	 * 列名(アルファベット表記)を取得します。
	 *
	 * @return 列名(アルファベット表記)
	 */
	public String getColumn() {

		column = convertToAlphabet(columnNumber);

		return column;
	}

	/**
	 * 列名(アルファベット表記)を設定します。
	 *
	 * @param column 列名(アルファベット表記)
	 */
	public void setColumn(String column) {

		this.column = column;
		columnNumber = convertToNumber(column);
	}

	/**
	 * 列番号を取得します。
	 *
	 * @return 列番号
	 */
	public int getColumnNumber() {

		return columnNumber;
	}

	/**
	 * 列番号を設定します。
	 *
	 * @param columnNumber 列番号
	 */
	public void setColumnNumber(int columnNumber) {

		this.columnNumber = columnNumber;
		this.column = convertToAlphabet(columnNumber);
	}

	/**
	 * 行番号を取得します。
	 *
	 * @return 行番号
	 */
	public int getRowNumber() {

		return rowNumber;
	}

	/**
	 * 行番号を設定します。
	 *
	 * @param rowNumber 行番号
	 */
	public void setRowNumber(int rowNumber) {

		this.rowNumber = rowNumber;
	}

	/**
	 * hashCode
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 *
	 * @see java.lang.Object#hashCode()
	 */
	@Override
	public int hashCode() {

		final int prime = 31;
		int result = 1;
		result = prime * result + columnNumber;
		result = prime * result + rowNumber;
		return result;
	}

	/**
	 * equals
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 *
	 * @see java.lang.Object#equals(java.lang.Object)
	 */
	@Override
	public boolean equals(Object obj) {

		if (this == obj) {
			return true;
		}
		if (obj == null) {
			return false;
		}
		if (getClass() != obj.getClass()) {
			return false;
		}
		Address other = (Address) obj;
		if (columnNumber != other.columnNumber) {
			return false;
		}
		if (rowNumber != other.rowNumber) {
			return false;
		}
		return true;
	}
}
