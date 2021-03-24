package com.kiruah.poi2cc.storage;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

import com.kiruah.poi2cc.Poi2ccConstants;
import com.kiruah.poi2cc.Poi2ccUtil;
import com.kiruah.poi2cc.storage.sub.SheetStyle;

/**
 * Excelシートの情報を保持します。
 *
 * @author Kiruah
 */
public class Sheet extends CellSet
		implements
			Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = -2029018211085633075L;

	/** シート名 */
	protected String name = Poi2ccConstants.EMPTY_STRING;

	/** 行高さ情報 */
	protected Map<Integer, Integer> rowSizeMap = new HashMap<Integer, Integer>();

	/** 列幅情報 */
	protected Map<Integer, Integer> columnSizeMap = new HashMap<Integer, Integer>();

	/** デフォルト行高さ */
	protected short defaultRowSize = 10;

	/** デフォルト列幅 */
	protected int defaultColumnSize = 8;

	/** シートのスタイル情報 */
	protected SheetStyle style = new SheetStyle();

	/**
	 * シート情報を生成するコンストラクタ
	 *
	 * @param name シート名
	 */
	public Sheet(
			String name) {

		this.name = name;
	}

	/**
	 * シート情報をクローンします。
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 * @return シート情報
	 * @see com.kiruah.poi2cc.storage.CellSet#clone()
	 */
	@Override
	public Sheet clone() {

		return Poi2ccUtil.cloneObject(this);
	}

	/**
	 * シート情報を文字列化し返却します。
	 *
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 * @return 文字列化されたシート情報
	 * @see com.kiruah.poi2cc.storage.CellSet#toString()
	 */
	@Override
	public String toString() {

		return getClass().getName() + ":" + name;
	}

	/**
	 * シート名を取得します。
	 * @return シート名
	 */
	public String getName() {

		return name;
	}

	/**
	 * シート名を設定します。
	 * @param name シート名
	 */
	public void setName(
			String name) {

		this.name = name;
	}

	/**
	 * 行高さ情報を取得します。
	 * @return 行高さ情報
	 */
	public Map<Integer, Integer> getRowSizeMap() {

		return rowSizeMap;
	}

	/**
	 * 行高さ情報を設定します。
	 * @param rowSizeMap 行高さ情報
	 */
	public void setRowSizeMap(
			Map<Integer, Integer> rowSizeMap) {

		this.rowSizeMap = rowSizeMap;
	}

	/**
	 * 列幅情報を取得します。
	 * @return 列幅情報
	 */
	public Map<Integer, Integer> getColumnSizeMap() {

		return columnSizeMap;
	}

	/**
	 * 列幅情報を設定します。
	 * @param columnSizeMap 列幅情報
	 */
	public void setColumnSizeMap(
			Map<Integer, Integer> columnSizeMap) {

		this.columnSizeMap = columnSizeMap;
	}

	/**
	 * デフォルト行高さを取得します。
	 * @return デフォルト行高さ
	 */
	public short getDefaultRowSize() {

		return defaultRowSize;
	}

	/**
	 * デフォルト行高さを設定します。
	 * @param defaultRowSize デフォルト行高さ
	 */
	public void setDefaultRowSize(
			short defaultRowSize) {

		this.defaultRowSize = defaultRowSize;
	}

	/**
	 * デフォルト列幅を取得します。
	 * @return デフォルト列幅
	 */
	public int getDefaultColumnSize() {

		return defaultColumnSize;
	}

	/**
	 * デフォルト列幅を設定します。
	 * @param defaultColumnSize デフォルト列幅
	 */
	public void setDefaultColumnSize(
			int defaultColumnSize) {

		this.defaultColumnSize = defaultColumnSize;
	}

	/**
	 * シートのスタイル情報を取得します。
	 * @return シートのスタイル情報
	 */
	public SheetStyle getStyle() {

		return style;
	}

	/**
	 * シートのスタイル情報を設定します。
	 * @param style シートのスタイル情報
	 */
	public void setStyle(
			SheetStyle style) {

		this.style = style;
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
		int result = super.hashCode();
		result = prime * result + ((columnSizeMap == null) ? 0 : columnSizeMap.hashCode());
		result = prime * result + defaultColumnSize;
		result = prime * result + defaultRowSize;
		result = prime * result + ((name == null) ? 0 : name.hashCode());
		result = prime * result + ((rowSizeMap == null) ? 0 : rowSizeMap.hashCode());
		result = prime * result + ((style == null) ? 0 : style.hashCode());
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
	public boolean equals(
			Object obj) {

		if (this == obj) {
			return true;
		}
		if (!super.equals(obj)) {
			return false;
		}
		if (getClass() != obj.getClass()) {
			return false;
		}
		Sheet other = (Sheet) obj;
		if (columnSizeMap == null) {
			if (other.columnSizeMap != null) {
				return false;
			}
		} else if (!columnSizeMap.equals(other.columnSizeMap)) {
			return false;
		}
		if (defaultColumnSize != other.defaultColumnSize) {
			return false;
		}
		if (defaultRowSize != other.defaultRowSize) {
			return false;
		}
		if (name == null) {
			if (other.name != null) {
				return false;
			}
		} else if (!name.equals(other.name)) {
			return false;
		}
		if (rowSizeMap == null) {
			if (other.rowSizeMap != null) {
				return false;
			}
		} else if (!rowSizeMap.equals(other.rowSizeMap)) {
			return false;
		}
		if (style == null) {
			if (other.style != null) {
				return false;
			}
		} else if (!style.equals(other.style)) {
			return false;
		}
		return true;
	}
}
