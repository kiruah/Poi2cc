package com.kiruah.poi2cc.storage.sub;

import java.io.Serializable;

/**
 * 結合セル範囲
 * 
 * @author Kiruah
 */
public class CellRange
		implements
			Serializable {
	
	/** シリアルバージョンUID */
	private static final long serialVersionUID = -7331658892547927359L;
	
	/** 開始列 */
	private int beginColumn = 0;
	
	/** 開始行 */
	private int beginRow = 0;
	
	/** 終了列 */
	private int lastColumn = 0;
	
	/** 終了行 */
	private int lastRow = 0;
	
	/**
	 * 結合セル範囲 コンストラクタ
	 */
	public CellRange() {
	
	}
	
	/**
	 * 結合セル範囲 コンストラクタ
	 * 
	 * @param beginColumn 開始列
	 * @param beginRow 開始行
	 * @param lastColumn 終了列
	 * @param lastRow 終了行
	 */
	public CellRange(
			int beginColumn,
			int beginRow,
			int lastColumn,
			int lastRow) {
	
		this.beginColumn = beginColumn;
		this.beginRow = beginRow;
		this.lastColumn = lastColumn;
		this.lastRow = lastRow;
	}
	
	/**
	 * 開始列を取得します。
	 * @return 開始列
	 */
	public int getBeginColumn() {
	
		return beginColumn;
	}
	
	/**
	 * 開始列を設定します。
	 * @param beginColumn 開始列
	 */
	public void setBeginColumn(
			int beginColumn) {
	
		this.beginColumn = beginColumn;
	}
	
	/**
	 * 開始行を取得します。
	 * @return 開始行
	 */
	public int getBeginRow() {
	
		return beginRow;
	}
	
	/**
	 * 開始行を設定します。
	 * @param beginRow 開始行
	 */
	public void setBeginRow(
			int beginRow) {
	
		this.beginRow = beginRow;
	}
	
	/**
	 * 終了列を取得します。
	 * @return 終了列
	 */
	public int getLastColumn() {
	
		return lastColumn;
	}
	
	/**
	 * 終了列を設定します。
	 * @param lastColumn 終了列
	 */
	public void setLastColumn(
			int lastColumn) {
	
		this.lastColumn = lastColumn;
	}
	
	/**
	 * 終了行を取得します。
	 * @return 終了行
	 */
	public int getLastRow() {
	
		return lastRow;
	}
	
	/**
	 * 終了行を設定します。
	 * @param lastRow 終了行
	 */
	public void setLastRow(
			int lastRow) {
	
		this.lastRow = lastRow;
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
		result = prime * result + beginColumn;
		result = prime * result + beginRow;
		result = prime * result + lastColumn;
		result = prime * result + lastRow;
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
		if (obj == null) {
			return false;
		}
		if (getClass() != obj.getClass()) {
			return false;
		}
		CellRange other = (CellRange) obj;
		if (beginColumn != other.beginColumn) {
			return false;
		}
		if (beginRow != other.beginRow) {
			return false;
		}
		if (lastColumn != other.lastColumn) {
			return false;
		}
		if (lastRow != other.lastRow) {
			return false;
		}
		return true;
	}
	
}
