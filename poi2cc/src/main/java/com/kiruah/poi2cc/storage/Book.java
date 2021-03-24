package com.kiruah.poi2cc.storage;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.kiruah.poi2cc.Poi2ccConstants;
import com.kiruah.poi2cc.Poi2ccUtil;

/**
 * Excelブック
 *
 * @author Kiruah
 */
public class Book implements Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = -214729840412172985L;

	/** シートリスト */
	private List<Sheet> sheetList = new ArrayList<Sheet>();

	/** ブック名 */
	private String name = Poi2ccConstants.EMPTY_STRING;

	/**
	 * ブックのコンストラクタ
	 *
	 * @param name ブック名
	 */
	public Book(String name) {

		this.name = name;
	}

	/**
	 * シートを追加します。
	 *
	 * @param sheet シート
	 */
	public void addSheet(Sheet sheet) {

		sheetList.add(sheet);
	}

	/**
	 * シートを取得します。
	 *
	 * @param sheetName シート名
	 * @return シート
	 */
	public Sheet getSheet(String sheetName) {

		for (Sheet sheet : sheetList) {
			if (Poi2ccUtil.equalString(sheet.getName(), sheetName) == true) {
				return sheet;
			}
		}

		return null;
	}

	/**
	 * シート名をリネームします。
	 *
	 * @param sheet シート
	 * @param name 新しい名前
	 */
	public void renameSheet(Sheet sheet, String name) {

		sheet.setName(name);
	}

	/**
	 * シートを削除します。
	 *
	 * @param sheet 削除対象のシート
	 */
	public void removeSheet(Sheet sheet) {

		sheetList.remove(sheet);
	}

	/**
	 * シート名から対象のシートを削除します。
	 *
	 * @param sheetName シート名
	 */
	public void removeSheet(String sheetName) {

		Sheet sheet = getSheet(sheetName);
		if (sheet != null) {
			removeSheet(sheet);
		}
	}

	/**
	 * 新規シートを作成します。
	 *
	 * @param sheetName 新規シート名
	 * @return 新規シート情報
	 */
	public Sheet newSheet(String sheetName) {

		Sheet sheet = new Sheet(sheetName);

		addSheet(sheet);

		return sheet;
	}

	/**
	 * ブックをクローンします。
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 * @return クローンされたブック
	 * @see java.lang.Object#clone()
	 */
	public Book clone() {

		return Poi2ccUtil.cloneObject(this);
	}

	/**
	 * ブック情報を文字列化し返却します。
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 * @return 文字列化されたブック情報
	 * @see java.lang.Object#toString()
	 */
	public String toString() {

		return getClass().getName() + ":" + name;
	}

	/**
	 * シートリストを取得します。
	 * @return シートリスト
	 */
	public List<Sheet> getSheetList() {

		return sheetList;
	}

	/**
	 * シートリストを設定します。
	 * @param sheetList シートリスト
	 */
	public void setSheetList(List<Sheet> sheetList) {

		this.sheetList = sheetList;
	}

	/**
	 * ブック名を取得します。
	 * @return ブック名
	 */
	public String getName() {

		return name;
	}

	/**
	 * ブック名を設定します。
	 * @param name ブック名
	 */
	public void setName(String name) {

		this.name = name;
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
		result = prime * result + ((name == null) ? 0 : name.hashCode());
		result = prime * result + ((sheetList == null) ? 0 : sheetList.hashCode());
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
		Book other = (Book) obj;
		if (name == null) {
			if (other.name != null) {
				return false;
			}
		} else if (!name.equals(other.name)) {
			return false;
		}
		if (sheetList == null) {
			if (other.sheetList != null) {
				return false;
			}
		} else if (!sheetList.equals(other.sheetList)) {
			return false;
		}
		return true;
	}
}
