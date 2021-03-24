package com.kiruah.poi2cc.storage;

import java.io.Serializable;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import com.kiruah.poi2cc.Poi2ccUtil;

/**
 * セルを保持するためのエリア領域を提供します。<br/>
 * 同時に保持しているセルに対する編集処理を提供します。
 *
 * @author Kiruah
 * @since 0.9
 */
public class CellSet implements Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = 1747045611125337339L;

	/** 保持可能な最大行数 */
	public static final int MAX_ROW = 1048576;

	/** 保持可能な最大桁数 */
	public static final int MAX_COLUMN = 16384;

	/** 保持しているセル情報 */
	protected LinkedHashMap<Address, Cell> cellMap = new LinkedHashMap<Address, Cell>();

	/** この範囲内での最大桁数 */
	protected int maxColumn = 0;

	/** この範囲内での最大行数 */
	protected int maxRow = 0;

	/**
	 * デフォルトコンストラクタ
	 */
	public CellSet() {

	}

	/**
	 * 編集された状態になったかどうかを得るために事前に実施する編集済みフラグの初期化処理
	 * <p>
	 * 編集されたかどうかを判定する必要がない場合は実施不要の初期化処理です。
	 * </p>
	 */
	public void initializeEdited() {

		Set<Entry<Address, Cell>> set = cellMap.entrySet();
		Iterator<Entry<Address, Cell>> iterator = set.iterator();

		while (iterator.hasNext()) {
			Entry<Address, Cell> entry = iterator.next();

			entry.getValue().setEdited(false);
		}
	}

	/**
	 * 指定アドレスにセルを設定します。
	 *
	 * @param address 設定先のアドレス
	 * @param cell 設定するセル情報
	 */
	public void setCell(Address address, Cell cell) {

		if (maxColumn <= address.getColumnNumber()) {
			maxColumn = address.getColumnNumber() + 1;
		}
		if (maxRow <= address.getRowNumber()) {
			maxRow = address.getRowNumber() + 1;
		}

		if (cell.getAddress() == null) {
			cell.setAddress(address);
		} else {
			if (cell.getMergedCell() != null) {
				int diffColumn = address.getColumnNumber() - cell.getAddress().getColumnNumber();
				int diffRow = address.getRowNumber() - cell.getAddress().getRowNumber();

				cell.getMergedCell().setBeginColumn(cell.getMergedCell().getBeginColumn() + diffColumn);
				cell.getMergedCell().setBeginRow(cell.getMergedCell().getBeginRow() + diffRow);
				cell.getMergedCell().setLastColumn(cell.getMergedCell().getLastColumn() + diffColumn);
				cell.getMergedCell().setLastRow(cell.getMergedCell().getLastRow() + diffRow);
			}

			cell.getAddress().setColumnNumber(address.getColumnNumber());
			cell.getAddress().setRowNumber(address.getRowNumber());
		}

		cell.edited = true;
		cellMap.put(address, cell);
	}

	/**
	 * 指定アドレスの相対位置にセルを設定します。
	 *
	 * @param address 設定先のアドレス
	 * @param relativeColumnNumber 相対列番号
	 * @param relativeRowNumber 相対行番号
	 * @param cell 設定するセル情報
	 */
	public void setCell(Address address, int relativeColumnNumber, int relativeRowNumber, Cell cell) {

		Address newAddress = address.step(relativeColumnNumber, relativeRowNumber);

		setCell(newAddress, cell);
	}

	/**
	 * 指定アドレスにセルを設定します。
	 *
	 * @param columnNumber 列番号
	 * @param rowNumber 行番号
	 * @param cell 設定するセル情報
	 */
	public void setCell(int columnNumber, int rowNumber, Cell cell) {

		Address address = new Address(columnNumber, rowNumber);

		setCell(address, cell);
	}

	/**
	 * 指定し地のセル情報を取得します。<br/>
	 * もしセル情報がない場合は空のセルを返却します。<br/>
	 * また、存在しえない不正なアドレスを指定した場合はnullを返却します。<br/>
	 *
	 * @param address セル情報の取得先アドレス
	 * @return セル情報
	 */
	public Cell getCell(Address address) {

		Cell cell = cellMap.get(address);

		if (cell == null) {
			if (address.getColumnNumber() < 0 || address.getRowNumber() < 0) {
				return null;
			}

			cell = new Cell(address);

			setCell(address, cell);
		}

		return cell;
	}

	/**
	 * 指定し地のセル情報を取得します。<br/>
	 * もしセル情報がない場合は空のセルを返却します。<br/>
	 * また、存在しえない不正なアドレスを指定した場合はnullを返却します。<br/>
	 *
	 * @param address セル情報の取得先アドレス
	 * @param relativeColumnNumber 相対列位置
	 * @param relativeRowNumber 相対行位置
	 * @return セル情報
	 */
	public Cell getCell(Address address, int relativeColumnNumber, int relativeRowNumber) {

		Address newAddress = address.step(relativeColumnNumber, relativeRowNumber);

		return getCell(newAddress);
	}

	/**
	 * 指定し地のセル情報を取得します。<br/>
	 * もしセル情報がない場合は空のセルを返却します。<br/>
	 * また、存在しえない不正なアドレスを指定した場合はnullを返却します。<br/>
	 *
	 * @param columnNumber 列番号
	 * @param rowNumber 行番号
	 * @return セル情報
	 */
	public Cell getCell(int columnNumber, int rowNumber) {

		Address address = new Address(columnNumber, rowNumber);

		return getCell(address);
	}

	/**
	 * 指定し地のセル情報を取得します。<br/>
	 * もしセル情報がない場合や存在しえない不正なアドレスを指定した場合はnullを返却します。<br/>
	 *
	 * @param address
	 * @return
	 */
	public Cell getActualCell(Address address) {

		return cellMap.get(address);
	}

	/**
	 * 指定した行のセルセットを返却します。
	 *
	 * @param rowNumber 行番号
	 * @param length 取得するセル数
	 * @return セルセット
	 */
	public CellSet getRows(int rowNumber, int length) {

		return getRange(0, rowNumber, maxColumn, rowNumber + length - 1);
	}

	/**
	 * 指定した列のセルセットを返却します。
	 *
	 * @param columnNumber 列番号
	 * @param length 取得するセル数
	 * @return セルセット
	 */
	public CellSet getColumns(int columnNumber, int length) {

		return getRange(columnNumber, 0, columnNumber + length - 1, maxRow);
	}

	/**
	 * 指定した列のセルセットを返却します。
	 *
	 * @param column 列名(アルファベット表記)
	 * @param length 取得するセル数
	 * @return セルセット
	 */
	public CellSet getColumns(String column, int length) {

		int columnNumber = Address.convertToNumber(column);

		return getRange(columnNumber, 0, columnNumber + length - 1, maxRow);
	}

	/**
	 * 指定したアドレス範囲内のセルセットを取得します。
	 *
	 * @param begin 取得開始範囲
	 * @param end 取得終了範囲
	 * @return セルセット
	 */
	public CellSet getRange(Address begin, Address end) {

		return getRange(begin.getColumnNumber(), begin.getRowNumber(), end.getColumnNumber(), end.getRowNumber());
	}

	/**
	 * 指定したアドレス範囲内のセルセットを取得します。
	 *
	 * @param beginColumnNumber 取得開始列番号
	 * @param beginRowNumber 取得開始行番号
	 * @param endColumnNumber 取得終了列番号
	 * @param endRowNumber 取得終了行番号
	 * @return セルセット
	 */
	public CellSet getRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber) {

		CellSet set = new CellSet();

		for (int i = beginRowNumber; i < endRowNumber; i++) {
			for (int j = beginColumnNumber; j < endColumnNumber; j++) {
				Address localAddress = new Address(j, i);
				Cell cell = cellMap.get(localAddress);

				Address address = new Address(j - beginColumnNumber, i - beginRowNumber);

				if (cell != null) {
					set.setCell(address, cell);
				}
			}
		}

		return set.clone();
	}

	/**
	 * 指定行のセルセットをクローンし取得します。
	 *
	 * @param rowNumber 取得行番号
	 */
	public CellSet getNewRows(int rowNumber, int length) {

		CellSet set = getRows(rowNumber, length);

		return set.clone();
	}

	/**
	 * 指定列のセルセットをクローンし取得します。
	 *
	 * @param columnNumber 取得列番号
	 * @param length 長さ
	 * @return セルセット
	 */
	public CellSet getNewColumns(int columnNumber, int length) {

		CellSet set = getColumns(columnNumber, length);

		return set.clone();
	}

	/**
	 * 指定列のセルセットをクローンし取得します。
	 *
	 * @param column 取得列名(アルファベット表記)
	 * @param length 長さ
	 * @return セルセット
	 */
	public CellSet getNewColumns(String column, int length) {

		CellSet set = getColumns(column, length);

		return set.clone();
	}

	/**
	 * 指定範囲のセルセットをクローンし取得します。
	 *
	 * @param begin 取得開始範囲アドレス
	 * @param end 取得終了範囲アドレス
	 * @return セルセット
	 */
	public CellSet getNewRange(Address begin, Address end) {

		CellSet set = getRange(begin, end);

		return set.clone();
	}

	/**
	 * 指定範囲のセルセットをクローンし取得します。
	 *
	 * @param beginColumnNumber 取得開始列番号
	 * @param beginRowNumber 取得開始行番号
	 * @param endColumnNumber 取得終了列番号
	 * @param endRowNumber 取得終了行番号
	 * @return セルセット
	 */
	public CellSet getNewRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber) {

		CellSet set = getRange(beginColumnNumber, beginRowNumber, endColumnNumber, endRowNumber);

		return set;
	}

	/**
	 * セルの内容を移動します。
	 *
	 * @param beginColumnNumber 移動元開始列番号
	 * @param beginRowNumber 移動元開始行番号
	 * @param endColumnNumber 移動元終了列番号
	 * @param endRowNumber 移動元終了行番号
	 * @param moveBeginColumnNumber 移動先列番号
	 * @param moveBeginRowNumber 移動先行番号
	 */
	public void moveRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber, int moveBeginColumnNumber, int moveBeginRowNumber) {

		CellSet cellSet = getRange(beginColumnNumber, beginRowNumber, endColumnNumber, endRowNumber);

		clearRange(beginColumnNumber, beginRowNumber, endColumnNumber, endRowNumber);

		Set<Entry<Address, Cell>> set = cellSet.getCellMap().entrySet();
		Iterator<Entry<Address, Cell>> iterator = set.iterator();

		while (iterator.hasNext() == true) {
			Entry<Address, Cell> entry = iterator.next();
			Address address = entry.getKey();
			Cell cell = entry.getValue();

			address = address.step(moveBeginColumnNumber, moveBeginRowNumber);

			setCell(address, cell);
		}
	}

	/**
	 * 指定範囲のセルのコピーを行います。
	 *
	 * @param beginColumnNumber コピー元開始列番号
	 * @param beginRowNumber コピー元開始行番号
	 * @param endColumnNumber コピー元終了列番号
	 * @param endRowNumber コピー元終了行番号
	 * @param copyBeginColumnNumber コピー先開始列番号
	 * @param copyBeginRowNumber コピー先開始行番号
	 */
	public void copyRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber, int copyBeginColumnNumber, int copyBeginRowNumber) {

		CellSet cellSet = getRange(beginColumnNumber, beginRowNumber, endColumnNumber, endRowNumber);

		Set<Entry<Address, Cell>> set = cellSet.getCellMap().entrySet();
		Iterator<Entry<Address, Cell>> iterator = set.iterator();

		while (iterator.hasNext() == true) {
			Entry<Address, Cell> entry = iterator.next();
			Address address = entry.getKey();
			Cell cell = entry.getValue();

			address = address.step(copyBeginColumnNumber, copyBeginRowNumber);

			setCell(address, cell);
		}
	}

	public void clearRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber) {

		Address address = new Address(0, 0);

		for (int i = beginRowNumber; i <= endRowNumber; i++) {
			for (int j = beginColumnNumber; j <= endColumnNumber; j++) {
				address.setAddress(j, i);
				cellMap.remove(address);
			}
		}
	}

	public void slideOverwriteRange(int beginColumnNumber, int beginRowNumber, int slideNumber, boolean moveHorizontal) {

		CellSet cellSet = null;
		int diffColumn = 0;
		int diffRow = 0;

		if (moveHorizontal == true) {
			cellSet = getRange(beginColumnNumber + 1, beginRowNumber, maxColumn, beginRowNumber);
			diffColumn = slideNumber;
		} else {
			cellSet = getRange(beginColumnNumber, beginRowNumber + 1, beginColumnNumber, maxRow);
			diffRow = slideNumber;
		}

		Set<Entry<Address, Cell>> set = cellSet.getCellMap().entrySet();
		Iterator<Entry<Address, Cell>> iterator = set.iterator();

		while (iterator.hasNext() == true) {
			Entry<Address, Cell> entry = iterator.next();
			Address address = entry.getKey();
			Cell cell = entry.getValue();

			address = address.step(diffColumn, diffRow);

			setCell(address, cell);
		}
	}

	public void deleteRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber, boolean moveHorizontal) {

		if (moveHorizontal == true) {
			slideOverwriteRange(endColumnNumber, endRowNumber, beginColumnNumber - endColumnNumber, moveHorizontal);
		} else {
			slideOverwriteRange(endColumnNumber, endRowNumber, beginRowNumber - endRowNumber, moveHorizontal);
		}
	}

	public void spaceRange(int beginColumnNumber, int beginRowNumber, int endColumnNumber, int endRowNumber, boolean moveHorizontal) {

		if (moveHorizontal == true) {
			slideOverwriteRange(beginColumnNumber, beginRowNumber, endColumnNumber - beginColumnNumber, moveHorizontal);
		} else {
			slideOverwriteRange(beginColumnNumber, beginRowNumber, endRowNumber - beginRowNumber, moveHorizontal);
		}
		clearRange(beginColumnNumber, beginRowNumber, endColumnNumber, endRowNumber);
	}

	public void insertRange(int beginColumnNumber, int beginRowNumber, boolean moveHorizontal, CellSet cellSet) {

		int endColumnNumber = beginColumnNumber + cellSet.maxColumn;
		int endRowNumber = beginRowNumber + cellSet.maxRow;

		if (moveHorizontal == true) {
			slideOverwriteRange(beginColumnNumber, beginRowNumber, endColumnNumber - beginColumnNumber, moveHorizontal);
		} else {
			slideOverwriteRange(beginColumnNumber, beginRowNumber, endRowNumber - beginRowNumber, moveHorizontal);
		}

		paste(beginColumnNumber, beginRowNumber, cellSet);
	}

	public void paste(Address address, CellSet cellSet) {

		paste(address.getColumnNumber(), address.getRowNumber(), cellSet);
	}

	public void paste(int beginColumnNumber, int beginRowNumber, CellSet cellSet) {

		cellSet = cellSet.clone();

		Set<Entry<Address, Cell>> set = cellSet.getCellMap().entrySet();
		Iterator<Entry<Address, Cell>> iterator = set.iterator();

		while (iterator.hasNext() == true) {
			Entry<Address, Cell> entry = iterator.next();
			Address address = entry.getKey();
			Cell cell = entry.getValue();

			address = address.step(beginColumnNumber, beginRowNumber);

			setCell(address, cell);
		}
	}

	/**
	 * 指定したセルの内容を消去します。
	 *
	 * @param address 消去アドレス
	 */
	public void clear(Address address) {

		clearRange(address.getColumnNumber(), address.getRowNumber(), 1, 1);
	}

	/**
	 * 指定したセル範囲の内容を消去します。
	 *
	 * @param beginAddress 消去開始アドレス
	 * @param endAddress 消去終了アドレス
	 */
	public void clear(Address beginAddress, Address endAddress) {

		clearRange(beginAddress.getColumnNumber(), beginAddress.getRowNumber(), endAddress.getColumnNumber(), endAddress.getRowNumber());
	}

	/**
	 * 指定した行の内容を消去します。
	 *
	 * @param beginAddress 開始アドレス(行のみ使用)
	 * @param length 列数
	 */
	public void clearRow(Address beginAddress, int length) {

		clearRange(0, beginAddress.getRowNumber(), maxColumn, beginAddress.getRowNumber() + length - 1);
	}

	/**
	 * 指定した行の内容を消去します。
	 *
	 * @param beginAddress 開始アドレス(行のみ使用)
	 * @param endAddress 終了アドレス(行のみ使用)
	 */
	public void clearRow(Address beginAddress, Address endAddress) {

		clearRange(0, beginAddress.getRowNumber(), maxColumn, endAddress.getRowNumber());
	}

	/**
	 * 指定した行の内容を消去します。
	 *
	 * @param rowNumber 行数
	 * @param length 長さ
	 */
	public void clearRow(int rowNumber, int length) {

		clearRange(0, rowNumber, maxColumn, rowNumber + length - 1);
	}

	/**
	 * 指定した列の内容を消去します。
	 *
	 * @param beginAddress 開始アドレス(列のみ使用)
	 * @param length 列数
	 */
	public void clearColumn(Address beginAddress, int length) {

		clearRange(beginAddress.getColumnNumber(), 0, beginAddress.getColumnNumber() + length - 1, maxRow);
	}

	/**
	 * 指定した列の内容を消去します。
	 *
	 * @param beginAddress 開始アドレス(列のみ使用)
	 * @param endAddress 終了アドレス(列のみ使用)
	 */
	public void clearColumn(Address beginAddress, Address endAddress) {

		clearRange(beginAddress.getColumnNumber(), 0, endAddress.getColumnNumber(), maxRow);
	}

	/**
	 * 指定した列の内容を消去します。
	 *
	 * @param column 列番号
	 * @param length 長さ
	 */
	public void clearColumn(int columnNumber, int length) {

		clearRange(columnNumber, 0, columnNumber + length - 1, maxRow);
	}

	/**
	 * 指定した列の内容を消去します。
	 *
	 * @param column 列名(アルファベット表記)
	 * @param length 長さ
	 */
	public void clearColumn(String column, int length) {

		int columnNumber = Address.convertToNumber(column);

		clearRange(columnNumber, 0, columnNumber + length - 1, maxRow);
	}

	/**
	 * 名前 - セル取得位置のマップ情報を引き渡すことで、一括で指定したアドレスのセルを取得します。
	 *
	 * @param cellGetterMap セル取得先のマップ名
	 * @return 名前 - セル のマップ情報
	 */
	public Map<String, Cell> getCell(Map<String, Address> cellGetterMap) {

		return getCell(cellGetterMap, 0, 0);
	}

	/**
	 * 名前 - セル取得位置のマップ情報を引き渡すことで、一括で指定したアドレスのセルを取得します。
	 *
	 * @param cellGetterMap セル取得先のマップ名
	 * @param skipColumn 指定したアドレスからの相対桁位置
	 * @param skipRow 指定したアドレスからの相対行位置
	 * @return 名前 - セル のマップ情報
	 */
	public Map<String, Cell> getCell(Map<String, Address> cellGetterMap, int skipColumn, int skipRow) {

		Iterator<Entry<String, Address>> iterator = Poi2ccUtil.getIterator(cellGetterMap);
		Map<String, Cell> cellSetterMap = new LinkedHashMap<String, Cell>();

		while (iterator.hasNext() == true) {
			Entry<String, Address> entry = iterator.next();

			Address address = entry.getValue().step(skipColumn, skipRow);

			Cell cell = cellMap.get(address);

			cellSetterMap.put(entry.getKey(), cell);
		}

		return cellSetterMap;
	}

	public boolean isRange(Address address) {

		int column = address.getColumnNumber();
		if (column < 0 || maxColumn <= column) {
			return false;
		}

		int row = address.getRowNumber();
		if (row < 0 || maxRow <= row) {
			return false;
		}

		return true;
	}

	/**
	 * デープコピーし返却します。
	 *
	 * @return デープコピーされたCellSet
	 * @see java.lang.Object#clone()
	 */
	public CellSet clone() {

		return Poi2ccUtil.cloneObject(this);
	}

	/**
	 * 現在のCellSetクラス名を返却します。
	 *
	 * @return クラス名
	 * @see java.lang.Object#toString()
	 */
	@Override
	public String toString() {

		return getClass().getName();
	}

	/**
	 * 保持しているセル情報を取得します。
	 * @return 保持しているセル情報
	 */
	public LinkedHashMap<Address, Cell> getCellMap() {

		return cellMap;
	}

	/**
	 * 保持しているセル情報を設定します。
	 * @param cellMap 保持しているセル情報
	 */
	public void setCellMap(LinkedHashMap<Address, Cell> cellMap) {

		this.cellMap = cellMap;
	}

	/**
	 * この範囲内での最大桁数を取得します。
	 * @return この範囲内での最大桁数
	 */
	public int getMaxColumn() {

		return maxColumn;
	}

	/**
	 * この範囲内での最大桁数を設定します。
	 * @param maxColumn この範囲内での最大桁数
	 */
	public void setMaxColumn(int maxColumn) {

		this.maxColumn = maxColumn;
	}

	/**
	 * この範囲内での最大行数を取得します。
	 * @return この範囲内での最大行数
	 */
	public int getMaxRow() {

		return maxRow;
	}

	/**
	 * この範囲内での最大行数を設定します。
	 * @param maxRow この範囲内での最大行数
	 */
	public void setMaxRow(int maxRow) {

		this.maxRow = maxRow;
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
		result = prime * result + ((cellMap == null) ? 0 : cellMap.hashCode());
		result = prime * result + maxColumn;
		result = prime * result + maxRow;
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
		CellSet other = (CellSet) obj;
		if (cellMap == null) {
			if (other.cellMap != null) {
				return false;
			}
		} else if (!cellMap.equals(other.cellMap)) {
			return false;
		}
		if (maxColumn != other.maxColumn) {
			return false;
		}
		if (maxRow != other.maxRow) {
			return false;
		}
		return true;
	}
}
