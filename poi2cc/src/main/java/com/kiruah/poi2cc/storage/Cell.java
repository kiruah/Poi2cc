package com.kiruah.poi2cc.storage;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.util.Date;

import com.kiruah.poi2cc.Poi2ccConstants;
import com.kiruah.poi2cc.Poi2ccUtil;
import com.kiruah.poi2cc.storage.sub.CellFont;
import com.kiruah.poi2cc.storage.sub.CellRange;
import com.kiruah.poi2cc.storage.sub.CellStyle;

/**
 * Excelシート内のセル情報を保持します。
 *
 * @author Kiruah
 */
public class Cell implements Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = -5584341563567300951L;

	protected ValueType valueType = ValueType.NONE;

	protected boolean valueBoolean = false;

	protected String valueText = null;

	protected BigDecimal valueNumber = null;

	protected LocalDateTime valueDate = null;

	/** セルの実際の式値 */
	protected String formulaValue = null;

	/** セルに対するコメント */
	protected String comment = null;

	/** 式が埋め込まれているかどうか */
	protected boolean formula = false;

	/** エラーがあるか */
	protected boolean error = false;

	/** セルのフォント */
	protected CellFont font = null;

	/** セルのスタイル */
	protected CellStyle style = null;

	/** 編集ロック状態かどうか */
	protected boolean locked = true;

	/** 結合セル情報 */
	protected CellRange mergedCell = null;

	/** 結合セルの場合の結合先頭セル */
	protected Cell parentCell = null;

	/** 編集されたかどうか */
	protected boolean edited = false;

	/** セル位置 */
	protected Address address = null;

	/**
	 * セル コンストラクタ
	 */
	public Cell() {

		comment = null;
		font = new CellFont();
		style = new CellStyle();
	}

	/**
	 * セル コンストラクタ
	 *
	 * @param address アドレス
	 */
	public Cell(Address address) {

		this();
		this.address = address;
	}

	/**
	 * セル コンストラクタ
	 *
	 * @param address アドレス
	 * @param font フォント情報
	 * @param style セルスタイル情報
	 */
	public Cell(Address address, CellFont font, CellStyle style) {

		this.address = address;
		this.font = font;
		this.style = style;
	}

	/**
	 * セル情報をクローンします。
	 * <br/>
	 * <p>(オーバーライドメソッド)</p>
	 * @return クローンされたセル情報
	 * @see java.lang.Object#clone()
	 */
	@Override
	public Cell clone() {

		return Poi2ccUtil.cloneObject(this);
	}

	public String toString() {

		return "Cell(" + valueType + ")=" + getValue();
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を取得します。
	 * @return セルの見た目上の値(画面表示上の値)
	 */
	public Object getValue() {

		if (valueType == ValueType.NONE) {
			return "";
		} else if (valueType == ValueType.BOOLEAN) {
			return valueBoolean;
		} else if (valueType == ValueType.TEXT) {
			return valueText;
		} else if (valueType == ValueType.NUMBER) {
			return valueNumber;
		} else if (valueType == ValueType.DATE) {
			return valueDate;
		}

		return null;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(boolean value) {

		valueType = ValueType.BOOLEAN;
		valueBoolean = value;
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(CharSequence value) {

		valueType = ValueType.TEXT;
		if (value == null) {
			valueText = null;
		} else {
			valueText = value.toString();
		}
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(BigDecimal value) {

		valueType = ValueType.NUMBER;
		if (value == null) {
			valueNumber = BigDecimal.ZERO;
		} else {
			valueNumber = value;
		}
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(Date value) {

		valueType = ValueType.DATE;

		Instant instant = value.toInstant();
		valueDate = LocalDateTime.ofInstant(instant, ZoneId.systemDefault());
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(LocalDateTime value) {

		valueType = ValueType.DATE;
		valueDate = value;
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(byte value) {

		valueType = ValueType.NUMBER;
		valueNumber = new BigDecimal(value);
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(char value) {

		valueType = ValueType.TEXT;
		valueText = new String(new char[] {value});
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(short value) {

		valueType = ValueType.NUMBER;
		valueNumber = new BigDecimal(value);
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(int value) {

		valueType = ValueType.NUMBER;
		valueNumber = new BigDecimal(value);
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(long value) {

		valueType = ValueType.NUMBER;
		valueNumber = new BigDecimal(value);
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(float value) {

		valueType = ValueType.NUMBER;
		valueNumber = new BigDecimal(value);
		this.formula = false;
	}

	/**
	 * セルの見た目上の値(画面表示上の値)を設定します。
	 * @param value セルの見た目上の値(画面表示上の値)
	 */
	public void setValue(double value) {

		valueType = ValueType.NUMBER;
		valueNumber = new BigDecimal(value);
		this.formula = false;
	}

	/**
	 * valueTypeを取得します。
	 * @return valueType
	 */
	public ValueType getValueType() {

		return valueType;
	}

	/**
	 * valueTypeを設定します。
	 * @param valueType valueType
	 */
	public void setValueType(ValueType valueType) {

		this.valueType = valueType;
	}

	/**
	 * valueTextを取得します。
	 * @return valueText
	 */
	public String getValueText() {

		if (valueType == ValueType.NONE) {
			return null;
		} else if (valueType == ValueType.BOOLEAN) {
			return Boolean.toString(valueBoolean);
		} else if (valueType == ValueType.NUMBER) {
			return valueNumber.toPlainString();
		} else if (valueType == ValueType.DATE) {
			String text = valueDate.format(Poi2ccConstants.DEFAULT_DATE_TIME_FORMAT);

			return text;
		} else if (valueType == ValueType.TEXT) {
			return valueText;
		}

		return null;
	}

	/**
	 * valueTextを設定します。
	 * @param valueText valueText
	 */
	public void setValueText(String valueText) {

		this.valueText = valueText;
		this.valueType = ValueType.TEXT;
		this.formula = false;
	}

	/**
	 * valueNumberを取得します。
	 * @return valueNumber
	 */
	public BigDecimal getValueNumber() {

		if (valueType == ValueType.NONE) {
			return null;
		} else if (valueType == ValueType.BOOLEAN) {
			if (valueBoolean == true) {
				return BigDecimal.ONE;
			}

			return BigDecimal.ZERO;
		} else if (valueType == ValueType.NUMBER) {
			return valueNumber;
		} else if (valueType == ValueType.DATE) {
			ZonedDateTime zdt = valueDate.atZone(ZoneOffset.systemDefault());
			long ms = zdt.toEpochSecond();
			BigDecimal bd = new BigDecimal(ms);

			return bd;
		} else if (valueType == ValueType.TEXT) {
			try {
				BigDecimal bd = new BigDecimal(valueText);
				return bd;
			} catch (Exception e) {
				return BigDecimal.ZERO;
			}
		}

		return null;
	}

	/**
	 * valueNumberを設定します。
	 * @param valueNumber valueNumber
	 */
	public void setValueNumber(BigDecimal valueNumber) {

		this.valueNumber = valueNumber;
		this.valueType = ValueType.NUMBER;
		this.formula = false;
	}

	/**
	 * valueDateを取得します。
	 * @return valueDate
	 */
	public LocalDateTime getValueDate() {

		if (valueType == ValueType.NONE) {
			return null;
		} else if (valueType == ValueType.BOOLEAN) {

		} else if (valueType == ValueType.NUMBER) {

		} else if (valueType == ValueType.DATE) {
			return valueDate;

		} else if (valueType == ValueType.TEXT) {

		}

		return null;
	}

	/**
	 * valueDateを設定します。
	 * @param valueDate valueDate
	 */
	public void setValueDate(LocalDateTime valueDate) {

		this.valueDate = valueDate;
		this.valueType = ValueType.DATE;
		this.formula = false;
	}

	/**
	 * valueBooleanを取得します。
	 * @return valueBoolean
	 */
	public boolean getValueBoolean() {

		if (valueType == ValueType.NONE) {
			return false;
		} else if (valueType == ValueType.BOOLEAN) {
			return valueBoolean;
		} else if (valueType == ValueType.NUMBER) {
			if (BigDecimal.ZERO.equals(valueNumber) == true) {
				return false;
			}

			return true;
		} else if (valueType == ValueType.TEXT) {
			boolean value = Boolean.valueOf(valueText);
			return value;
		}

		return valueBoolean;
	}

	/**
	 * valueBooleanを設定します。
	 * @param valueBoolean valueBoolean
	 */
	public void setValueBoolean(boolean valueBoolean) {

		this.valueBoolean = valueBoolean;
		this.valueType = ValueType.BOOLEAN;
		this.formula = false;
	}

	/**
	 * セルの実際の式値を取得します。
	 * @return セルの実際の式値
	 */
	public String getFormulaValue() {

		return formulaValue;
	}

	/**
	 * セルの実際の式値を設定します。
	 * @param formulaValue セルの実際の式値
	 */
	public void setFormulaValue(String formulaValue) {

		this.formulaValue = formulaValue;
		this.formula = true;
	}

	/**
	 * セルに対するコメントを取得します。
	 * @return セルに対するコメント
	 */
	public String getComment() {

		return comment;
	}

	/**
	 * セルに対するコメントを設定します。
	 * @param comment セルに対するコメント
	 */
	public void setComment(String comment) {

		this.comment = comment;
	}

	/**
	 * 式が埋め込まれているかどうかを取得します。
	 * @return 式が埋め込まれているかどうか
	 */
	public boolean isFormula() {

		return formula;
	}

	/**
	 * 式が埋め込まれているかどうかを設定します。
	 * @param formula 式が埋め込まれているかどうか
	 */
	public void setFormula(boolean formula) {

		this.formula = formula;
	}

	/**
	 * エラーがあるかを取得します。
	 * @return エラーがあるか
	 */
	public boolean isError() {

		return error;
	}

	/**
	 * エラーがあるかを設定します。
	 * @param error エラーがあるか
	 */
	public void setError(boolean error) {

		this.error = error;
	}

	/**
	 * セルのフォントを取得します。
	 * @return セルのフォント
	 */
	public CellFont getFont() {

		return font;
	}

	/**
	 * セルのフォントを設定します。
	 * @param font セルのフォント
	 */
	public void setFont(CellFont font) {

		this.font = font;
	}

	/**
	 * セルのスタイルを取得します。
	 * @return セルのスタイル
	 */
	public CellStyle getStyle() {

		return style;
	}

	/**
	 * セルのスタイルを設定します。
	 * @param style セルのスタイル
	 */
	public void setStyle(CellStyle style) {

		this.style = style;
	}

	/**
	 * 編集ロック状態かどうかを取得します。
	 * @return 編集ロック状態かどうか
	 */
	public boolean isLocked() {

		return locked;
	}

	/**
	 * 編集ロック状態かどうかを設定します。
	 * @param locked 編集ロック状態かどうか
	 */
	public void setLocked(boolean locked) {

		this.locked = locked;
	}

	/**
	 * 結合セル情報を取得します。
	 * @return 結合セル情報
	 */
	public CellRange getMergedCell() {

		return mergedCell;
	}

	/**
	 * 結合セル情報を設定します。
	 * @param mergedCell 結合セル情報
	 */
	public void setMergedCell(CellRange mergedCell) {

		this.mergedCell = mergedCell;
	}

	/**
	 * 結合セルの場合の結合先頭セルを取得します。
	 * @return 結合セルの場合の結合先頭セル
	 */
	public Cell getParentCell() {

		return parentCell;
	}

	/**
	 * 結合セルの場合の結合先頭セルを設定します。
	 * @param parentCell 結合セルの場合の結合先頭セル
	 */
	public void setParentCell(Cell parentCell) {

		this.parentCell = parentCell;
	}

	/**
	 * 編集されたかどうかを取得します。
	 * @return 編集されたかどうか
	 */
	public boolean isEdited() {

		return edited;
	}

	/**
	 * 編集されたかどうかを設定します。
	 * @param edited 編集されたかどうか
	 */
	public void setEdited(boolean edited) {

		this.edited = edited;
	}

	/**
	 * セル位置を取得します。
	 * @return セル位置
	 */
	public Address getAddress() {

		return address;
	}

	/**
	 * セル位置を設定します。
	 * @param address セル位置
	 */
	public void setAddress(Address address) {

		this.address = address;
	}
}
