package com.kiruah.poi2cc.storage.sub;

import java.io.Serializable;

import com.kiruah.poi2cc.Poi2ccUtil;

/**
 * セルのフォント
 *
 * @author Kiruah
 */
public class CellFont implements Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = -2796470750661779102L;

	/** ボールド */
	protected boolean fontBold = false;

	/** 色 */
	protected short fontColor = 0;

	/** フォントサイズ */
	protected short fontSize = 0;

	/** フォント名 */
	protected String fontName = null;

	/** イタリック */
	protected boolean fontItalic = false;

	/** 取り消し線 */
	protected boolean fontStrikeout = false;

	/** 下線 */
	protected byte fontUnderline = 0;

	/**
	 * ボールドを取得します。
	 *
	 * @return ボールド
	 */
	public boolean getFontBold() {

		return fontBold;
	}

	/**
	 * ボールドを設定します。
	 *
	 * @param fontBold ボールド
	 */
	public void setFontBold(boolean fontBold) {

		this.fontBold = fontBold;
	}

	/**
	 * 色を取得します。
	 *
	 * @return 色
	 */
	public short getFontColor() {

		return fontColor;
	}

	/**
	 * 色を設定します。
	 *
	 * @param fontColor 色
	 */
	public void setFontColor(short fontColor) {

		this.fontColor = fontColor;
	}

	/**
	 * フォントサイズを取得します。
	 *
	 * @return フォントサイズ
	 */
	public short getFontSize() {

		return fontSize;
	}

	/**
	 * フォントサイズを設定します。
	 *
	 * @param fontSize フォントサイズ
	 */
	public void setFontSize(short fontSize) {

		this.fontSize = fontSize;
	}

	/**
	 * フォント名を取得します。
	 *
	 * @return フォント名
	 */
	public String getFontName() {

		return fontName;
	}

	/**
	 * フォント名を設定します。
	 *
	 * @param fontName フォント名
	 */
	public void setFontName(String fontName) {

		this.fontName = fontName;
	}

	/**
	 * イタリックを取得します。
	 *
	 * @return イタリック
	 */
	public boolean isFontItalic() {

		return fontItalic;
	}

	/**
	 * イタリックを設定します。
	 *
	 * @param fontItalic イタリック
	 */
	public void setFontItalic(boolean fontItalic) {

		this.fontItalic = fontItalic;
	}

	/**
	 * 取り消し線を取得します。
	 *
	 * @return 取り消し線
	 */
	public boolean isFontStrikeout() {

		return fontStrikeout;
	}

	/**
	 * 取り消し線を設定します。
	 *
	 * @param fontStrikeout 取り消し線
	 */
	public void setFontStrikeout(boolean fontStrikeout) {

		this.fontStrikeout = fontStrikeout;
	}

	/**
	 * 下線を取得します。
	 *
	 * @return 下線
	 */
	public byte getFontUnderline() {

		return fontUnderline;
	}

	/**
	 * 下線を設定します。
	 *
	 * @param fontUnderline 下線
	 */
	public void setFontUnderline(byte fontUnderline) {

		this.fontUnderline = fontUnderline;
	}

	/**
	 * インスタンスをクローンします。
	 *
	 * @return クローンされたインスタンス
	 * @see java.lang.Object#clone()
	 */
	@Override
	public CellFont clone() {

		return Poi2ccUtil.cloneObject(this);
	}

	@Override
	public int hashCode() {

		final int prime = 31;
		int result = 1;
		result = prime * result + (fontBold ? 1231 : 1237);
		result = prime * result + fontColor;
		result = prime * result + (fontItalic ? 1231 : 1237);
		result = prime * result + ((fontName == null) ? 0 : fontName.hashCode());
		result = prime * result + fontSize;
		result = prime * result + (fontStrikeout ? 1231 : 1237);
		result = prime * result + fontUnderline;
		return result;
	}

	@Override
	public boolean equals(Object obj) {

		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		CellFont other = (CellFont) obj;
		if (fontBold != other.fontBold)
			return false;
		if (fontColor != other.fontColor)
			return false;
		if (fontItalic != other.fontItalic)
			return false;
		if (fontName == null) {
			if (other.fontName != null)
				return false;
		} else if (!fontName.equals(other.fontName))
			return false;
		if (fontSize != other.fontSize)
			return false;
		if (fontStrikeout != other.fontStrikeout)
			return false;
		if (fontUnderline != other.fontUnderline)
			return false;
		return true;
	}
}
