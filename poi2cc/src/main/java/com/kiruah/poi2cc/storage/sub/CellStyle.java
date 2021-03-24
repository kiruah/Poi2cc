package com.kiruah.poi2cc.storage.sub;

import java.awt.Color;
import java.io.Serializable;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * セルのスタイル情報
 *
 * @author Kiruah
 */
public class CellStyle implements Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = -3813059074469574750L;

	/** 書式 */
	protected String format = null;

	/** 前面色 */
	protected Color foregroundColor = Color.BLACK;

	/** 背景色 */
	protected Color backgroundColor = Color.WHITE;

	/** 網掛パターン */
	protected FillPatternType fillPattern = FillPatternType.NO_FILL;

	/** 横方向整列 */
	protected HorizontalAlignment alignment = HorizontalAlignment.GENERAL;

	/** 縦方向整列 */
	protected VerticalAlignment verticalAlignment = VerticalAlignment.TOP;

	/** インデント */
	protected short indent = 0;

	/** 文字列の改行 */
	protected boolean wrapText = false;

	/** 回転 */
	protected short rotation = 0;

	/** 罫線(下線) */
	protected BorderStyle borderBottom = BorderStyle.NONE;

	/** 罫線(左線) */
	protected BorderStyle borderLeft = BorderStyle.NONE;

	/** 罫線(右線) */
	protected BorderStyle borderRight = BorderStyle.NONE;

	/** 罫線(上線) */
	protected BorderStyle borderTop = BorderStyle.NONE;

	/* (非 Javadoc)
	 * @see java.lang.Object#hashCode()
	 */
	@Override
	public int hashCode() {

		final int prime = 31;
		int result = 1;
		result = prime * result + ((alignment == null) ? 0 : alignment.hashCode());
		result = prime * result + ((backgroundColor == null) ? 0 : backgroundColor.hashCode());
		result = prime * result + ((borderBottom == null) ? 0 : borderBottom.hashCode());
		result = prime * result + ((borderLeft == null) ? 0 : borderLeft.hashCode());
		result = prime * result + ((borderRight == null) ? 0 : borderRight.hashCode());
		result = prime * result + ((borderTop == null) ? 0 : borderTop.hashCode());
		result = prime * result + ((fillPattern == null) ? 0 : fillPattern.hashCode());
		result = prime * result + ((foregroundColor == null) ? 0 : foregroundColor.hashCode());
		result = prime * result + ((format == null) ? 0 : format.hashCode());
		result = prime * result + indent;
		result = prime * result + rotation;
		result = prime * result + ((verticalAlignment == null) ? 0 : verticalAlignment.hashCode());
		result = prime * result + (wrapText ? 1231 : 1237);
		return result;
	}

	/* (非 Javadoc)
	 * @see java.lang.Object#equals(java.lang.Object)
	 */
	@Override
	public boolean equals(Object obj) {

		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		CellStyle other = (CellStyle) obj;
		if (alignment != other.alignment)
			return false;
		if (backgroundColor == null) {
			if (other.backgroundColor != null)
				return false;
		} else if (!backgroundColor.equals(other.backgroundColor))
			return false;
		if (borderBottom != other.borderBottom)
			return false;
		if (borderLeft != other.borderLeft)
			return false;
		if (borderRight != other.borderRight)
			return false;
		if (borderTop != other.borderTop)
			return false;
		if (fillPattern != other.fillPattern)
			return false;
		if (foregroundColor == null) {
			if (other.foregroundColor != null)
				return false;
		} else if (!foregroundColor.equals(other.foregroundColor))
			return false;
		if (format == null) {
			if (other.format != null)
				return false;
		} else if (!format.equals(other.format))
			return false;
		if (indent != other.indent)
			return false;
		if (rotation != other.rotation)
			return false;
		if (verticalAlignment != other.verticalAlignment)
			return false;
		if (wrapText != other.wrapText)
			return false;
		return true;
	}

	/**
	 * 書式を取得します。
	 * @return 書式
	 */
	public String getFormat() {

		return format;
	}

	/**
	 * 書式を設定します。
	 * @param format 書式
	 */
	public void setFormat(String format) {

		this.format = format;
	}

	/**
	 * 前面色を取得します。
	 * @return 前面色
	 */
	public Color getForegroundColor() {

		return foregroundColor;
	}

	/**
	 * 前面色を設定します。
	 * @param foregroundColor 前面色
	 */
	public void setForegroundColor(Color foregroundColor) {

		this.foregroundColor = foregroundColor;
	}

	/**
	 * 背景色を取得します。
	 * @return 背景色
	 */
	public Color getBackgroundColor() {

		return backgroundColor;
	}

	/**
	 * 背景色を設定します。
	 * @param backgroundColor 背景色
	 */
	public void setBackgroundColor(Color backgroundColor) {

		this.backgroundColor = backgroundColor;
	}

	/**
	 * 網掛パターンを取得します。
	 * @return 網掛パターン
	 */
	public FillPatternType getFillPattern() {

		return fillPattern;
	}

	/**
	 * 網掛パターンを設定します。
	 * @param fillPattern 網掛パターン
	 */
	public void setFillPattern(FillPatternType fillPattern) {

		this.fillPattern = fillPattern;
	}

	/**
	 * 横方向整列を取得します。
	 * @return 横方向整列
	 */
	public HorizontalAlignment getAlignment() {

		return alignment;
	}

	/**
	 * 横方向整列を設定します。
	 * @param alignment 横方向整列
	 */
	public void setAlignment(HorizontalAlignment alignment) {

		this.alignment = alignment;
	}

	/**
	 * 縦方向整列を取得します。
	 * @return 縦方向整列
	 */
	public VerticalAlignment getVerticalAlignment() {

		return verticalAlignment;
	}

	/**
	 * 縦方向整列を設定します。
	 * @param verticalAlignment 縦方向整列
	 */
	public void setVerticalAlignment(VerticalAlignment verticalAlignment) {

		this.verticalAlignment = verticalAlignment;
	}

	/**
	 * インデントを取得します。
	 * @return インデント
	 */
	public short getIndent() {

		return indent;
	}

	/**
	 * インデントを設定します。
	 * @param indent インデント
	 */
	public void setIndent(short indent) {

		this.indent = indent;
	}

	/**
	 * 文字列の改行を取得します。
	 * @return 文字列の改行
	 */
	public boolean isWrapText() {

		return wrapText;
	}

	/**
	 * 文字列の改行を設定します。
	 * @param wrapText 文字列の改行
	 */
	public void setWrapText(boolean wrapText) {

		this.wrapText = wrapText;
	}

	/**
	 * 回転を取得します。
	 * @return 回転
	 */
	public short getRotation() {

		return rotation;
	}

	/**
	 * 回転を設定します。
	 * @param rotation 回転
	 */
	public void setRotation(short rotation) {

		this.rotation = rotation;
	}

	/**
	 * 罫線(下線)を取得します。
	 * @return 罫線(下線)
	 */
	public BorderStyle getBorderBottom() {

		return borderBottom;
	}

	/**
	 * 罫線(下線)を設定します。
	 * @param borderBottom 罫線(下線)
	 */
	public void setBorderBottom(BorderStyle borderBottom) {

		this.borderBottom = borderBottom;
	}

	/**
	 * 罫線(左線)を取得します。
	 * @return 罫線(左線)
	 */
	public BorderStyle getBorderLeft() {

		return borderLeft;
	}

	/**
	 * 罫線(左線)を設定します。
	 * @param borderLeft 罫線(左線)
	 */
	public void setBorderLeft(BorderStyle borderLeft) {

		this.borderLeft = borderLeft;
	}

	/**
	 * 罫線(右線)を取得します。
	 * @return 罫線(右線)
	 */
	public BorderStyle getBorderRight() {

		return borderRight;
	}

	/**
	 * 罫線(右線)を設定します。
	 * @param borderRight 罫線(右線)
	 */
	public void setBorderRight(BorderStyle borderRight) {

		this.borderRight = borderRight;
	}

	/**
	 * 罫線(上線)を取得します。
	 * @return 罫線(上線)
	 */
	public BorderStyle getBorderTop() {

		return borderTop;
	}

	/**
	 * 罫線(上線)を設定します。
	 * @param borderTop 罫線(上線)
	 */
	public void setBorderTop(BorderStyle borderTop) {

		this.borderTop = borderTop;
	}
}
