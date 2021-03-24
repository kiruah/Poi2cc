package com.kiruah.poi2cc.storage.sub;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.kiruah.poi2cc.Poi2ccConstants;
import com.kiruah.poi2cc.Poi2ccUtil;

/**
 * シートのスタイル・印刷スタイル情報
 *
 * @author Kiruah
 */
public class SheetStyle
		implements
			Serializable, Cloneable {

	/** シリアルバージョンUID */
	private static final long serialVersionUID = -8950049424594794809L;

	/** グリッド線の表示 */
	protected boolean displayGridLine = false;

	/** ゼロ値の表示 */
	protected boolean displayZeros = false;

	/** 自動改ページ */
	protected boolean autoBreaks = false;

	/** GUTSの表示 */
	protected boolean displayGuts = false;

	/** ページフィット */
	protected boolean fitToPage = false;

	/** 水平方向センタリング */
	protected boolean horizontallyCenter = false;

	/** 垂直方向センタリング */
	protected boolean verticallyCenter = false;

	/** グリッド線の印刷 */
	protected boolean printGridlines = false;

	/** ヘッダーの有無 */
	protected boolean header = false;

	/** ヘッダー文字列左部 */
	protected String headerLeft = Poi2ccConstants.EMPTY_STRING;

	/** ヘッダー文字列中央部 */
	protected String headerCenter = Poi2ccConstants.EMPTY_STRING;

	/** ヘッダー文字列右部 */
	protected String headerRight = Poi2ccConstants.EMPTY_STRING;

	/** フッターの有無 */
	protected boolean footer = false;

	/** フッター文字列左部 */
	protected String footerLeft = Poi2ccConstants.EMPTY_STRING;

	/** フッター文字列中央部 */
	protected String footerCenter = Poi2ccConstants.EMPTY_STRING;

	/** フッター文字列右部 */
	protected String footerRight = Poi2ccConstants.EMPTY_STRING;

	/** 列の改ページ */
	protected List<Integer> columnBreakList = new ArrayList<Integer>();

	/** 行の改ページ */
	protected List<Integer> rowBreakList = new ArrayList<Integer>();

	/** 印刷設定済 */
	protected boolean printerSetup = false;

	/** フィットさせる高さ */
	protected int fitHeight = 0;

	/** フィットさせる横幅 */
	protected int fitWidth = 0;

	/** ドラフト */
	protected boolean draft = false;

	/** フッターマージン */
	protected double footerMargin = 0.0;

	/** ヘッダーマージン */
	protected double headerMargin = 0.0;

	/** 水平方向 解像度 */
	protected int horizontallyResolution = 0;

	/** 垂直方向 解像度 */
	protected int verticallyResolution = 0;

	/** ランドスケープ */
	protected boolean landscape = false;

	/** 左から右表示 */
	protected boolean leftToRight = false;

	/** 回転有無 */
	protected boolean noOrientation = false;

	/** 白黒印刷 */
	protected boolean noColor = false;

	/** 備考 */
	protected boolean notes = false;

	/** 印刷用紙サイズ */
	protected int paperSize = 0;

	/** 印刷開始ページ */
	protected int pageStart = 0;

	/** 印刷サイズ */
	protected int scale = 0;

	/** ページ設定の利用 */
	protected boolean usePage = false;

	/** 設定の妥当性 */
	protected boolean validSettings = false;

	/** ページ上部マージン */
	protected double marginTop = 0.0;

	/** ページ下部マージン */
	protected double marginBottom = 0.0;

	/** ページ左部マージン */
	protected double marginLeft = 0.0;

	/** ページ右部マージン */
	protected double marginRight = 0.0;

	/**
	 * グリッド線の表示を取得します。
	 * @return グリッド線の表示
	 */
	public boolean isDisplayGridLine() {

		return displayGridLine;
	}

	/**
	 * グリッド線の表示を設定します。
	 * @param displayGridLine グリッド線の表示
	 */
	public void setDisplayGridLine(
			boolean displayGridLine) {

		this.displayGridLine = displayGridLine;
	}

	/**
	 * ゼロ値の表示を取得します。
	 * @return ゼロ値の表示
	 */
	public boolean isDisplayZeros() {

		return displayZeros;
	}

	/**
	 * ゼロ値の表示を設定します。
	 * @param displayZeros ゼロ値の表示
	 */
	public void setDisplayZeros(
			boolean displayZeros) {

		this.displayZeros = displayZeros;
	}

	/**
	 * 自動改ページを取得します。
	 * @return 自動改ページ
	 */
	public boolean isAutoBreaks() {

		return autoBreaks;
	}

	/**
	 * 自動改ページを設定します。
	 * @param autoBreaks 自動改ページ
	 */
	public void setAutoBreaks(
			boolean autoBreaks) {

		this.autoBreaks = autoBreaks;
	}

	/**
	 * GUTSの表示を取得します。
	 * @return GUTSの表示
	 */
	public boolean isDisplayGuts() {

		return displayGuts;
	}

	/**
	 * GUTSの表示を設定します。
	 * @param displayGuts GUTSの表示
	 */
	public void setDisplayGuts(
			boolean displayGuts) {

		this.displayGuts = displayGuts;
	}

	/**
	 * ページフィットを取得します。
	 * @return ページフィット
	 */
	public boolean isFitToPage() {

		return fitToPage;
	}

	/**
	 * ページフィットを設定します。
	 * @param fitToPage ページフィット
	 */
	public void setFitToPage(
			boolean fitToPage) {

		this.fitToPage = fitToPage;
	}

	/**
	 * 水平方向センタリングを取得します。
	 * @return 水平方向センタリング
	 */
	public boolean isHorizontallyCenter() {

		return horizontallyCenter;
	}

	/**
	 * 水平方向センタリングを設定します。
	 * @param horizontallyCenter 水平方向センタリング
	 */
	public void setHorizontallyCenter(
			boolean horizontallyCenter) {

		this.horizontallyCenter = horizontallyCenter;
	}

	/**
	 * 垂直方向センタリングを取得します。
	 * @return 垂直方向センタリング
	 */
	public boolean isVerticallyCenter() {

		return verticallyCenter;
	}

	/**
	 * 垂直方向センタリングを設定します。
	 * @param verticallyCenter 垂直方向センタリング
	 */
	public void setVerticallyCenter(
			boolean verticallyCenter) {

		this.verticallyCenter = verticallyCenter;
	}

	/**
	 * グリッド線の印刷を取得します。
	 * @return グリッド線の印刷
	 */
	public boolean isPrintGridlines() {

		return printGridlines;
	}

	/**
	 * グリッド線の印刷を設定します。
	 * @param printGridlines グリッド線の印刷
	 */
	public void setPrintGridlines(
			boolean printGridlines) {

		this.printGridlines = printGridlines;
	}

	/**
	 * ヘッダーの有無を取得します。
	 * @return ヘッダーの有無
	 */
	public boolean isHeader() {

		return header;
	}

	/**
	 * ヘッダーの有無を設定します。
	 * @param header ヘッダーの有無
	 */
	public void setHeader(
			boolean header) {

		this.header = header;
	}

	/**
	 * ヘッダー文字列左部を取得します。
	 * @return ヘッダー文字列左部
	 */
	public String getHeaderLeft() {

		return headerLeft;
	}

	/**
	 * ヘッダー文字列左部を設定します。
	 * @param headerLeft ヘッダー文字列左部
	 */
	public void setHeaderLeft(
			String headerLeft) {

		this.headerLeft = headerLeft;
		this.header = true;
	}

	/**
	 * ヘッダー文字列中央部を取得します。
	 * @return ヘッダー文字列中央部
	 */
	public String getHeaderCenter() {

		return headerCenter;
	}

	/**
	 * ヘッダー文字列中央部を設定します。
	 * @param headerCenter ヘッダー文字列中央部
	 */
	public void setHeaderCenter(
			String headerCenter) {

		this.headerCenter = headerCenter;
		this.header = true;
	}

	/**
	 * ヘッダー文字列右部を取得します。
	 * @return ヘッダー文字列右部
	 */
	public String getHeaderRight() {

		return headerRight;
	}

	/**
	 * ヘッダー文字列右部を設定します。
	 * @param headerRight ヘッダー文字列右部
	 */
	public void setHeaderRight(
			String headerRight) {

		this.headerRight = headerRight;
		this.header = true;
	}

	/**
	 * フッターの有無を取得します。
	 * @return フッターの有無
	 */
	public boolean isFooter() {

		return footer;
	}

	/**
	 * フッターの有無を設定します。
	 * @param footer フッターの有無
	 */
	public void setFooter(
			boolean footer) {

		this.footer = footer;
	}

	/**
	 * フッター文字列左部を取得します。
	 * @return フッター文字列左部
	 */
	public String getFooterLeft() {

		return footerLeft;
	}

	/**
	 * フッター文字列左部を設定します。
	 * @param footerLeft フッター文字列左部
	 */
	public void setFooterLeft(
			String footerLeft) {

		this.footerLeft = footerLeft;
		this.footer = true;
	}

	/**
	 * フッター文字列中央部を取得します。
	 * @return フッター文字列中央部
	 */
	public String getFooterCenter() {

		return footerCenter;
	}

	/**
	 * フッター文字列中央部を設定します。
	 * @param footerCenter フッター文字列中央部
	 */
	public void setFooterCenter(
			String footerCenter) {

		this.footerCenter = footerCenter;
		this.footer = true;
	}

	/**
	 * フッター文字列右部を取得します。
	 * @return フッター文字列右部
	 */
	public String getFooterRight() {

		return footerRight;
	}

	/**
	 * フッター文字列右部を設定します。
	 * @param footerRight フッター文字列右部
	 */
	public void setFooterRight(
			String footerRight) {

		this.footerRight = footerRight;
		this.footer = true;
	}

	/**
	 * 列の改ページを取得します。
	 * @return 列の改ページ
	 */
	public List<Integer> getColumnBreakList() {

		return columnBreakList;
	}

	/**
	 * 列の改ページを設定します。
	 * @param columnBreakList 列の改ページ
	 */
	public void setColumnBreakList(
			List<Integer> columnBreakList) {

		this.columnBreakList = columnBreakList;
	}

	/**
	 * 行の改ページを取得します。
	 * @return 行の改ページ
	 */
	public List<Integer> getRowBreakList() {

		return rowBreakList;
	}

	/**
	 * 行の改ページを設定します。
	 * @param rowBreakList 行の改ページ
	 */
	public void setRowBreakList(
			List<Integer> rowBreakList) {

		this.rowBreakList = rowBreakList;
	}

	/**
	 * 印刷設定済を取得します。
	 * @return 印刷設定済
	 */
	public boolean isPrinterSetup() {

		return printerSetup;
	}

	/**
	 * 印刷設定済を設定します。
	 * @param printerSetup 印刷設定済
	 */
	public void setPrinterSetup(
			boolean printerSetup) {

		this.printerSetup = printerSetup;
	}

	/**
	 * フィットさせる高さを取得します。
	 * @return フィットさせる高さ
	 */
	public int getFitHeight() {

		return fitHeight;
	}

	/**
	 * フィットさせる高さを設定します。
	 * @param fitHeight フィットさせる高さ
	 */
	public void setFitHeight(
			int fitHeight) {

		this.fitHeight = fitHeight;
		this.printerSetup = true;
	}

	/**
	 * フィットさせる横幅を取得します。
	 * @return フィットさせる横幅
	 */
	public int getFitWidth() {

		return fitWidth;
	}

	/**
	 * フィットさせる横幅を設定します。
	 * @param fitWidth フィットさせる横幅
	 */
	public void setFitWidth(
			int fitWidth) {

		this.fitWidth = fitWidth;
		this.printerSetup = true;
	}

	/**
	 * ドラフトを取得します。
	 * @return ドラフト
	 */
	public boolean isDraft() {

		return draft;
	}

	/**
	 * ドラフトを設定します。
	 * @param draft ドラフト
	 */
	public void setDraft(
			boolean draft) {

		this.draft = draft;
		this.printerSetup = true;
	}

	/**
	 * フッターマージンを取得します。
	 * @return フッターマージン
	 */
	public double getFooterMargin() {

		return footerMargin;
	}

	/**
	 * フッターマージンを設定します。
	 * @param footerMargin フッターマージン
	 */
	public void setFooterMargin(
			double footerMargin) {

		this.footerMargin = footerMargin;
		this.printerSetup = true;
	}

	/**
	 * ヘッダーマージンを取得します。
	 * @return ヘッダーマージン
	 */
	public double getHeaderMargin() {

		return headerMargin;
	}

	/**
	 * ヘッダーマージンを設定します。
	 * @param headerMargin ヘッダーマージン
	 */
	public void setHeaderMargin(
			double headerMargin) {

		this.headerMargin = headerMargin;
		this.printerSetup = true;
	}

	/**
	 * 水平方向 解像度を取得します。
	 * @return 水平方向 解像度
	 */
	public int getHorizontallyResolution() {

		return horizontallyResolution;
	}

	/**
	 * 水平方向 解像度を設定します。
	 * @param horizontallyResolution 水平方向 解像度
	 */
	public void setHorizontallyResolution(
			int horizontallyResolution) {

		this.horizontallyResolution = horizontallyResolution;
		this.printerSetup = true;
	}

	/**
	 * 垂直方向 解像度を取得します。
	 * @return 垂直方向 解像度
	 */
	public int getVerticallyResolution() {

		return verticallyResolution;
	}

	/**
	 * 垂直方向 解像度を設定します。
	 * @param verticallyResolution 垂直方向 解像度
	 */
	public void setVerticallyResolution(
			int verticallyResolution) {

		this.verticallyResolution = verticallyResolution;
		this.printerSetup = true;
	}

	/**
	 * ランドスケープを取得します。
	 * @return ランドスケープ
	 */
	public boolean isLandscape() {

		return landscape;
	}

	/**
	 * ランドスケープを設定します。
	 * @param landscape ランドスケープ
	 */
	public void setLandscape(
			boolean landscape) {

		this.landscape = landscape;
		this.printerSetup = true;
	}

	/**
	 * 左から右表示を取得します。
	 * @return 左から右表示
	 */
	public boolean isLeftToRight() {

		return leftToRight;
	}

	/**
	 * 左から右表示を設定します。
	 * @param leftToRight 左から右表示
	 */
	public void setLeftToRight(
			boolean leftToRight) {

		this.leftToRight = leftToRight;
		this.printerSetup = true;
	}

	/**
	 * 回転有無を取得します。
	 * @return 回転有無
	 */
	public boolean isNoOrientation() {

		return noOrientation;
	}

	/**
	 * 回転有無を設定します。
	 * @param noOrientation 回転有無
	 */
	public void setNoOrientation(
			boolean noOrientation) {

		this.noOrientation = noOrientation;
		this.printerSetup = true;
	}

	/**
	 * 白黒印刷を取得します。
	 * @return 白黒印刷
	 */
	public boolean isNoColor() {

		return noColor;
	}

	/**
	 * 白黒印刷を設定します。
	 * @param noColor 白黒印刷
	 */
	public void setNoColor(
			boolean noColor) {

		this.noColor = noColor;
		this.printerSetup = true;
	}

	/**
	 * 備考を取得します。
	 * @return 備考
	 */
	public boolean isNotes() {

		return notes;
	}

	/**
	 * 備考を設定します。
	 * @param notes 備考
	 */
	public void setNotes(
			boolean notes) {

		this.notes = notes;
		this.printerSetup = true;
	}

	/**
	 * 印刷用紙サイズを取得します。
	 * @return 印刷用紙サイズ
	 */
	public int getPaperSize() {

		return paperSize;
	}

	/**
	 * 印刷用紙サイズを設定します。
	 * @param paperSize 印刷用紙サイズ
	 */
	public void setPaperSize(
			int paperSize) {

		this.paperSize = paperSize;
		this.printerSetup = true;
	}

	/**
	 * 印刷開始ページを取得します。
	 * @return 印刷開始ページ
	 */
	public int getPageStart() {

		return pageStart;
	}

	/**
	 * 印刷開始ページを設定します。
	 * @param pageStart 印刷開始ページ
	 */
	public void setPageStart(
			int pageStart) {

		this.pageStart = pageStart;
		this.printerSetup = true;
	}

	/**
	 * 印刷サイズを取得します。
	 * @return 印刷サイズ
	 */
	public int getScale() {

		return scale;
	}

	/**
	 * 印刷サイズを設定します。
	 * @param scale 印刷サイズ
	 */
	public void setScale(
			int scale) {

		this.scale = scale;
		this.printerSetup = true;
	}

	/**
	 * ページ設定の利用を取得します。
	 * @return ページ設定の利用
	 */
	public boolean isUsePage() {

		return usePage;
	}

	/**
	 * ページ設定の利用を設定します。
	 * @param usePage ページ設定の利用
	 */
	public void setUsePage(
			boolean usePage) {

		this.usePage = usePage;
		this.printerSetup = true;
	}

	/**
	 * 設定の妥当性を取得します。
	 * @return 設定の妥当性
	 */
	public boolean isValidSettings() {

		return validSettings;
	}

	/**
	 * 設定の妥当性を設定します。
	 * @param validSettings 設定の妥当性
	 */
	public void setValidSettings(
			boolean validSettings) {

		this.validSettings = validSettings;
		this.printerSetup = true;
	}

	/**
	 * ページ上部マージンを取得します。
	 * @return ページ上部マージン
	 */
	public double getMarginTop() {

		return marginTop;
	}

	/**
	 * ページ上部マージンを設定します。
	 * @param marginTop ページ上部マージン
	 */
	public void setMarginTop(
			double marginTop) {

		this.marginTop = marginTop;
	}

	/**
	 * ページ下部マージンを取得します。
	 * @return ページ下部マージン
	 */
	public double getMarginBottom() {

		return marginBottom;
	}

	/**
	 * ページ下部マージンを設定します。
	 * @param marginBottom ページ下部マージン
	 */
	public void setMarginBottom(
			double marginBottom) {

		this.marginBottom = marginBottom;
	}

	/**
	 * ページ左部マージンを取得します。
	 * @return ページ左部マージン
	 */
	public double getMarginLeft() {

		return marginLeft;
	}

	/**
	 * ページ左部マージンを設定します。
	 * @param marginLeft ページ左部マージン
	 */
	public void setMarginLeft(
			double marginLeft) {

		this.marginLeft = marginLeft;
	}

	/**
	 * ページ右部マージンを取得します。
	 * @return ページ右部マージン
	 */
	public double getMarginRight() {

		return marginRight;
	}

	/**
	 * ページ右部マージンを設定します。
	 * @param marginRight ページ右部マージン
	 */
	public void setMarginRight(
			double marginRight) {

		this.marginRight = marginRight;
	}

	/**
	 * インスタンスをクローンします。
	 *
	 * @return クローンされたインスタンス
	 * @see java.lang.Object#clone()
	 */
	@Override
	public SheetStyle clone() {

		return Poi2ccUtil.cloneObject(this);
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
		result = prime * result + (autoBreaks ? 1231 : 1237);
		result = prime * result + ((columnBreakList == null) ? 0 : columnBreakList.hashCode());
		result = prime * result + (displayGridLine ? 1231 : 1237);
		result = prime * result + (displayGuts ? 1231 : 1237);
		result = prime * result + (displayZeros ? 1231 : 1237);
		result = prime * result + (draft ? 1231 : 1237);
		result = prime * result + fitHeight;
		result = prime * result + (fitToPage ? 1231 : 1237);
		result = prime * result + fitWidth;
		result = prime * result + (footer ? 1231 : 1237);
		result = prime * result + ((footerCenter == null) ? 0 : footerCenter.hashCode());
		result = prime * result + ((footerLeft == null) ? 0 : footerLeft.hashCode());
		long temp;
		temp = Double.doubleToLongBits(footerMargin);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		result = prime * result + ((footerRight == null) ? 0 : footerRight.hashCode());
		result = prime * result + (header ? 1231 : 1237);
		result = prime * result + ((headerCenter == null) ? 0 : headerCenter.hashCode());
		result = prime * result + ((headerLeft == null) ? 0 : headerLeft.hashCode());
		temp = Double.doubleToLongBits(headerMargin);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		result = prime * result + ((headerRight == null) ? 0 : headerRight.hashCode());
		result = prime * result + (horizontallyCenter ? 1231 : 1237);
		result = prime * result + horizontallyResolution;
		result = prime * result + (landscape ? 1231 : 1237);
		result = prime * result + (leftToRight ? 1231 : 1237);
		temp = Double.doubleToLongBits(marginBottom);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		temp = Double.doubleToLongBits(marginLeft);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		temp = Double.doubleToLongBits(marginRight);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		temp = Double.doubleToLongBits(marginTop);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		result = prime * result + (noColor ? 1231 : 1237);
		result = prime * result + (noOrientation ? 1231 : 1237);
		result = prime * result + (notes ? 1231 : 1237);
		result = prime * result + pageStart;
		result = prime * result + paperSize;
		result = prime * result + (printGridlines ? 1231 : 1237);
		result = prime * result + (printerSetup ? 1231 : 1237);
		result = prime * result + ((rowBreakList == null) ? 0 : rowBreakList.hashCode());
		result = prime * result + scale;
		result = prime * result + (usePage ? 1231 : 1237);
		result = prime * result + (validSettings ? 1231 : 1237);
		result = prime * result + (verticallyCenter ? 1231 : 1237);
		result = prime * result + verticallyResolution;
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
		SheetStyle other = (SheetStyle) obj;
		if (autoBreaks != other.autoBreaks) {
			return false;
		}
		if (columnBreakList == null) {
			if (other.columnBreakList != null) {
				return false;
			}
		} else if (!columnBreakList.equals(other.columnBreakList)) {
			return false;
		}
		if (displayGridLine != other.displayGridLine) {
			return false;
		}
		if (displayGuts != other.displayGuts) {
			return false;
		}
		if (displayZeros != other.displayZeros) {
			return false;
		}
		if (draft != other.draft) {
			return false;
		}
		if (fitHeight != other.fitHeight) {
			return false;
		}
		if (fitToPage != other.fitToPage) {
			return false;
		}
		if (fitWidth != other.fitWidth) {
			return false;
		}
		if (footer != other.footer) {
			return false;
		}
		if (footerCenter == null) {
			if (other.footerCenter != null) {
				return false;
			}
		} else if (!footerCenter.equals(other.footerCenter)) {
			return false;
		}
		if (footerLeft == null) {
			if (other.footerLeft != null) {
				return false;
			}
		} else if (!footerLeft.equals(other.footerLeft)) {
			return false;
		}
		if (Double.doubleToLongBits(footerMargin) != Double.doubleToLongBits(other.footerMargin)) {
			return false;
		}
		if (footerRight == null) {
			if (other.footerRight != null) {
				return false;
			}
		} else if (!footerRight.equals(other.footerRight)) {
			return false;
		}
		if (header != other.header) {
			return false;
		}
		if (headerCenter == null) {
			if (other.headerCenter != null) {
				return false;
			}
		} else if (!headerCenter.equals(other.headerCenter)) {
			return false;
		}
		if (headerLeft == null) {
			if (other.headerLeft != null) {
				return false;
			}
		} else if (!headerLeft.equals(other.headerLeft)) {
			return false;
		}
		if (Double.doubleToLongBits(headerMargin) != Double.doubleToLongBits(other.headerMargin)) {
			return false;
		}
		if (headerRight == null) {
			if (other.headerRight != null) {
				return false;
			}
		} else if (!headerRight.equals(other.headerRight)) {
			return false;
		}
		if (horizontallyCenter != other.horizontallyCenter) {
			return false;
		}
		if (horizontallyResolution != other.horizontallyResolution) {
			return false;
		}
		if (landscape != other.landscape) {
			return false;
		}
		if (leftToRight != other.leftToRight) {
			return false;
		}
		if (Double.doubleToLongBits(marginBottom) != Double.doubleToLongBits(other.marginBottom)) {
			return false;
		}
		if (Double.doubleToLongBits(marginLeft) != Double.doubleToLongBits(other.marginLeft)) {
			return false;
		}
		if (Double.doubleToLongBits(marginRight) != Double.doubleToLongBits(other.marginRight)) {
			return false;
		}
		if (Double.doubleToLongBits(marginTop) != Double.doubleToLongBits(other.marginTop)) {
			return false;
		}
		if (noColor != other.noColor) {
			return false;
		}
		if (noOrientation != other.noOrientation) {
			return false;
		}
		if (notes != other.notes) {
			return false;
		}
		if (pageStart != other.pageStart) {
			return false;
		}
		if (paperSize != other.paperSize) {
			return false;
		}
		if (printGridlines != other.printGridlines) {
			return false;
		}
		if (printerSetup != other.printerSetup) {
			return false;
		}
		if (rowBreakList == null) {
			if (other.rowBreakList != null) {
				return false;
			}
		} else if (!rowBreakList.equals(other.rowBreakList)) {
			return false;
		}
		if (scale != other.scale) {
			return false;
		}
		if (usePage != other.usePage) {
			return false;
		}
		if (validSettings != other.validSettings) {
			return false;
		}
		if (verticallyCenter != other.verticallyCenter) {
			return false;
		}
		if (verticallyResolution != other.verticallyResolution) {
			return false;
		}
		return true;
	}
}
