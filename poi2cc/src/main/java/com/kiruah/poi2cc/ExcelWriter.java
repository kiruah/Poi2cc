package com.kiruah.poi2cc;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.kiruah.poi2cc.storage.Address;
import com.kiruah.poi2cc.storage.Book;
import com.kiruah.poi2cc.storage.Cell;
import com.kiruah.poi2cc.storage.Sheet;
import com.kiruah.poi2cc.storage.ValueType;
import com.kiruah.poi2cc.storage.sub.CellFont;
import com.kiruah.poi2cc.storage.sub.CellRange;
import com.kiruah.poi2cc.storage.sub.SheetStyle;

/**
 * Excelファイルライター
 *
 * @author Kiruah
 */
public class ExcelWriter {

	/** カラムサイズデフォルト値補正数 */
	protected static final int COLUMN_CORRECTING_VALUE = 256;

	/** 最大カラム数 */
	protected static final int MAX_COLUMN_SIZE = 255;

	/** 現在処理中のブック */
	protected XSSFWorkbook fileBook = null;

	/** 現在処理中のデータフォーマット */
	protected XSSFDataFormat fileDataFormat = null;

	/** 現在処理中のシート */
	protected XSSFSheet fileSheet = null;

	/** 現在処理中の行 */
	protected XSSFRow row = null;

	/** 現在処理中のセル */
	protected XSSFCell fileCell = null;

	/** 現在処理中のすべてのフォントキャッシュ */
	protected Map<CellFont, XSSFFont> fontCache = new HashMap<CellFont, XSSFFont>();

	/** 現在処理中のすべてのセルスタイルキャッシュ */
	protected Map<com.kiruah.poi2cc.storage.sub.CellStyle, XSSFCellStyle> styleCache = new HashMap<com.kiruah.poi2cc.storage.sub.CellStyle, XSSFCellStyle>();

	/**
	 * Excelブック情報を指定したファイル名で2003以前の形式で保存します。
	 *
	 * @param book ブック
	 * @param file ファイル名
	 * @throws IOException IO例外
	 */
	public static void save(Book book, File file) throws IOException {

		String fileName = null;

		if (file != null) {
			fileName = file.getAbsolutePath();
		} else {
			fileName = book.getName();
		}

		save(book, fileName);
	}

	/**
	 * Excelブック情報を指定したファイル名で保存します。
	 *
	 * @param book ブック
	 * @param fileName ファイル名
	 * @throws IOException IO例外
	 */
	public static void save(Book book, String fileName) throws IOException {

		ExcelWriter writer = new ExcelWriter();
		XSSFWorkbook workBook = new XSSFWorkbook();

		if (fileName == null) {
			fileName = book.getName();
		}

		writer.saveBook(book, workBook, fileName);
	}

	/**
	 * Excelブック情報を指定したApache POIワークブック形式で保存します。
	 *
	 * @param book ブック
	 * @param workBook 保存するワークブック形式
	 * @param fileName ファイル名
	 * @throws IOException IO例外
	 */
	public void saveBook(Book book, XSSFWorkbook workBook, String fileName) throws IOException {

		if (book == null) {
			throw new Poi2ccRuntimeException("保存対象のブックがnullのため保存できません");
		}
		if (Poi2ccUtil.isEmptyString(fileName) == true) {
			throw new Poi2ccRuntimeException("ファイル名が指定されていません");
		}
		if (book.getSheetList().size() == 0) {
			throw new Poi2ccRuntimeException("保存するシートが存在しないため、保存処理を継続できません : " + fileName);
		}

		fileBook = workBook;
		fileDataFormat = fileBook.createDataFormat();

		createBook(book);

		output(fileName);
	}

	/**
	 * 生成したメモリ上のExcelブック情報をファイルに出力します。
	 *
	 * @param fileName ファイル名
	 * @throws IOException IO例外
	 */
	protected void output(String fileName) throws IOException {

		FileOutputStream out = null;

		try {
			out = new FileOutputStream(fileName);

			fileBook.write(out);
		} finally {
			Poi2ccUtil.close(out);
		}
	}

	/**
	 * ブックを作成します。
	 *
	 * @param book ブック
	 */
	protected void createBook(Book book) {

		for (Sheet sheet : book.getSheetList()) {
			createSheet(sheet);
		}
	}

	/**
	 * シートのスタイル情報を生成します。
	 *
	 * @param style スタイル情報
	 */
	protected void createSheetStyle(SheetStyle style) {

		fileSheet.setDisplayGridlines(style.isDisplayGridLine());
		fileSheet.setDisplayZeros(style.isDisplayZeros());
		fileSheet.setMargin(org.apache.poi.ss.usermodel.Sheet.TopMargin, style.getMarginTop());
		fileSheet.setMargin(org.apache.poi.ss.usermodel.Sheet.BottomMargin, style.getMarginBottom());
		fileSheet.setMargin(org.apache.poi.ss.usermodel.Sheet.LeftMargin, style.getMarginLeft());
		fileSheet.setMargin(org.apache.poi.ss.usermodel.Sheet.RightMargin, style.getMarginRight());

		fileSheet.setAutobreaks(style.isAutoBreaks());
		fileSheet.setDisplayGuts(style.isDisplayGuts());
		fileSheet.setFitToPage(style.isFitToPage());
		fileSheet.setHorizontallyCenter(style.isHorizontallyCenter());
		fileSheet.setVerticallyCenter(style.isVerticallyCenter());
		fileSheet.setPrintGridlines(style.isPrintGridlines());

		if (style.isHeader() == true) {
			Header header = fileSheet.getHeader();

			header.setCenter(style.getHeaderCenter());
			header.setLeft(style.getHeaderLeft());
			header.setRight(style.getHeaderRight());
		}
		if (style.isFooter() == true) {
			Footer footer = fileSheet.getFooter();

			footer.setCenter(style.getFooterCenter());
			footer.setLeft(style.getFooterLeft());
			footer.setRight(style.getFooterRight());
		}

		if (style.isPrinterSetup() == true) {
			PrintSetup ps = fileSheet.getPrintSetup();

			ps.setHeaderMargin(style.getHeaderMargin());
			ps.setFooterMargin(style.getFooterMargin());
			ps.setFitHeight((short) style.getFitHeight());
			ps.setFitWidth((short) style.getFitWidth());
			ps.setDraft(style.isDraft());
			ps.setHeaderMargin(style.getHeaderMargin());
			ps.setFooterMargin(style.getFooterMargin());
			ps.setHResolution((short) style.getHorizontallyResolution());
			ps.setLandscape(style.isLandscape());
			ps.setLeftToRight(style.isLeftToRight());
			ps.setNoOrientation(style.isNoOrientation());
			ps.setNoColor(style.isNoColor());
			ps.setNotes(style.isNotes());
			ps.setPaperSize((short) style.getPaperSize());
			ps.setPageStart((short) style.getPageStart());
			ps.setScale((short) style.getScale());
			ps.setUsePage(style.isUsePage());
			ps.setValidSettings(style.isValidSettings());
			ps.setVResolution((short) style.getVerticallyResolution());
		}

		Iterator<Integer> columnIterator = style.getColumnBreakList().iterator();

		while (columnIterator.hasNext() == true) {
			fileSheet.setRowBreak(columnIterator.next());
		}

		Iterator<Integer> rowIterator = style.getRowBreakList().iterator();

		while (rowIterator.hasNext() == true) {
			fileSheet.setRowBreak(rowIterator.next());
		}
	}

	/**
	 * シートを生成します。
	 *
	 * @param sheet シート
	 */
	protected void createSheet(Sheet sheet) {

		fileSheet = fileBook.createSheet(sheet.getName());

		fileSheet.setDefaultColumnWidth(sheet.getDefaultColumnSize() / COLUMN_CORRECTING_VALUE);
		fileSheet.setDefaultRowHeight(sheet.getDefaultRowSize());

		fileSheet.getHeader().setLeft(sheet.getStyle().getHeaderLeft());
		fileSheet.getHeader().setCenter(sheet.getStyle().getHeaderCenter());
		fileSheet.getHeader().setRight(sheet.getStyle().getHeaderRight());

		fileSheet.getFooter().setLeft(sheet.getStyle().getFooterLeft());
		fileSheet.getFooter().setCenter(sheet.getStyle().getFooterCenter());
		fileSheet.getFooter().setRight(sheet.getStyle().getFooterRight());

		SheetStyle style = sheet.getStyle();

		createSheetStyle(style);

		int maxColumn = sheet.getMaxColumn();

		//		if (maxColumn > MAX_COLUMN_SIZE) {
		//			maxColumn = MAX_COLUMN_SIZE;
		//		}

		for (int i = 0; i < maxColumn; i++) {
			if (sheet.getColumnSizeMap().get(i) != null) {
				fileSheet.setColumnWidth(i, sheet.getColumnSizeMap().get(i));
			} else {
				fileSheet.setColumnWidth(i, sheet.getDefaultColumnSize());
			}
		}

		int maxRow = sheet.getMaxRow();

		for (int i = 0; i < maxRow; i++) {
			createRow(sheet, i);
		}
	}

	/**
	 * 行を生成します。
	 *
	 * @param sheet シート
	 * @param rowNumber 行番号
	 */
	protected void createRow(Sheet sheet, int rowNumber) {

		row = fileSheet.createRow(rowNumber);

		if (sheet.getRowSizeMap().get(rowNumber) != null) {
			row.setHeight(sheet.getRowSizeMap().get(rowNumber).shortValue());
		}

		int maxColumn = sheet.getMaxColumn();

		//		if (maxColumn > MAX_COLUMN_SIZE) {
		//			maxColumn = MAX_COLUMN_SIZE;
		//		}
		//
		for (int i = 0; i < maxColumn; i++) {
			createCell(sheet, i, rowNumber);
		}
	}

	/**
	 * セルを生成します。
	 *
	 * @param sheet シート
	 * @param columnNumber 列番号
	 * @param rowNumber 行番号
	 */
	protected void createCell(Sheet sheet, int columnNumber, int rowNumber) {

		fileCell = row.createCell(columnNumber);
		Cell cell = sheet.getCell(new Address(columnNumber, rowNumber));

		CellRange cellRange = cell.getMergedCell();

		if (cellRange != null) {
			CellRangeAddress range = new CellRangeAddress(cellRange.getBeginRow(), cellRange.getLastRow(), cellRange.getBeginColumn(), cellRange.getLastColumn());

			fileSheet.addMergedRegion(range);
		}

		setCellValue(cell);
		setCellStyle(cell);
	}

	/**
	 * セルに値を設定します。
	 *
	 * @param cell セル
	 */
	protected void setCellValue(Cell cell) {

		if (cell.getParentCell() != null || cell.getValueType() == ValueType.NONE) {
			return;
		} else if (cell.getValueType() == ValueType.BOOLEAN) {
			fileCell.setCellValue(cell.getValueDate());
		} else if (cell.getValueType() == ValueType.NUMBER) {
			fileCell.setCellValue(cell.getValueNumber().doubleValue());
		} else if (cell.getValueType() == ValueType.DATE) {
			fileCell.setCellValue(cell.getValueDate());
		} else {
			String text = cell.getValueText();

			if (Poi2ccUtil.isEmptyString(text) == false) {
				fileCell.setCellValue(text);
			}
		}
		if (cell.isFormula() == true) {
			try {
				fileCell.setCellFormula(cell.getFormulaValue());
			} catch (FormulaParseException e) {
			}
		}
	}

	/**
	 * セルにスタイルを設定します。
	 *
	 * @param cell セル
	 */
	protected void setCellStyle(Cell cell) {

		CellFont cellFont = cell.getFont();
		XSSFFont font = fontCache.get(cellFont);

		if (font == null && Poi2ccUtil.isEmptyString(cellFont.getFontName()) == false) {
			font = fileBook.createFont();

			font.setBold(cellFont.getFontBold());
			font.setFontName(cellFont.getFontName());
			font.setColor(cellFont.getFontColor());
			font.setFontHeight(cellFont.getFontSize());
			font.setItalic(cellFont.isFontItalic());
			font.setStrikeout(cellFont.isFontStrikeout());
			font.setUnderline(cellFont.getFontUnderline());

			fontCache.put(cellFont, font);
		}

		com.kiruah.poi2cc.storage.sub.CellStyle cellStyle = cell.getStyle();
		XSSFCellStyle style = styleCache.get(cellStyle);

		if (style == null || font == null) {
			if (style == null) {
				style = fileBook.createCellStyle();

				style.setLocked(cell.isLocked());

				style.setAlignment(cellStyle.getAlignment());
				style.setVerticalAlignment(cellStyle.getVerticalAlignment());
				style.setWrapText(cellStyle.isWrapText());
				style.setRotation(cellStyle.getRotation());
				style.setIndention(cellStyle.getIndent());

				IndexedColorMap colorMap = fileBook.getStylesSource().getIndexedColors();
				byte[] rgbBackgroundArray = new byte[] {(byte) cellStyle.getBackgroundColor().getRed(), (byte) cellStyle.getBackgroundColor().getGreen(), (byte) cellStyle.getBackgroundColor().getBlue()};
				byte[] rgbForegroundArray = new byte[] {(byte) cellStyle.getForegroundColor().getRed(), (byte) cellStyle.getForegroundColor().getGreen(), (byte) cellStyle.getForegroundColor().getBlue()};

				style.setFillBackgroundColor(new XSSFColor(rgbBackgroundArray, colorMap));
				style.setFillForegroundColor(new XSSFColor(rgbForegroundArray, colorMap));
				style.setFillPattern(cellStyle.getFillPattern());

				style.setBorderBottom(cellStyle.getBorderBottom());
				style.setBorderLeft(cellStyle.getBorderLeft());
				style.setBorderRight(cellStyle.getBorderRight());
				style.setBorderTop(cellStyle.getBorderTop());

				if (Poi2ccUtil.isEmptyString(cellStyle.getFormat()) == false) {
					short dataFormat = fileDataFormat.getFormat(cellStyle.getFormat());

					style.setDataFormat(dataFormat);
				}

				styleCache.put(cellStyle, style);
			}

			if (font != null) {
				style.setFont(font);
			}
		}

		fileCell.setCellStyle(style);
	}

	/**
	 * ExcelWriter コンストラクタ
	 */
	protected ExcelWriter() {

	}
}
