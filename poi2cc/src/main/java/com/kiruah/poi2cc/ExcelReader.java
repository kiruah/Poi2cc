package com.kiruah.poi2cc;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import com.kiruah.poi2cc.storage.Address;
import com.kiruah.poi2cc.storage.Book;
import com.kiruah.poi2cc.storage.Cell;
import com.kiruah.poi2cc.storage.Sheet;
import com.kiruah.poi2cc.storage.sub.CellFont;
import com.kiruah.poi2cc.storage.sub.CellRange;
import com.kiruah.poi2cc.storage.sub.SheetStyle;

/**
 * Excelファイルをロードします。
 *
 * @author Kiruah
 */
public class ExcelReader {

	/** カラムサイズデフォルト値補正数 */
	protected static final int COLUMN_CORRECTING_VALUE = 256;

	/** 現在処理中のブック */
	protected org.apache.poi.ss.usermodel.Workbook fileBook = null;

	/** 現在処理中のシート */
	protected org.apache.poi.ss.usermodel.Sheet fileSheet = null;

	/** 現在処理中の行 */
	protected org.apache.poi.ss.usermodel.Row fileRow = null;

	/** 現在処理中のセル */
	protected org.apache.poi.ss.usermodel.Cell fileCell = null;

	/**
	 * Excelファイルロードします。
	 *
	 * @param fileName ファイル名
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws FileNotFoundException ファイルがない場合の例外
	 * @throws IOException IO例外
	 */
	public static Book loadBook(String fileName) throws InvalidFormatException, FileNotFoundException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBook(fileName, new FileInputStream(fileName));
	}

	/**
	 * システムリソースからファイルをロードします。
	 *
	 * @param fileName ファイル名
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookFromResource(String fileName) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBook(fileName, Thread.currentThread().getContextClassLoader().getSystemResourceAsStream(fileName));
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param file ファイルオブジェクト
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws FileNotFoundException ファイルがない場合の例外
	 * @throws IOException IO例外
	 */
	public static Book loadBook(File file) throws InvalidFormatException, FileNotFoundException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBook(file.getAbsolutePath(), new FileInputStream(file));
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param uri URI
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBook(URI uri) throws InvalidFormatException, IOException {

		return loadBook(uri.toURL());
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param url URL
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBook(URL url) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBook(url.toExternalForm(), url.openStream());
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param uri URI(ファイル名の元情報として利用します。)
	 * @param in Excelファイルを読み取る入力ストリーム
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBook(URI uri, InputStream in) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();
		String name = null;

		if (uri != null) {
			name = uri.toString();
		}

		return reader.loadBook(name, in);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param fileName ファイル名(ファイル名の元情報として利用します。)
	 * @param in Excelファイルを読み取る入力ストリーム
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public Book loadBook(String fileName, InputStream in) throws InvalidFormatException, IOException {

		Book book = new Book(fileName);

		try {
			fileBook = WorkbookFactory.create(in);

			int sheetNumber = fileBook.getNumberOfSheets();

			for (int i = 0; i < sheetNumber; i++) {
				String sheetName = fileBook.getSheetName(i);
				fileSheet = fileBook.getSheetAt(i);

				Sheet sheet = loadSheet(sheetName);

				book.addSheet(sheet);
			}
		} finally {
			Poi2ccUtil.close(in);
		}

		return book;
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param fileName ファイル名
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws FileNotFoundException ファイルがない場合の例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedIncludeSheet(String fileName, String... includeSheetNameArray) throws InvalidFormatException, FileNotFoundException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedIncludeSheet(fileName, new FileInputStream(fileName), includeSheetNameArray);
	}

	/**
	 * システムリソースからファイルをロードします。
	 *
	 * @param fileName ファイル名
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookFromResourceSpecifiedIncludeSheet(String fileName, String... includeSheetNameArray) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedIncludeSheet(fileName, Thread.currentThread().getContextClassLoader().getSystemResourceAsStream(fileName), includeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param file ファイルオブジェクト
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws FileNotFoundException ファイルがない場合の例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedIncludeSheet(File file, String... includeSheetNameArray) throws InvalidFormatException, FileNotFoundException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedIncludeSheet(file.getAbsolutePath(), new FileInputStream(file), includeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param uri URI
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedIncludeSheet(URI uri, String... includeSheetNameArray) throws InvalidFormatException, IOException {

		return loadBookSpecifiedIncludeSheet(uri.toURL(), includeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param url URL
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedIncludeSheet(URL url, String... includeSheetNameArray) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedIncludeSheet(url.toExternalForm(), url.openStream(), includeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param uri URI(ファイル名の元情報として利用します。)
	 * @param in Excelファイルを読み取る入力ストリーム
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedIncludeSheet(URI uri, InputStream in, String... includeSheetNameArray) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedIncludeSheet(uri.toString(), in, includeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param fileName ファイル名(ファイル名の元情報として利用します。)
	 * @param in Excelファイルを読み取る入力ストリーム
	 * @param includeSheetNameArray ロードするシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public Book loadBookSpecifiedIncludeSheet(String fileName, InputStream in, String... includeSheetNameArray) throws InvalidFormatException, IOException {

		Book book = new Book(fileName);

		Map<String, String> map = new HashMap<String, String>();

		for (String sheetName : includeSheetNameArray) {
			map.put(sheetName, sheetName);
		}

		try {
			fileBook = WorkbookFactory.create(in);

			int sheetNumber = fileBook.getNumberOfSheets();

			for (int i = 0; i < sheetNumber; i++) {
				String sheetName = fileBook.getSheetName(i);

				if (map.containsKey(sheetName) == true) {
					fileSheet = fileBook.getSheetAt(i);

					Sheet sheet = loadSheet(sheetName);

					book.addSheet(sheet);
				}
			}
		} finally {
			Poi2ccUtil.close(in);
		}

		return book;
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param fileName ファイル名
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws FileNotFoundException ファイルがない場合の例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedExcludeSheet(String fileName, String... excludeSheetNameArray) throws InvalidFormatException, FileNotFoundException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedExcludeSheet(fileName, new FileInputStream(fileName), excludeSheetNameArray);
	}

	/**
	 * システムリソースからファイルをロードします。
	 *
	 * @param fileName ファイル名
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookFromResourceSpecifiedExcludeSheet(String fileName, String... excludeSheetNameArray) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedExcludeSheet(fileName, Thread.currentThread().getContextClassLoader().getSystemResourceAsStream(fileName), excludeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param file ファイルオブジェクト
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws FileNotFoundException ファイルがない場合の例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedExcludeSheet(File file, String... excludeSheetNameArray) throws InvalidFormatException, FileNotFoundException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedExcludeSheet(file.getAbsolutePath(), new FileInputStream(file), excludeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param uri URI
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedExcludeSheet(URI uri, String... excludeSheetNameArray) throws InvalidFormatException, IOException {

		return loadBookSpecifiedExcludeSheet(uri.toURL(), excludeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param url URL
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedExcludeSheet(URL url, String... excludeSheetNameArray) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedExcludeSheet(url.toExternalForm(), url.openStream(), excludeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param uri URI(ファイル名の元情報として利用します。)
	 * @param in Excelファイルを読み取る入力ストリーム
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public static Book loadBookSpecifiedExcludeSheet(URI uri, InputStream in, String... excludeSheetNameArray) throws InvalidFormatException, IOException {

		ExcelReader reader = new ExcelReader();

		return reader.loadBookSpecifiedExcludeSheet(uri.toString(), in, excludeSheetNameArray);
	}

	/**
	 * Excelファイルロードします。
	 *
	 * @param fileName ファイル名(ファイル名の元情報として利用します。)
	 * @param in Excelファイルを読み取る入力ストリーム
	 * @param excludeSheetNameArray ロード除外対象のシート名の配列
	 * @return Excelファイル内容を保持したブック情報
	 * @throws InvalidFormatException 不正なフォーマット例外
	 * @throws IOException IO例外
	 */
	public Book loadBookSpecifiedExcludeSheet(String fileName, InputStream in, String... excludeSheetNameArray) throws InvalidFormatException, IOException {

		Book book = new Book(fileName);

		Map<String, String> map = new HashMap<String, String>();

		for (String sheetName : excludeSheetNameArray) {
			map.put(sheetName, sheetName);
		}

		try {
			fileBook = WorkbookFactory.create(in);

			int sheetNumber = fileBook.getNumberOfSheets();

			for (int i = 0; i < sheetNumber; i++) {
				String sheetName = fileBook.getSheetName(i);

				if (map.containsKey(sheetName) == false) {
					fileSheet = fileBook.getSheetAt(i);

					Sheet sheet = loadSheet(sheetName);

					book.addSheet(sheet);
				}
			}
		} finally {
			Poi2ccUtil.close(in);
		}

		return book;
	}

	protected void loadSheetStyle(SheetStyle style) {

		style.setDisplayGridLine(fileSheet.isDisplayGridlines());
		style.setDisplayZeros(fileSheet.isDisplayZeros());

		style.setMarginTop(fileSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.TopMargin));
		style.setMarginBottom(fileSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.BottomMargin));
		style.setMarginLeft(fileSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.LeftMargin));
		style.setMarginRight(fileSheet.getMargin(org.apache.poi.ss.usermodel.Sheet.RightMargin));

		style.setAutoBreaks(fileSheet.getAutobreaks());
		style.setDisplayGuts(fileSheet.getDisplayGuts());
		style.setFitToPage(fileSheet.getFitToPage());
		style.setHorizontallyCenter(fileSheet.getHorizontallyCenter());
		style.setVerticallyCenter(fileSheet.getVerticallyCenter());
		style.setPrintGridlines(fileSheet.isPrintGridlines());

		Header header = fileSheet.getHeader();

		if (header != null) {
			style.setHeaderCenter(header.getCenter());
			style.setHeaderLeft(header.getLeft());
			style.setHeaderRight(header.getRight());
		}

		Footer footer = fileSheet.getFooter();

		if (footer != null) {
			style.setFooterCenter(footer.getCenter());
			style.setFooterLeft(footer.getLeft());
			style.setFooterRight(footer.getRight());
		}

		PrintSetup printSetup = fileSheet.getPrintSetup();

		if (printSetup != null) {
			style.setHeaderMargin(printSetup.getHeaderMargin());
			style.setFooterMargin(printSetup.getFooterMargin());
			style.setFitHeight(printSetup.getFitHeight());
			style.setFitWidth(printSetup.getFitWidth());
			style.setDraft(printSetup.getDraft());
			style.setHorizontallyResolution(printSetup.getHResolution());
			style.setLandscape(printSetup.getLandscape());
			style.setLeftToRight(printSetup.getLeftToRight());
			style.setNoOrientation(printSetup.getNoOrientation());
			style.setNoColor(printSetup.getNoColor());
			style.setNotes(printSetup.getNotes());
			style.setPaperSize(printSetup.getPaperSize());
			style.setPageStart(printSetup.getPageStart());
			style.setScale(printSetup.getScale());
			style.setUsePage(printSetup.getUsePage());
			style.setValidSettings(printSetup.getValidSettings());
			style.setVerticallyResolution(printSetup.getVResolution());
		}

		int[] columnBreakArray = fileSheet.getColumnBreaks();

		if (columnBreakArray != null) {
			for (int value : columnBreakArray) {
				style.getColumnBreakList().add(value);
			}
		}

		int[] rowBreakArray = fileSheet.getRowBreaks();

		if (rowBreakArray != null) {
			for (int value : rowBreakArray) {
				style.getRowBreakList().add(value);
			}
		}
	}

	/**
	 * シート情報をロードします。
	 *
	 * @param sheetName シート名
	 * @return ロードしたシート情報
	 */
	protected Sheet loadSheet(String sheetName) {

		Sheet sheet = new Sheet(sheetName);

		sheet.setDefaultColumnSize(fileSheet.getDefaultColumnWidth() * COLUMN_CORRECTING_VALUE);
		sheet.setDefaultRowSize(fileSheet.getDefaultRowHeight());

		SheetStyle style = sheet.getStyle();

		loadSheetStyle(style);

		int beginRow = fileSheet.getFirstRowNum();
		int lastRow = fileSheet.getLastRowNum();

		for (int i = beginRow; i <= lastRow; i++) {
			fileRow = fileSheet.getRow(i);

			if (fileRow != null) {
				sheet.getRowSizeMap().put(i, (int) fileRow.getHeight());

				loadRow(i, sheet);
			}
		}

		for (int i = 0; i < fileSheet.getNumMergedRegions(); i++) {
			CellRangeAddress range = fileSheet.getMergedRegion(i);

			loadMergedCell(sheet, range);
		}

		int maxRow = sheet.getMaxRow();
		int maxColumn = sheet.getMaxColumn();
		Address address = new Address();

		for (int column = 0; column < maxColumn; column++) {
			for (int row = 0; row < maxRow; row++) {
				address.setAddress(column, row);

				Cell cell = sheet.getCell(address);

				if (cell == null) {
					cell = new Cell();
					sheet.setCell(address, cell);
				}

				cell.setEdited(false);
			}
		}

		return sheet;
	}

	/**
	 * 結合セル情報をロードします。
	 *
	 * @param sheet シート
	 * @param range セルの結合情報
	 */
	protected void loadMergedCell(Sheet sheet, CellRangeAddress range) {

		CellRange cellRange = new CellRange(range.getFirstColumn(), range.getFirstRow(), range.getLastColumn(), range.getLastRow());
		Address address = new Address(range.getFirstColumn(), range.getFirstRow());
		Cell baseCell = sheet.getCell(address);

		for (int i = range.getFirstRow(); i <= range.getLastRow(); i++) {
			for (int j = range.getFirstColumn(); j <= range.getLastColumn(); j++) {
				address.setAddress(j, i);

				sheet.getCell(address).setParentCell(baseCell);
			}
		}

		baseCell.setMergedCell(cellRange);
		baseCell.setParentCell(null);
	}

	/**
	 * 特定の行をロードします。
	 *
	 * @param rowNumber 行番号
	 * @param sheet シート
	 */
	protected void loadRow(int rowNumber, Sheet sheet) {

		int endColumn = fileRow.getLastCellNum();

		for (int columnNumber = 0; columnNumber <= endColumn; columnNumber++) {
			sheet.getColumnSizeMap().put(columnNumber, fileSheet.getColumnWidth(columnNumber));

			Address address = new Address(columnNumber, rowNumber);

			try {
				fileCell = fileRow.getCell(columnNumber);

				Cell cell = getCell();
				sheet.setCell(address, cell);
			} catch (Exception e) {
				Cell cell = new Cell();

				sheet.setCell(address, cell);
			}
		}
	}

	/**
	 * 特定のセルをロードします。
	 *
	 * @return セル情報
	 */
	protected Cell getCell() {

		if (fileCell == null) {
			return new Cell();
		}

		Cell cell = new Cell(null, null, null);

		setCellValue(cell);
		setCellComment(cell);

		CellStyle style = fileCell.getCellStyle();

		cell.setLocked(style.getLocked());
		setCellFont(style, cell);
		setCellStyle(style, cell);

		return cell;
	}

	protected void setCellValue(Cell cell, CellType cellType) {

		if (cellType == CellType.BLANK) {
			cell.setValue(Poi2ccConstants.EMPTY_STRING);
			cell.setError(false);
		} else if (cellType == CellType.ERROR) {
			cell.setValue(fileCell.getErrorCellValue());
			cell.setError(true);
		} else if (cellType == CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(fileCell) == true) {
				cell.setValue(fileCell.getLocalDateTimeCellValue());
			} else {
				cell.setValue(fileCell.getNumericCellValue());
			}
			cell.setError(false);
		} else if (cellType == CellType.STRING) {
			cell.setValue(fileCell.getStringCellValue());
			cell.setError(false);
		} else if (cellType == CellType.BOOLEAN) {
			cell.setValue(fileCell.getBooleanCellValue());
			cell.setError(false);
		}
	}

	/**
	 * セル値を設定します。
	 *
	 * @param cell セル情報
	 */
	protected void setCellValue(Cell cell) {

		CellType cellType = fileCell.getCellType();

		if (cellType == CellType.FORMULA) {
			CellType innerCellType = fileCell.getCachedFormulaResultType();
			setCellValue(cell, innerCellType);
			cell.setFormulaValue(fileCell.getCellFormula());
			cell.setFormula(true);
		} else {
			setCellValue(cell, cellType);
		}
	}

	/**
	 * セルのフォント情報をロードし設定します。
	 *
	 * @param style スタイル
	 * @param cell セル情報
	 */
	protected void setCellFont(CellStyle style, Cell cell) {

		int fontIndex = style.getFontIndexAsInt();
		CellFont cellFont = null;

		Font font = fileBook.getFontAt(fontIndex);
		cellFont = new CellFont();

		cellFont.setFontBold(font.getBold());
		cellFont.setFontColor(font.getColor());
		cellFont.setFontSize(font.getFontHeight());
		cellFont.setFontName(font.getFontName());
		cellFont.setFontItalic(font.getItalic());
		cellFont.setFontStrikeout(font.getStrikeout());
		cellFont.setFontUnderline(font.getUnderline());

		cell.setFont(cellFont);
	}

	protected Color getColor(CellStyle style, org.apache.poi.ss.usermodel.Color color, Color defaultColor) {

		try {
			if (color instanceof ExtendedColor == true) {
				ExtendedColor c = (ExtendedColor) color;
				byte[] rgbArray = c.getRGBWithTint();

				if (rgbArray != null) {
					int red = rgbArray[0] & 0xFF;
					int green = rgbArray[1] & 0xFF;
					int blue = rgbArray[2] & 0xFF;
					Color awtColor = new Color(red, green, blue);

					return awtColor;
				}
			} else if (color instanceof HSSFColor == true) {
				HSSFColor c = (HSSFColor) color;

				short[] rgbArray = c.getTriplet();
				if (rgbArray != null) {
					Color awtColor = new Color(rgbArray[0], rgbArray[1], rgbArray[2]);

					return awtColor;
				}
			}
		} catch (Exception e) {
		}

		return defaultColor;
	}

	/**
	 * セルのスタイル情報をロードし設定します。
	 *
	 * @param style スタイル
	 * @param cell セル情報
	 */
	protected void setCellStyle(CellStyle style, Cell cell) {

		com.kiruah.poi2cc.storage.sub.CellStyle cellStyle = new com.kiruah.poi2cc.storage.sub.CellStyle();

		// format
		cellStyle.setFormat(style.getDataFormatString());
		cellStyle.setAlignment(style.getAlignment());
		cellStyle.setVerticalAlignment(style.getVerticalAlignment());
		cellStyle.setIndent(style.getIndention());
		cellStyle.setWrapText(style.getWrapText());
		cellStyle.setRotation(style.getRotation());

		// color
		cellStyle.setForegroundColor(getColor(style, style.getFillForegroundColorColor(), Color.BLACK));
		cellStyle.setBackgroundColor(getColor(style, style.getFillBackgroundColorColor(), Color.WHITE));
		cellStyle.setFillPattern(style.getFillPattern());

		// line
		cellStyle.setBorderBottom(style.getBorderBottom());
		cellStyle.setBorderLeft(style.getBorderLeft());
		cellStyle.setBorderRight(style.getBorderRight());
		cellStyle.setBorderTop(style.getBorderTop());

		cell.setStyle(cellStyle);
	}

	/**
	 * セルのコメントをロードしセルに設定します。
	 *
	 * @param cell セル情報
	 */
	protected void setCellComment(Cell cell) {

		Comment comment = fileCell.getCellComment();

		if (fileCell.getCellComment() != null) {
			RichTextString text = comment.getString();

			if (text != null) {
				cell.setComment(text.getString());
			}
		}
	}

	/**
	 * ExcelReader コンストラクタ
	 */
	protected ExcelReader() {

	}
}
