package demo;



import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;

import javax.swing.filechooser.FileSystemView;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public class Read2 {
	public static void main(String[] args) throws IOException {
		FileSystemView fsv = FileSystemView.getFileSystemView();
		String desktop = fsv.getHomeDirectory().getPath();
		String filePath = "E:/bb.xls";
		FileInputStream fileInputStream = new FileInputStream(filePath);
		BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
		POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
		// 创建 Excel文件对象
		HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
		// 返回Sheet对象
		HSSFSheet sheet = workbook.getSheet("Sheet1");

		int lastRowIndex = sheet.getLastRowNum(); // 索引�?0�?�?

		JSONArray jsonArray = new JSONArray();

		JSONObject rowobj1 = new JSONObject();
		JSONArray jsonArray1 = new JSONArray();

		JSONObject rowobj2 = new JSONObject();
		JSONArray jsonArray2 = new JSONArray();

		JSONObject rowobj3 = new JSONObject();
		JSONArray jsonArray3 = new JSONArray();

		JSONObject rowobj4 = new JSONObject();
		JSONArray jsonArray4 = new JSONArray();

		JSONObject rowobj5 = new JSONObject();
		JSONArray jsonArray5 = new JSONArray();

		for (int i = 0; i <= lastRowIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 得到�?
			if (row == null) {
				break;
			}

			// �?5�? -- 5�?
			HSSFCell cell5 = row.getCell(4);
			if (cell5 != null && !"".equals(cell5)) {
				String cellValue5 = row.getCell(4).getStringCellValue();
				Result mergedRegion5 = isMergedRegion(sheet, i, 4);
				if (mergedRegion5.merged == false) {
					// 不是合并�?
					if (!"".equals(cellValue5)) {
						rowobj5.put("text", cellValue5);
						jsonArray5.add(rowobj5);
					}
				}
			}

			// �?4�? -- 4�?
			HSSFCell cell4 = row.getCell(3);
			if (cell4 != null && !"".equals(cell4)) {
				String cellValue4 = row.getCell(3).getStringCellValue();
				Result mergedRegion4 = isMergedRegion(sheet, i, 3);
				if (mergedRegion4.merged == false) {// 不是合并�?
					if (!"".equals(cellValue4)) {
						rowobj4.put("text", cellValue4);
						if (rowobj5 != null && jsonArray5.size() > 0) {
							rowobj4.put("children", jsonArray5);
						}
						jsonArray4.add(rowobj4);
						rowobj5.clear();
						jsonArray5.clear();
					}
				} else { // 是合并的
					if (i == mergedRegion4.startRow - 1) {
						rowobj4.put("text", cellValue4);
					}
					if (i == mergedRegion4.endRow - 1) {
						rowobj4.put("children", jsonArray5);
						jsonArray4.add(rowobj4);
						rowobj5.clear();
						jsonArray5.clear();
					}
				}
			}

			// �?3�? -- 三级
			HSSFCell cell3 = row.getCell(2);
			if (cell3 != null && !"".equals(cell3)) {
				String cellValue3 = row.getCell(2).getStringCellValue();
				Result mergedRegion3 = isMergedRegion(sheet, i, 2);
				if (mergedRegion3.merged == false) {
					// 不是合并�?
					if (!"".equals(cellValue3)) {

						rowobj3.put("text", cellValue3);
						if (jsonArray4.size() > 0) {
							rowobj3.put("children", jsonArray4);
						}

						jsonArray3.add(rowobj3);
						rowobj4.clear();
						jsonArray4.clear();
					}
				} else { // 是合并的
					if (i == mergedRegion3.startRow - 1) {
						rowobj3.put("text", cellValue3);
					}
					if (i == mergedRegion3.endRow - 1) {
						rowobj3.put("children", jsonArray4);
						jsonArray3.add(rowobj3);
						rowobj4.clear();
						jsonArray4.clear();
					}
				}

			}

			// �?2�? - 二级目录
			HSSFCell cell2 = row.getCell(1);
			if (cell2 != null && !"".equals(cell2)) {
				String cellValue2 = row.getCell(1).getStringCellValue();
				Result mergedRegion2 = isMergedRegion(sheet, i, 1);
				if (mergedRegion2.merged == false) {
					// 不是合并�?
					if (!"".equals(cellValue2)) {
						rowobj2.put("text", cellValue2);
						if (jsonArray3.size() > 0) {
							rowobj2.put("children", jsonArray3);
						}
						jsonArray2.add(rowobj2);
						rowobj3.clear();
						jsonArray3.clear();
					}

				} else {
					if (i == mergedRegion2.startRow - 1) {
						rowobj2.put("text", cellValue2);
					}
					if (i == mergedRegion2.endRow - 1) {
						rowobj2.put("children", jsonArray3);
						jsonArray2.add(rowobj2);
						rowobj3.clear();
						jsonArray3.clear();
					}

				}
			}

			// �?1�? -- �?级目�?
			HSSFCell cell1 = row.getCell(0);
			if (cell1 != null && !"".equals(cell1)) {
				String cellValue1 = row.getCell(0).getStringCellValue();
				Result mergedRegion1 = isMergedRegion(sheet, i, 0);
				if (mergedRegion1.merged == false) {
					// 不是合并�?
					if (!"".equals(cellValue1)) {
						rowobj1.put("text", cellValue1);
						if (jsonArray2.size() > 0) {
							rowobj1.put("children", jsonArray2);
						}
						jsonArray.add(rowobj1);
						rowobj2.clear();
						jsonArray2.clear();
						rowobj1.clear();
					}

				} else {
					if (i == mergedRegion1.startRow - 1) {
						rowobj1.put("text", cellValue1);
					}
					if (i == mergedRegion1.endRow - 1) {
						rowobj1.put("children", jsonArray2);
						jsonArray.add(rowobj1);
						rowobj2.clear();
						jsonArray2.clear();
						rowobj1.clear();
					}
				}
			}

		}
		System.out.println(jsonArray);
		bufferedInputStream.close();
	}
	// 判断是否是合并单元格
	private static Result isMergedRegion(HSSFSheet sheet, int row, int column) {
		int sheetMergeCount = ((org.apache.poi.ss.usermodel.Sheet) sheet).getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = ((org.apache.poi.ss.usermodel.Sheet) sheet).getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return new Result(true, firstRow + 1, lastRow + 1, firstColumn + 1, lastColumn + 1);
				}
			}
		}
		return new Result(false, 0, 0, 0, 0);
	}
}
