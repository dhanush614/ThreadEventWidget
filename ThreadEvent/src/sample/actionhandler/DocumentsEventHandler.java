package sample.actionhandler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.collection.FolderSet;
import com.filenet.api.constants.AutoClassify;
import com.filenet.api.constants.AutoUniqueName;
import com.filenet.api.constants.CheckinType;
import com.filenet.api.constants.DefineSecurityParentage;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.constants.RefreshMode;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.core.ReferentialContainmentRelationship;
import com.filenet.api.engine.EventActionHandler;
import com.filenet.api.events.ObjectChangeEvent;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.Properties;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.util.Id;

import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.P8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;

public class DocumentsEventHandler implements EventActionHandler {
	public void onEvent(ObjectChangeEvent event, Id subId) {
		System.out.println("Inside onEvent method");
		CaseMgmtContext origCmctx = null;
		try {
			P8ConnectionCache connCache = new SimpleP8ConnectionCache();
			origCmctx = CaseMgmtContext.set(new CaseMgmtContext(new SimpleVWSessionCache(), connCache));
			ObjectStore os = event.getObjectStore();
			System.out.println("OS" + os);
			ObjectStoreReference targetOsRef = new ObjectStoreReference(os);
			System.out.println("TOS" + targetOsRef);
			Id id = event.get_SourceObjectId();
			FilterElement fe = new FilterElement(null, null, null, "Owner Name", null);
			PropertyFilter pf = new PropertyFilter();
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_SIZE, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_ELEMENTS, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.FOLDERS_FILED_IN, null));
			pf.addIncludeProperty(fe);
			Document doc = Factory.Document.fetchInstance(os, id, pf);
			System.out.println("Document Name" + doc.get_Name());
			ContentElementList docContentList = doc.get_ContentElements();
			Iterator iter = docContentList.iterator();
			while (iter.hasNext()) {
				ContentTransfer ct = (ContentTransfer) iter.next();
				InputStream stream = ct.accessContentStream();
				String docTitle = doc.get_Name();
				HashMap<Integer, HashMap<String, Object>> excelRows = new HashMap<Integer, HashMap<String, Object>>();
				HashMap<Integer, HashMap<String, Object>> responseData = new HashMap<Integer, HashMap<String, Object>>();
				XSSFWorkbook workbook = new XSSFWorkbook(stream);
				excelRows = readExcelRows(os, workbook);
				responseData = threadExecMethod(excelRows, docTitle, os);
				updateDocument(responseData, os, doc, workbook, stream);
			}
		} catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
			throw new RuntimeException(e);
		} finally {
			CaseMgmtContext.set(origCmctx);
		}
	}

	public static HashMap<Integer, HashMap<String, Object>> threadExecMethod(
			HashMap<Integer, HashMap<String, Object>> excelRows, String docTitle, ObjectStore targetOS) {
		HashMap<Integer, HashMap<String, Object>> responseMap = new HashMap<Integer, HashMap<String, Object>>();
		ExecutorService threadExecutor = Executors.newFixedThreadPool(10);
		List<Future<HashMap<Integer, HashMap<String, Object>>>> responseList = new ArrayList<Future<HashMap<Integer, HashMap<String, Object>>>>();
		Iterator<Entry<Integer, HashMap<String, Object>>> excelRow = excelRows.entrySet().iterator();
		while (excelRow.hasNext()) {
			try {
				Entry<Integer, HashMap<String, Object>> propertyPair = excelRow.next();
				Future<HashMap<Integer, HashMap<String, Object>>> threadList = threadExecutor
						.submit(new ThreadClass(propertyPair.getKey(), propertyPair.getValue(), docTitle, targetOS));
				responseList.add(threadList);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		for (Future<HashMap<Integer, HashMap<String, Object>>> object : responseList) {
			try {
				HashMap<Integer, HashMap<String, Object>> map = object.get();
				Iterator<Entry<Integer, HashMap<String, Object>>> caseProperty = map.entrySet().iterator();
				while (caseProperty.hasNext()) {
					Entry<Integer, HashMap<String, Object>> propertyPair = caseProperty.next();
					responseMap.put(propertyPair.getKey(), propertyPair.getValue());
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		threadExecutor.shutdown();
		return responseMap;
	}

	public static HashMap<Integer, HashMap<String, Object>> readExcelRows(ObjectStore targetOS, XSSFWorkbook workbook)
			throws IOException {
		int rowLastCell = 0;
		HashMap<Integer, String> headers = new HashMap<Integer, String>();
		HashMap<String, String> propDescMap = new HashMap<String, String>();
		HashMap<Integer, HashMap<String, Object>> caseProperties = new HashMap<Integer, HashMap<String, Object>>();
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFSheet sheet1 = workbook.getSheetAt(1);
		Iterator<Row> rowIterator = sheet.iterator();
		Iterator<Row> rowIterator1 = sheet1.iterator();
		while (rowIterator1.hasNext()) {
			Row row = rowIterator1.next();
			if (row.getRowNum() > 0) {
				String key = null, value = null;
				key = row.getCell(0).getStringCellValue();
				value = row.getCell(1).getStringCellValue();
				if (key != null && value != null) {
					propDescMap.put(key, value);
				}
			}
		}
		String headerValue;
		int rowNum = 0;
		if (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			int colNum = 0;
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				headerValue = cell.getStringCellValue();
				if (headerValue.contains("*")) {
					if (headerValue.contains("datetime")) {
						headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
						headerValue += "dateField";
					} else {
						headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
					}
				}
				if (headerValue.contains("datetime")) {
					headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
					headerValue += "dateField";
				} else {
					headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
				}
				headers.put(colNum++, headerValue);
			}
			rowLastCell = row.getLastCellNum();
			Cell cell1 = row.createCell(rowLastCell, Cell.CELL_TYPE_STRING);
			if (row.getRowNum() == 0) {
				cell1.setCellValue("Status");
			}
		}
		int rowStart = sheet.getFirstRowNum() + 1;
		int rowEnd = sheet.getLastRowNum();
		for (int rowNumber = rowStart; rowNumber <= rowEnd; rowNumber++) {
			Row row = sheet.getRow(rowNumber);
			if (row == null) {
				break;
			} else {
				HashMap<String, Object> rowValue = new HashMap<String, Object>();
				// Row row = rowIterator.next();
				int colNum = 0;
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
					try {
						if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							colNum++;
						} else {
							if (headers.get(colNum).contains("dateField")) {
								String symName = headers.get(colNum).replace("dateField", "");
								if (HSSFDateUtil.isCellDateFormatted(cell)) {
									Date date = cell.getDateCellValue();
									rowValue.put(propDescMap.get(symName), date);
									colNum++;
								}
							} else {
								rowValue.put(propDescMap.get(headers.get(colNum)), getCharValue(cell));
								colNum++;
							}
						}
					} catch (Exception e) {
						System.out.println(e);
						e.printStackTrace();
					}

				}
				caseProperties.put(++rowNum, rowValue);
			}
		}
		return caseProperties;
	}

	public static void updateDocument(HashMap<Integer, HashMap<String, Object>> responseMap, ObjectStore os,
			Document doc, XSSFWorkbook workbook, InputStream stream) throws IOException {
		// TODO Auto-generated method stub
		XSSFSheet sheet = workbook.getSheetAt(0);
		int lastCellNum = sheet.getRow(0).getLastCellNum();
		int rowNum;
		Iterator<Entry<Integer, HashMap<String, Object>>> caseProperty = responseMap.entrySet().iterator();
		while (caseProperty.hasNext()) {
			try {
				Entry<Integer, HashMap<String, Object>> propertyPair = caseProperty.next();
				rowNum = propertyPair.getKey();
				Row row = sheet.getRow(rowNum);
				Cell cell = row.createCell(lastCellNum - 1);
				HashMap<String, Object> propertyValues = (propertyPair.getValue());
				String status = propertyValues.get("Status").toString();
				if (status == "Success") {
					cell.setCellValue("Success");
				} else {
					cell.setCellValue("Failure");
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		InputStream is = null;
		ByteArrayOutputStream bos = null;
		try {
			bos = new ByteArrayOutputStream();
			workbook.write(bos);
			byte[] barray = bos.toByteArray();
			is = new ByteArrayInputStream(barray);
			String docTitle = doc.get_Name();
			FolderSet folderSet = doc.get_FoldersFiledIn();
			Folder folder = null;
			Iterator<Folder> folderSetIterator = folderSet.iterator();
			if (folderSetIterator.hasNext()) {
				folder = folderSetIterator.next();
			}
			String folderPath = folder.get_PathName();
			folderPath += " Response";
			Folder responseFolder = Factory.Folder.fetchInstance(os, folderPath, null);
			String docClassName = doc.getClassName() + "Response";
			Document updateDoc = Factory.Document.createInstance(os, docClassName);
			ContentElementList contentList = Factory.ContentElement.createList();
			ContentTransfer contentTransfer = Factory.ContentTransfer.createInstance();
			contentTransfer.setCaptureSource(is);
			contentTransfer.set_RetrievalName(docTitle + ".xlsx");
			contentTransfer.set_ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			contentList.add(contentTransfer);

			updateDoc.set_ContentElements(contentList);
			updateDoc.checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION);
			Properties p = updateDoc.getProperties();
			p.putValue("DocumentTitle", docTitle);
			updateDoc.setUpdateSequenceNumber(null);
			updateDoc.save(RefreshMode.REFRESH);
			ReferentialContainmentRelationship rc = responseFolder.file(updateDoc, AutoUniqueName.AUTO_UNIQUE, docTitle,
					DefineSecurityParentage.DO_NOT_DEFINE_SECURITY_PARENTAGE);
			rc.save(RefreshMode.REFRESH);
		} catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
		} finally {
			if (bos != null) {
				bos.close();
			}
			if (is != null) {
				is.close();
			}
			if (stream != null) {
				stream.close();
			}
		}

	}

	private static Object getCharValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();

		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}
}
