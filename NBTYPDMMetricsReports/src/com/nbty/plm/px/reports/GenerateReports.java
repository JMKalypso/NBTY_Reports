package com.nbty.plm.px.reports;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAdmin;
import com.agile.api.IAgileClass;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IItem;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ITransferOrder;
import com.agile.api.ITwoWayIterator;
import com.agile.api.ItemConstants;
import com.agile.api.QueryConstants;
import com.agile.api.TransferOrderConstants;

public class GenerateReports {
	
	private static final Logger logger = Logger.getLogger("ExtractBulkChangesPXLog");
	
	public void doAction(String reportType, String fromDate,
			String toDate, HttpServletRequest request, HttpServletResponse response, String email) {
		
		try {
			response.getWriter().append("\nGenerating...");
			// Load properties
			Properties props = getProps();
			logger.info("Properties loaded.");	
			
			// Get Agile session
			IAgileSession session = new AgileSession().getSession(
					props.getProperty(ExtractConstants.URL_PROPERTY),
					ExtractConstants.PXUSER, ExtractConstants.PXPASS);
			logger.info("Session started.");	
			
			if (reportType.equals(ExtractConstants.MBR_TYPE)) {
				logger.info("MBR Changes selected. Creating Report...");
				doMBRChangesReport(session, fromDate, toDate, props);
			} else if (reportType.equals(ExtractConstants.CUBULK_TYPE)) {
				logger.info("CU-Bulk Report selected. Creating Report...");
				doCUBulkReport(session, fromDate, toDate, props);
			} else if (reportType.equals(ExtractConstants.CU_TYPE)) {
				logger.info("CU Changes selected. Creating Report...");
				doCUChangesReport(session, fromDate, toDate, props);
			} else {
				response.getWriter().append("Not a valid report to generate.");
			}
			response.getWriter().append("\nE-mail sent.");
		} catch (APIException e) {
			try {
				e.printStackTrace(response.getWriter());
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		} catch (Exception e) {
			try {
				response.getWriter().append(e.getMessage());
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		}
	}
	
	private void doMBRChangesReport(IAgileSession session, String fromDate, String toDate, Properties props) {
		
		OutputStream out = null;
		try {
			IAdmin admin = session.getAdminInstance();
			IAgileClass cls = admin.getAgileClass("Bulk");
			
			// Query for Bulks
			IQuery query = (IQuery) session.createObject(IQuery.OBJECT_TYPE, cls);
			query.setCaseSensitive(false);
			query.setCriteria("[2002] between ('" + fromDate + "' , '" + toDate + "')");
			
			// Create the Excel file.
			XSSFWorkbook wb = new XSSFWorkbook();
			File outputFile = new File(System.getProperty("java.io.tmpdir") + File.separator + UUID.randomUUID().toString() + ".xlsx");
			out = new FileOutputStream(outputFile);
			logger.info("Excel file created: " + outputFile.getAbsolutePath().toString());
			
			
			// Fill the file with data
			buildMBRSheet(query, "BulkChanges", session, wb);
			logger.info("Done exporting...");
			
			// Writing file 
			wb.write(out);
			logger.info("File written.");
			
			// Send report via email
			EmailUtils.sendEmail(props.getProperty("email.to"), props.getProperty("email.from"), props.getProperty("email.subject"),
					props.getProperty("email.messageBody"), outputFile.getAbsolutePath(), "MBRChangesReport.xlsx",
					props.getProperty("email.username"), props.getProperty("email.password"), props);
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
			}
			
		}
	}
	
	
	private void doCUBulkReport(IAgileSession session, String fromDate, String toDate, Properties props) {
		
		OutputStream out = null;
		try {
			IAdmin admin = session.getAdminInstance();
			IAgileClass cls = admin.getAgileClass(ExtractConstants.BULK_SUBCLASS);
			
			// Query for Bulks
			IQuery query = (IQuery) session.createObject(IQuery.OBJECT_TYPE, cls);
			query.setCaseSensitive(false);
			query.setCriteria("[2002] between ('" + fromDate + "' , '" + toDate + "')");
			
			// Create the Excel file.
			XSSFWorkbook wb = new XSSFWorkbook();
			File outputFile = new File(System.getProperty("java.io.tmpdir") + File.separator + UUID.randomUUID().toString() + ".xlsx");
			out = new FileOutputStream(outputFile);
			logger.info("Excel file created: " + outputFile.getAbsolutePath().toString());
			
			
			// Fill the file with data
			buildBulkCuSheet(query, "Bulk-CU relationship", session, wb);
			logger.info("Done exporting...");
			
			// Writing file 
			wb.write(out);
			logger.info("File written.");
			
			// Send report via email
			EmailUtils.sendEmail(props.getProperty("email.to"), props.getProperty("email.from"), props.getProperty("email.subject"),
					props.getProperty("email.messageBody"), outputFile.getAbsolutePath(), "CuBulkRelationshipReport.xlsx",
					props.getProperty("email.username"), props.getProperty("email.password"), props);
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
			}
			
		}
	}
	
	private void doCUChangesReport(IAgileSession session, String fromDate, String toDate, Properties props) {
		
		OutputStream out = null;
		try {
			IAdmin admin = session.getAdminInstance();
			IAgileClass cls = admin.getAgileClass(ExtractConstants.CU_SUBCLASS);
			
			// Query for Bulks
			IQuery query = (IQuery) session.createObject(IQuery.OBJECT_TYPE, cls);
			query.setCaseSensitive(false);
			query.setCriteria("[2002] between ('" + fromDate + "' , '" + toDate + "')");
			
			// Create the Excel file.
			XSSFWorkbook wb = new XSSFWorkbook();
			File outputFile = new File(System.getProperty("java.io.tmpdir") + File.separator + UUID.randomUUID().toString() + ".xlsx");
			out = new FileOutputStream(outputFile);
			logger.info("Excel file created: " + outputFile.getAbsolutePath().toString());
			
			
			// Fill the file with data
			buildCUSheet(query, "Consumer Unit Changes", session, wb);
			logger.info("Done exporting...");
			
			// Writing file 
			wb.write(out);
			logger.info("File written.");
			
			// Send report via email
			EmailUtils.sendEmail(props.getProperty("email.to"), props.getProperty("email.from"), props.getProperty("email.subject"),
					props.getProperty("email.messageBody"), outputFile.getAbsolutePath(), "CUChangesReport.xlsx",
					props.getProperty("email.username"), props.getProperty("email.password"), props);
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
			}
			
		}
	}



	private Properties getProps() throws Exception {
		Properties props = new Properties();
		InputStream propIn = GenerateReports.class.getResourceAsStream("/px/GenerateReports/GenerateReports.properties");
		props.load(propIn);
		propIn.close();
		return props;
	}
	
	
	public void buildMBRSheet(IQuery query, String sheetName, IAgileSession session, XSSFWorkbook wb) throws Exception {
		logger.info("Exporting...");
		int rowIdx = 0;
		try {
			
			// Headers
			String [] headers = new String[]{
					"Bulk Oracle Item Number",
					"Bulk Creation Date",
					"Bulk Revision",
					"MBR Item Number",
					"MBR Creation Date", //4
					"MBR Revision",
					"MBR Rev-ECO Number",
					"MBR ECO Originated Date", // 7
					"MBR ECO Date to ERP",
					"ATO Number for MBR to ERP"
			};
			
			// Create Worksheet from Workbook
			XSSFSheet ws = wb.createSheet(sheetName);
			logger.info("Worksheet created.");
			XSSFCellStyle dateCellStyle = wb.createCellStyle();
			CreationHelper createHelper = wb.getCreationHelper();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
			
			// Add headers to Worksheet
			addRow(headers, ws, rowIdx);
			rowIdx++;
			// Execute query for Bulks
			ITable results = query.execute();
			logger.info(results.size() + " bulks found.");
			
			// Build data 
			String [] data = new String[headers.length];
			ITwoWayIterator iter = results.getTableIterator();
			
			// Iterate results 
			while(iter.hasNext()) {
			 	IRow row = (IRow) iter.next();
			 	
			 	// Get all revisions
			 	IItem item = (IItem)row.getReferent();
			 	logger.info("Bulk number: " + item.getName());
			 	
			 	Map<?,?> revisions = item.getRevisions();
			 	
			 	Set<?> set = revisions.entrySet();
			 	Iterator<?> it = set.iterator();
			 	
			 	// Iterate each revision
			 	while (it.hasNext()) {
			 		Map.Entry<?,?> entry = (Map.Entry<?,?>)it.next();
			 		String rev = (String)entry.getValue();
			 		logger.info("Revision " + rev + " change: " + entry.getKey());
			 		
			 		// Only care for ECOs
			 		if (rev.trim().length() <= 0 || rev.trim().length() > 2) {
			 			//This is not a revision, ignore
			 		} else {
			 			// This is a revision
			 			item.setRevision(rev);
				 		
				 		data[0] = item.getName(); 												// Bulk Oracle Item Number
					 	data[1] = item.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); 	// Bulk Creation Date
					 	data[2] = item.getRevision();
					 	logger.info("Bulk Oracle Item Number: " + data[0]); 
					 	logger.info("Bulk Creation Date: " + data[1]);
					 	logger.info("Bulk Revision: " + data[2]);
					 	
					 	ITable bomTable = item.getTable(ItemConstants.TABLE_BOM);
					 	logger.info("BOM size: " + bomTable.size());
					 	
					 	ITwoWayIterator itBOM = bomTable.getTableIterator();
					 	
					 	// Iterate the BOM of the Bulk, looking for the MBR
					 	while(itBOM.hasNext()) {
					 		IRow rowBOM = (IRow)itBOM.next();
					 		IItem itemBOM = (IItem)rowBOM.getReferent();
					 		
					 		String itemtypeBOM = (String)itemBOM.getAgileClass().getName();
					 		
					 		// Is this the MBR?
					 		if (itemtypeBOM.equals(ExtractConstants.MBR_SUBCLASS)) {
					 			logger.info("BOM Item: " + itemBOM.getName());
					 			logger.info("BOM Item Type: " + itemtypeBOM);
					 			
					 			data[3] = itemBOM.getName(); // MBR Item Number
					 			data[4] = itemBOM.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); // MBR Creation Date
					 			data[5] = (String)rowBOM.getValue(ItemConstants.ATT_BOM_ITEM_REV);	 
					 			
					 			logger.info("MBR Complete Revision: " + data[5]);
					 			
					 			// Get revision 
					 			String [] revChange = data[5].trim().split(" +");
					 			if (revChange.length > 0) {
					 				data[5] = revChange[0]; // MBR Revision
					 				
					 				// Only care for ECOs
							 		if (revChange[0].trim().length() <= 0 || revChange[0].trim().length() > 2) {
							 			//This is not a revision, ignore
							 		} else {
							 			// This is a revision
							 			logger.info("MBR Revision: " + data[5]);
							 			logger.info("MBR Revision Change: " + revChange[1]);
						 				
						 				IChange ecoMBR = (IChange)session.getObject(ChangeConstants.CLASS_CHANGE_BASE_CLASS, revChange[1]);
						 				logger.info(ecoMBR.getAgileClass().toString());
						 				if (ecoMBR.getAgileClass().toString().equals(ExtractConstants.ECO_SUBCLASS)) {
						 					//ecoMBR = (IChange)session.getObject(ChangeConstants.CLASS_ECO, revChange[1]);
						 					logger.info("It is an ECO");
						 					// Get the ECO related to this revision
							 				data[6] = ecoMBR.getName(); // MBR Rev-ECO Number
							 				logger.info("MBR Rev-ECO Number: " + data[6]);
							 				data[7] = ecoMBR.getValue(ChangeConstants.ATT_COVER_PAGE_DATE_ORIGINATED).toString();
							 				logger.info("MBR ECO Originated Date: " + data[7]);
							 				
							 				// Find the ATO that sent this change to the ERP
							 				IQuery atoQuery = (IQuery) session.createObject(IQuery.OBJECT_TYPE, TransferOrderConstants.CLASS_ATO);
							 				atoQuery.setSearchType(QueryConstants.TRANSFER_ORDER_SELECTED_CONTENT);
							 				atoQuery.setRelatedContentClass(ChangeConstants.CLASS_ECO);
							 				atoQuery.setCaseSensitive(false);
							 				atoQuery.setCriteria(" [Selected Content.ECO.Cover Page.Number] contains '" + data[6]  + "' ");
							 				
							 				ITable resATO = atoQuery.execute();
							 				ITwoWayIterator atoIter = resATO.getTableIterator();
							 				
							 				while(atoIter.hasNext()) {
							 					IRow rowATO = (IRow)atoIter.next();
							 					ITransferOrder ato = (ITransferOrder) rowATO.getReferent();
							 					data[8] = ato.getValue(TransferOrderConstants.ATT_COVER_PAGE_FINAL_COMPLETE_DATE).toString();
							 					data[9] = ato.toString(); // ATO Number for MBR to ERP
							 					logger.info("ATO Number for MBR to ERP: " + data[9]); 
							 					
							 					// Add data to Worksheet
							 					addRow(data, ws, rowIdx, dateCellStyle, headers);
							 					rowIdx++;
							 					data = new String[headers.length];
							 					
							 					
							 					break;
							 				}
						 				} 
							 		}
					 				
					 			} // rev CLOSE
					 			
					 		} // THIS IS A MBR
					 	} // BULK BOM
			 		} // ECO Revision
			 	} // EACH REVISION
			 	
			 }

		} catch (Exception e1) {
			throw e1;
		} 
	}

	public void buildBulkCuSheet(IQuery query, String sheetName, IAgileSession session, XSSFWorkbook wb) throws Exception {
		logger.info("Exporting...");
		int rowIdx = 0;
		try {
			
			// Headers
			String [] headers = new String[]{
					"Bulk Oracle Item Number",
					"Bulk Creation Date",
					"Bulk Revision",
					"CU Oracle Number",
					"CU Revision"
			};
			
			// Create Worksheet from Workbook
			XSSFSheet ws = wb.createSheet(sheetName);
			logger.info("Worksheet created.");
			XSSFCellStyle dateCellStyle = wb.createCellStyle();
			CreationHelper createHelper = wb.getCreationHelper();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
			
			// Add headers to Worksheet
			addRow(headers, ws, rowIdx);
			rowIdx++;
			// Execute query for Bulks
			ITable results = query.execute();
			logger.info(results.size() + " bulks found.");
			
			// Build data 
			String [] data = new String[headers.length];
			ITwoWayIterator iter = results.getTableIterator();
			
			// Iterate results 
			while(iter.hasNext()) {
			 	IRow row = (IRow) iter.next();
			 	
			 	// Get all revisions
			 	IItem item = (IItem)row.getReferent();
			 	logger.info("Bulk number: " + item.getName());
			 	
			 	Map<?,?> revisions = item.getRevisions();
			 	
			 	Set<?> set = revisions.entrySet();
			 	Iterator<?> it = set.iterator();
			 	
			 	// Iterate each revision
			 	while (it.hasNext()) {
			 		Map.Entry<?,?> entry = (Map.Entry<?,?>)it.next();
			 		String rev = (String)entry.getValue();
			 		logger.info("Revision " + rev + " change: " + entry.getKey());
			 		
			 		// Only care for ECOs
			 		if (rev.trim().length() <= 0 || rev.trim().length() > 2) {
			 			//This is not a revision, ignore
			 		} else {
			 			// This is a revision
			 			item.setRevision(rev);
				 		
				 		data[0] = item.getName(); // Bulk Oracle Item Number
					 	data[1] = item.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); 	// Bulk Creation Date
					 	data[2] = item.getRevision(); // Bulk Revision
					 	
					 	logger.info("Bulk Oracle Item Number: " + data[0]); 
					 	logger.info("Bulk Creation Date: " + data[1]);
					 	logger.info("Bulk Revision: " + data[2]);
					 	
					 	ITable whereUsedTable = item.getTable(ItemConstants.TABLE_WHEREUSED);
					 	logger.info("Where Used size: " + whereUsedTable.size());
					 	
					 	ITwoWayIterator itWhereUsed = whereUsedTable.getTableIterator();
					 	
					 	// Iterate the Where Used table of the Bulk, looking for the CU
					 	while(itWhereUsed.hasNext()) {
					 		IRow rowWU = (IRow)itWhereUsed.next();
					 		IItem itemWU = (IItem)rowWU.getReferent();
					 		
					 		String itemtypeWU = (String)itemWU.getAgileClass().getName();
					 		
					 		// Is this the CU?
					 		if (itemtypeWU.equals(ExtractConstants.CU_SUBCLASS)) {
					 			logger.info("BOM Item: " + itemWU.getName());
					 			logger.info("BOM Item Type: " + itemtypeWU);
					 			
					 			data[3] = itemWU.getName(); // CU Item Number
					 			//data[3] = itemBOM.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); // MBR Creation Date
					 			data[4] = (String)rowWU.getValue(ItemConstants.ATT_WHERE_USED_ITEM_REV); // CU Revision
					 			
					 			logger.info("CU Complete Revision: " + data[4]);
					 			
					 			// Add data to Worksheet
			 					addRow(data, ws, rowIdx, dateCellStyle, headers);
			 					rowIdx++;
					 			
					 		} // THIS IS A CU
					 	} // BULK WHEREUSED
			 		} // ECO Revision
			 	} // EACH REVISION
			 	data = new String[headers.length];
			 }

		} catch (Exception e1) {
			throw e1;
		} 
	}
	
	public void buildCUSheet(IQuery query, String sheetName, IAgileSession session, XSSFWorkbook wb) throws Exception {
		logger.info("Exporting...");
		int rowIdx = 0;
		try {
			
			// Headers 
			String [] headers = new String[]{
					"CU Oracle Item Number",
					"CU Creation Date",
					"CU Revision",
					"CU Rev-ECO Number", // 3
					"CU Rev-ECO Date Originated",
					"AS400 Integration Item",
					"AS400 Integration Item Revision",
					"AS400 Integration Item Rev-ECO",
					"Bulk Item Number", // 8
					"Bulk Revision",
					"CTO Number", // 10
					"CTO Sent to AS400 Date"
			};
			
			// Create Worksheet from Workbook
			XSSFSheet ws = wb.createSheet(sheetName);
			logger.info("Worksheet created.");
			XSSFCellStyle dateCellStyle = wb.createCellStyle();
			CreationHelper createHelper = wb.getCreationHelper();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
			
			// Add headers to Worksheet
			addRow(headers, ws, rowIdx);
			rowIdx++;
			// Execute query for Bulks
			ITable results = query.execute();
			logger.info(results.size() + " Consumer Units found.");
			
			// Build data 
			String [] data = new String[headers.length];
			ITwoWayIterator iter = results.getTableIterator();
			
			// Iterate results 
			while(iter.hasNext()) {
			 	IRow row = (IRow) iter.next();
			 	
			 	// Get the items
			 	IItem item = (IItem)row.getReferent();
			 	logger.info("CU number: " + item.getName());
			 	
			 	// Get all revisions
			 	Map<?,?> revisions = item.getRevisions();
			 	
			 	Set<?> set = revisions.entrySet();
			 	Iterator<?> it = set.iterator();
			 	
			 	// Iterate each revision
			 	while (it.hasNext()) {
			 		logger.info("Next revision...");
			 		Map.Entry<?,?> entry = (Map.Entry<?,?>)it.next();
			 		String rev = (String)entry.getValue();
			 		logger.info("Revision " + rev + " change: " + entry.getKey());
			 		
			 		// Only care for ECOs
			 		if (rev.trim().length() <= 0 || rev.trim().length() > 2) {
			 			//This is not a revision, ignore
			 		} else {
			 			
			 			IChange change = (IChange)session.getObject(IChange.OBJECT_TYPE, entry.getKey().toString());
			 			logger.info("Change class: " + change.getAgileClass());
			 			logger.info("Change: " + change.getName());
			 			if (change.getAgileClass().getName().equals(ExtractConstants.ECO_SUBCLASS)){
				 			// This is a revision
				 			// Set to this revision on both items
			 				if (!item.getRevision().equals(rev)) {
			 					item.setRevision(rev);
			 				}
				 			logger.info("Revision " + rev + " set in CU.");
				 			
				 			IItem as400 = (IItem)session.getObject(IItem.OBJECT_TYPE, item.getName() + ".AS400");
						 	
						 	// If there is no AS400 Integration Item, then continue
						 	if (as400 == null) {
						 		logger.info("AS400 Integration Item doesn't exist.");
						 	} else {
						 		logger.info("AS400 Item: " + as400);
						 	}
				 			
				 			try {
				 				as400.setRevision(rev);
				 				logger.info("Revision " + rev + " set in AS400 Item.");
				 			} catch (APIException e) {
				 				logger.info("AS400 Item doesn't have this revision. ");
				 				as400 = null;
				 			} catch(Exception e) {
				 				logger.info(e.getMessage());
				 			}
				 			
					 		data[0] = item.getName(); // CU Oracle Item Number
						 	data[1] = item.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); 	// CU Creation Date
						 	
						 	logger.info("CU Oracle Item Number: " + data[0]); 
						 	logger.info("CU Creation Date: " + data[1]);
						 	
						 	data[2] = item.getRevision(); // CU Revision
						 	data[3] = entry.getKey().toString(); // CU Revision Change
						 	
						 	IChange revECO = item.getChange();
						 	data[4] = revECO.getValue(ChangeConstants.ATT_COVER_PAGE_DATE_ORIGINATED).toString(); // CU Rev-ECO Date Originated
						 	
						 	// Get BOM to look for Bulk
			 				ITable bomTable = item.getTable(ItemConstants.TABLE_BOM);
						 	logger.info("BOM size: " + bomTable.size());
						 	
						 	ITwoWayIterator itBOM = bomTable.getTableIterator();
			 				
						 	// Iterate the BOM of the CU, looking for the Bulk
						 	while(itBOM.hasNext()) {
						 		IRow rowBOM = (IRow)itBOM.next();
						 		IItem itemBOM = (IItem)rowBOM.getReferent();
						 		
						 		String itemtypeBOM = (String)itemBOM.getAgileClass().getName();
						 		
						 		// Is this the Bulk?
						 		if (itemtypeBOM.equals(ExtractConstants.BULK_SUBCLASS)) {
						 			logger.info("BOM Item Type: " + itemtypeBOM);
						 			
						 			data[8] = itemBOM.getName(); // Bulk Item Number 
						 			data[9] = itemBOM.getRevision();
						 			
						 			logger.info("Bulk Number: " + data[8]);
						 			logger.info("Bulk Revision: " + data[9]);
						 			
						 			break;
						 		} 
							 		
						 	} // CU BOM
						 	
						 	
						 	data[5] = as400 == null ? "" : as400.getName(); // AS400 Integration Item 
						 	data[6] = as400 == null ? "" : as400.getRevision(); // AS400 Integration Item Revision
						 	
						 	logger.info("CU Revision: " + data[2]);
						 	logger.info("CU Revision Change: " + data[3]);
						 	logger.info("ECO Rev-Change Date Originated: " + data[4]);
						 	logger.info("AS400 Integration Item: " + data[5]);
						 	logger.info("AS400 Integration Item Revision: " + data[6]);
						 	
						 	if (as400 != null) {
						 		// Get all changes within this revision that were sent to AS400 
						 		ITable as400Changes = as400.getTable(ItemConstants.TABLE_CHANGEHISTORY);
						 		ITwoWayIterator chgIt = as400Changes.getTableIterator();
						 		logger.info("Change History... ");
						 		
						 		while (chgIt.hasNext()) {
						 			IRow rowChange = (IRow)chgIt.next();
						 			String rowRev = rowChange.getCells()[1].toString();
						 			logger.info("Row revision: " + rowRev);
						 			
						 			if (rowRev.equals(as400.getRevision())) {
						 				IChange as400Change = (IChange)rowChange.getReferent();
						 				
						 				data[7] = as400Change.getName(); // AS400 Integration Item Revision Change 
						 				logger.info("AS400 Item Change: " + data[7]);
						 				
									 	// Find the CTO that sent this change to the ERP
						 				IQuery ctoQuery = (IQuery) session.createObject(IQuery.OBJECT_TYPE, TransferOrderConstants.CLASS_CTO);
						 				ctoQuery.setSearchType(QueryConstants.TRANSFER_ORDER_SELECTED_CONTENT);
						 				ctoQuery.setRelatedContentClass(ChangeConstants.CLASS_CHANGE_BASE_CLASS);
						 				ctoQuery.setCaseSensitive(false);
						 				ctoQuery.setCriteria(" [Selected Content.Changes.Cover Page.Number] contains '" + data[7]  + "' ");
						 				
						 				ITable resCTO = ctoQuery.execute();
						 				ITwoWayIterator ctoIter = resCTO.getTableIterator();
						 				logger.info(resCTO.size() + " Transfer Orders found. ");
						 				
						 				while(ctoIter.hasNext()) {
						 					IRow rowCTO = (IRow)ctoIter.next();
						 					ITransferOrder cto = (ITransferOrder) rowCTO.getReferent();
						 					
						 					data[10] = cto.toString(); // CTO Number for CU to ERP
						 					data[11] = cto.getValue(TransferOrderConstants.ATT_COVER_PAGE_FINAL_COMPLETE_DATE) != null ? 
						 							cto.getValue(TransferOrderConstants.ATT_COVER_PAGE_FINAL_COMPLETE_DATE).toString() : ""; // CTO Sent to AS400 Date
						 					logger.info("CTO Number for MBR to ERP: " + data[10]);
						 					logger.info("CTO Complete: " + data[11]);
						 					
										 	// Add data to Worksheet
											addRow(data, ws, rowIdx, dateCellStyle, headers);
											rowIdx++;
						 				}
						 				
						 				if (data[10] == null) {
						 					logger.info("No CTO found.");
						 					// Add data to Worksheet
											addRow(data, ws, rowIdx, dateCellStyle, headers);
											rowIdx++;
						 				}

						 			}
						 		}
				 				
						 	}
			 			}

			 		} // ECO Revision
			 		data = new String[headers.length];
			 	} // EACH REVISION
			 }

		} catch (Exception e1) {
			logger.info(e1.getMessage()); 
		} 
	}
	
	private void addRow(String[] data, XSSFSheet ws, int rowIdx,
			XSSFCellStyle dateCellStyle, String[] headers) {

		try {
			// Create row in Excel 
			XSSFRow row = ws.createRow(rowIdx);

			// Iterate through all data in array
			for (int i = 0; i < data.length; i++) {
				XSSFCell cell = row.createCell(i);
				cell.setCellValue(data[i]);
				
				// Is this a date value?
				if (headers[i].contains("Date")) {
					cell.setCellStyle(dateCellStyle);
					// Work the dates
					Date myDate = tryParse(data[i]);
					cell.setCellValue(myDate);	
				}
			}
		} catch (Exception e) {
			logger.info(e.getMessage());
		}
		
	}

	// Use this to add a row into an Excel Worksheet
	private void addRow(String[] data, XSSFSheet ws, int rowIdx) {
		// Create row in Excel 
		XSSFRow row = ws.createRow(rowIdx);

		// Iterate through all data in array
		for (int i = 0; i < data.length; i++) {
			row.createCell(i).setCellValue(data[i]);
		}
		
	}
	
	// Use this to parse dates into a single format
	private Date tryParse(String dateString)
	{
		String[] formatStrings = {"yyyy-MM-dd", "EEE MMM dd HH:mm:ss z yyyy"};
	    for (String formatString : formatStrings)
	    {
	        try
	        {
	        	Date myDate = new SimpleDateFormat(formatString).parse(dateString);
	        	logger.info(dateString + " formated using " + formatString);
	        	SimpleDateFormat myDateFormat = new SimpleDateFormat("MM/dd/yyyy");
	        	myDateFormat.format(myDate);
	            return myDate;
	        }
	        catch (ParseException e) {
	        	// Do nothing...
	        }
	    }
	    logger.info("Parse format not found.");
	    return null;
	}
}
