package com.nbty.plm.px.extbulkchgs;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.UUID;

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
import com.agile.api.INode;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ITransferOrder;
import com.agile.api.ITwoWayIterator;
import com.agile.api.ItemConstants;
import com.agile.api.QueryConstants;
import com.agile.api.TransferOrderConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;

public class ExtractBulkChanges implements IEventAction {
	
	private static final Logger logger = Logger.getLogger("ExtractBulkChangesPXLog");

	@Override
	public EventActionResult doAction(IAgileSession session, INode node,
			IEventInfo eventInfo) {
		
		OutputStream out = null;
		try {
			
			// Load properties
			Properties props = getProps();
			logger.info("Properties loaded.");
			
			IAdmin admin = session.getAdminInstance();
			IAgileClass cls = admin.getAgileClass("Bulk");
			
			// Query for Bulks
			IQuery query = (IQuery) session.createObject(IQuery.OBJECT_TYPE, cls);
			query.setCaseSensitive(false);
			// Get dates for criteria
			String fromDate = props.getProperty("fromDate");
			String toDate = props.getProperty("toDate");
			logger.info(fromDate);
			logger.info(toDate);
			query.setCriteria("[2002] between ('" + fromDate + "' , '" + toDate + "')");
			
			// Create the Excel file.
			XSSFWorkbook wb = new XSSFWorkbook();
			File outputFile = new File(System.getProperty("java.io.tmpdir") + File.separator + UUID.randomUUID().toString() + ".xlsx");
			out = new FileOutputStream(outputFile);
			logger.info("Excel file created: " + outputFile.getAbsolutePath().toString());
			
			// Fill the file with data
			doExport(query, "BulkChanges", session, wb);
			logger.info("Done exporting...");
			
			// Writing file 
			wb.write(out);
			logger.info("File written.");
			
			EmailUtils.sendEmail(props.getProperty("email.to"), props.getProperty("email.from"), props.getProperty("email.subject"),
						props.getProperty("email.messageBody"), outputFile.getAbsolutePath(), props.getProperty("outputFilename"),
						props.getProperty("email.username"), props.getProperty("email.password"), props);
			logger.info("Email sent.");
//			IUser user1 = (IUser)session.getObject(UserConstants.CLASS_USER, "jlozano_kalypso");
//			IUser[] users = new IUser[]{user1};
//			
//			List<IUser> col = Arrays.asList(users);
//			IChange agileObject = (IChange)session.getObject(IChange.OBJECT_TYPE, "C000012752");
//			
//			agileObject.send(users, "Comments");
			 
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.STRING, "Success"));
		} catch (APIException e) {
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.EXCEPTION, e));
		} catch (FileNotFoundException e) {
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.EXCEPTION, e));
		} catch (Exception e) {
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.EXCEPTION, e));
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
			}
		}
	}//session.sendNotification(agileObject, "users - User Send", col, true, "Comments");//session.sendMail(users, "Hello");
	
	private Properties getProps() throws Exception {
		Properties props = new Properties();
		InputStream propIn = ExtractBulkChanges.class.getResourceAsStream("/px/ExtractBulkChanges/ExtractBulkChanges.properties");
		props.load(propIn);
		propIn.close();
		return props;
	}
	
	
	public void doExport(IQuery query, String sheetName, IAgileSession session, XSSFWorkbook wb) throws Exception {
		logger.info("Exporting...");
		int rowIdx = 0;
		try {
			
			// Headers
			String [] headers = new String[]{
					"Bulk Oracle Item Number",
					"Bulk Creation Date",
					"MBR Item Number",
					"MBR Creation Date", //3
					"MBR Revision",
					"MBR Rev-ECO Number",
					"MBR ECO Originated Date", // 6
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
			int counter = 0;
			
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
					 	logger.info("Bulk Oracle Item Number: " + data[0]); 
					 	logger.info("Bulk Creation Date: " + data[1]);
					 	
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
					 			
					 			data[2] = itemBOM.getName(); // MBR Item Number
					 			data[3] = itemBOM.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); // MBR Creation Date
					 			data[4] = (String)rowBOM.getValue(ItemConstants.ATT_BOM_ITEM_REV);	 
					 			
					 			logger.info("MBR Complete Revision: " + data[4]);
					 			
					 			// Get revision 
					 			String [] revChange = data[4].trim().split(" +");
					 			if (revChange.length > 0) {
					 				data[4] = revChange[0]; // MBR Revision
					 				
					 				// Only care for ECOs
							 		if (revChange[0].trim().length() <= 0 || revChange[0].trim().length() > 2) {
							 			//This is not a revision, ignore
							 		} else {
							 			// This is a revision
							 			logger.info("MBR Revision: " + data[4]);
							 			logger.info("MBR Revision Change: " + revChange[1]);
						 				
						 				IChange ecoMBR = (IChange)session.getObject(ChangeConstants.CLASS_CHANGE_BASE_CLASS, revChange[1]);
						 				logger.info(ecoMBR.getAgileClass().toString());
						 				if (ecoMBR.getAgileClass().toString().equals(ExtractConstants.ECO_SUBCLASS)) {
						 					//ecoMBR = (IChange)session.getObject(ChangeConstants.CLASS_ECO, revChange[1]);
						 					logger.info("It is an ECO");
						 					// Get the ECO related to this revision
							 				data[5] = ecoMBR.getName(); // MBR Rev-ECO Number
							 				logger.info("MBR Rev-ECO Number: " + data[5]);
							 				data[6] = ecoMBR.getValue(ChangeConstants.ATT_COVER_PAGE_DATE_ORIGINATED).toString();
							 				logger.info("MBR ECO Originated Date: " + data[6]);
							 				
							 				// Find the ATO that sent this change to the ERP
							 				IQuery atoQuery = (IQuery) session.createObject(IQuery.OBJECT_TYPE, TransferOrderConstants.CLASS_ATO);
							 				atoQuery.setSearchType(QueryConstants.TRANSFER_ORDER_SELECTED_CONTENT);
							 				atoQuery.setRelatedContentClass(ChangeConstants.CLASS_ECO);
							 				atoQuery.setCaseSensitive(false);
							 				atoQuery.setCriteria(" [Selected Content.ECO.Cover Page.Number] contains '" + data[5]  + "' ");
							 				
							 				ITable resATO = atoQuery.execute();
							 				ITwoWayIterator atoIter = resATO.getTableIterator();
							 				
							 				while(atoIter.hasNext()) {
							 					IRow rowATO = (IRow)atoIter.next();
							 					ITransferOrder ato = (ITransferOrder) rowATO.getReferent();
							 					data[7] = ato.getValue(TransferOrderConstants.ATT_COVER_PAGE_FINAL_COMPLETE_DATE).toString();
							 					data[8] = ato.toString(); // ATO Number for MBR to ERP
							 					logger.info("ATO Number for MBR to ERP: " + data[8]); 
							 					
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
				if (headers[i].equals("Bulk Creation Date") || 
						headers[i].equals("MBR ECO Date to ERP") || 
						headers[i].equals("MBR Creation Date") ||
						headers[i].equals("MBR ECO Creation Date")) {
					cell.setCellStyle(dateCellStyle);
					// Work the dates
					Date myDate = tryParse(data[i]);
					cell.setCellValue(myDate.toString());	
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
