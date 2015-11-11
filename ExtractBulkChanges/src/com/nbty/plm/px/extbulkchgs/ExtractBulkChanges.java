package com.nbty.plm.px.extbulkchgs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
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
import com.agile.api.IAdmin;
import com.agile.api.IAgileClass;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ITwoWayIterator;
import com.agile.api.IUser;
import com.agile.api.ItemConstants;
import com.agile.api.UserConstants;
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
			
			// Create the Excel file.
			//XSSFWorkbook wb = null;
			File outputFile = new File(System.getProperty("java.io.tmpdir") + File.separator + UUID.randomUUID().toString() + ".xlsx");
			out = new FileOutputStream(outputFile);
			logger.info("Excel file created: " + outputFile.getAbsolutePath().toString());
			
			doExport(query, "BulkChanges", out);
			
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
			 
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.STRING, "Test"));
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
	
	
	public void doExport(IQuery query, String sheetName, OutputStream out) throws Exception {
		logger.info("Exporting...");
		XSSFWorkbook wb = null;
		int rowIdx = 1;
		try {
			
			// Headers
			String [] headers = new String[]{
					"Bulk Oracle Item Number",
					"Bulk Creation Date",
					"MBR Item Number",
					"MBR Revision",
					"MBR Rev-ECO Number",
					"MBR ECO Date to ERP",
					"ATO Number for MBR to ERP"
			};
			
			// Create Worksheet from Workbook
			wb = new XSSFWorkbook();
			XSSFSheet ws = wb.createSheet(sheetName);
			logger.info("Worksheet created.");
			// Add headers to Worksheet
			addRow(headers, ws, rowIdx);
			
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
			 	Map revisions = item.getRevisions();
			 	Set set = revisions.entrySet();
			 	Iterator it = set.iterator();
			 	
			 	// Iterate each revision
			 	while (it.hasNext()) {
			 		Map.Entry entry = (Map.Entry)it.next();
			 		String rev = (String)entry.getValue();
			 		item.setRevision(rev);
			 					 		
			 		data[0] = item.getName(); 												// Bulk Oracle Item Number
				 	data[1] = item.getValue(ItemConstants.ATT_PAGE_TWO_DATE01).toString(); 	// Bulk Creation Date
				 	logger.info("Bulk Oracle Item Number: " + data[0]); 
				 	logger.info("Bulk Creation Date: " + data[1]);
				 	
				 	
				 	ITable bomTable = item.getTable(ItemConstants.TABLE_BOM);
				 	logger.info(bomTable.size());
				 	ITwoWayIterator itBOM = bomTable.getTableIterator();
				 	while(itBOM.hasNext()) {
				 		IRow rowBOM = (IRow)itBOM.next();
				 		IItem itemBOM = (IItem)rowBOM.getReferent();
				 		String itemtypeBOM = (String)rowBOM.getValue(ItemConstants.ATT_BOM_ITEM_TYPE);
				 		logger.info(itemtypeBOM);
				 		if (itemtypeBOM.equals("Master Batch Records")) {
				 			data[2] = (String)row.getValue(ItemConstants.ATT_BOM_ITEM_TYPE); // MBR Item Number
				 			data[3] = (String)row.getValue(ItemConstants.ATT_BOM_ITEM_REV);	 // MBR Revision
				 			logger.info("MBR Item Number: " + data[2]);
				 			logger.info("MBR Revision: " + data[3]);
				 			
				 			// Get revision 
				 			String revBOM = itemBOM.getRevision();
				 			logger.info(revBOM);
				 			String [] revChange = revBOM.split(" ");
				 			logger.info(revChange.length);
				 			
				 			
				 			//data[4] =  // MBR Rev-ECO Number
				 		}
				 	}
			 	}
			 	
			 	
			 	// temp break
			 	break;
			 }
			
			
			
			/*
			XSSFCellStyle dateCellStyle = wb.createCellStyle();
			CreationHelper createHelper = wb.getCreationHelper();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
			
			for (int i = 0; i < result.getValue().length; i++) {

				ProductProposal productProposal = result.getValue()[i];

				// if there are Costs associated with proposal; add them
				Cost[] costs = productProposal.getCost();
				if (costs != null && costs.length > 0) {
					for (int costIdx = 0; costIdx < costs.length; costIdx++) {
						ws.createRow(rowIdx);
						populateProposalAttributes(ws, rowIdx, productProposal, dateCellStyle);
						populateCostAttributes(ws, rowIdx, costs[costIdx], dateCellStyle);
						rowIdx++;
					}
				}

				// if there are Revenues associated with proposal; add them
				Revenue[] revenues = productProposal.getRevenue();
				if (revenues != null && revenues.length > 0) {
					for (int revenueIdx = 0; revenueIdx < revenues.length; revenueIdx++) {
						ws.createRow(rowIdx);
						populateProposalAttributes(ws, rowIdx, productProposal, dateCellStyle);
						populateRevenueAttributes(ws, rowIdx, revenues[revenueIdx], dateCellStyle);
						rowIdx++;
					}
				}

				if ((costs == null || costs.length < 1) && (revenues == null || revenues.length < 1)) {
					// populate the attributes
					ws.createRow(rowIdx);
					populateProposalAttributes(ws, rowIdx, productProposal, dateCellStyle);
					rowIdx++;
				}

			}*/

			wb.write(out);

		} catch (Exception e1) {
			throw e1;
		} finally {
			if (wb != null) {
				wb.close();
			}
		}

	}

	private void addRow(String[] data, XSSFSheet ws, int rowIdx) {
		// Create row in Excel 
		XSSFRow row = ws.createRow(rowIdx++);

		// Iterate through all data in array
		for (int i = 0; i < data.length; i++) {
			row.createCell(i).setCellValue(data[i]);
		}
		
	}
}
