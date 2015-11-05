package com.nbty.plm.px.extbulkchgs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Properties;
import java.util.UUID;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.INode;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ITwoWayIterator;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;

public class ExtractBulkChanges implements IEventAction {
	
	private static final Logger logger = Logger.getLogger("ExtractBulkChanges.log");

	@Override
	public EventActionResult doAction(IAgileSession session, INode node,
			IEventInfo eventInfo) {
		OutputStream out = null;
		try {
			
			// Load properties
			Properties props = getProps();
			
			IQuery query = (IQuery) session.createObject(IQuery.OBJECT_TYPE, ItemConstants.CLASS_PART);
			query.setCaseSensitive(false);
			query.setCriteria("[Item Type] Equal To 'Bulk' AND ");
			
			// Create the Excel file.
			XSSFWorkbook wb = null;
			File outputFile = new File(System.getProperty("java.io.tmpdir") + File.separator + UUID.randomUUID().toString() + "xlsm");
			out = new FileOutputStream(outputFile);
			doExport(query, props.getProperty("templateSheetname"), out);
			
			
			
			 EmailUtils.sendEmail(props.getProperty("email.to"), props.getProperty("email.from"), props.getProperty("email.subject"),
						props.getProperty("email.messageBody"), outputFile.getAbsolutePath(), props.getProperty("outputFilename"),
						props.getProperty("email.username"), props.getProperty("email.password"), props);
				
			 
			return new EventActionResult(eventInfo, new ActionResult(ActionResult.STRING, "Test " + logger.getName()));
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
	}
	
	private Properties getProps() throws Exception {
		Properties props = new Properties();
		InputStream propIn = ExtractBulkChanges.class.getResourceAsStream("/px/ExtractBulkChanges/ExtractBulkChanges.properties");
		props.load(propIn);
		propIn.close();
		return props;
	}
	
	
	public void doExport(IQuery query, String sheetName, OutputStream out) throws Exception {

		XSSFWorkbook wb = null;
		try {
			
			ITable results = query.execute();
			logger.info(results.size());
			
			ITwoWayIterator iter = results.getTableIterator();
			 while(iter.hasNext()) {
			 	IRow row = (IRow) iter.next();
			 	String itemNumber = row.getCell(ItemConstants.ATT_TITLE_BLOCK_NUMBER).toString();
			 	logger.info(itemNumber);
			 	
			 	
			 	
			 	// temp break
			 	break;
			 }
			
			wb = new XSSFWorkbook();
			XSSFSheet ws = wb.getSheet(sheetName);
			
			/*
			XSSFCellStyle dateCellStyle = wb.createCellStyle();
			CreationHelper createHelper = wb.getCreationHelper();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
			int rowIdx = 2;
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
}
