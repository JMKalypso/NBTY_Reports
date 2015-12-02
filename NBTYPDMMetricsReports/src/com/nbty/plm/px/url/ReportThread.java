package com.nbty.plm.px.url;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.nbty.plm.px.reports.GenerateReports;

public class ReportThread implements Runnable {

	private String reportType;
	private String fromDate;
	private String toDate;
	private HttpServletRequest request;
	private HttpServletResponse response;
	private String email;
	
	/**
	 * @param reportType
	 * @param fromDate
	 * @param toDate
	 * @param request
	 * @param response
	 * @param email
	 */
	public ReportThread(String reportType, String fromDate, String toDate, HttpServletRequest request,
			HttpServletResponse response, String email) {
		this.reportType = reportType;
		this.fromDate = fromDate;
		this.toDate = toDate;
		this.request = request;
		this.response = response;
		this.email = email;
	}

	@Override
	public void run() {
		// Call the procedure that will generate the reports
		GenerateReports generator = new GenerateReports(); 
		generator.doAction(reportType, fromDate, toDate, request, response, email);
	}

}
