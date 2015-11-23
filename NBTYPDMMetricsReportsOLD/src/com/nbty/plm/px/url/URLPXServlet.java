package com.nbty.plm.px.url;

import java.io.IOException;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.nbty.plm.px.reports.GenerateReports;

/**
 * Servlet implementation class URLPXServlet
 */
@WebServlet("/URLPXServlet")
public class URLPXServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public URLPXServlet() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		String reportType = request.getParameter("reports");
		String fromDate = request.getParameter("fromDate");
		String toDate = request.getParameter("toDate");
		
		response.getWriter().append("Sending...");
		
		GenerateReports generator = new GenerateReports(); 
		generator.doAction(reportType, fromDate, toDate);
		
		response.getWriter().append("Email Sent.");
		
		//doGet(request, response);
	}

}
