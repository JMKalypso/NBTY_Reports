package com.nbty.plm.px.url;

import java.io.IOException;
import java.io.PrintWriter;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.validator.routines.EmailValidator;

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
		// Get the params
		String reportType = request.getParameter("reports");
		String fromDate = request.getParameter("fromDate");
		String toDate = request.getParameter("toDate");
		String email = request.getParameter("email");
		
		/*
		 * Server not configured to enable the use of cookies
		 */
		// Get the cookies
		//Cookie[] cookies = request.getCookies();
		//printCookies(cookies,response);
		if (validateEmail(email)) {
			response.getWriter().append("Sending...");
			GenerateReports generator = new GenerateReports(); 
			generator.doAction(reportType, fromDate, toDate, request, response, email);
		} else {
			request.setAttribute("errorMessage", "Not a valid e-mail address.");
			request.setAttribute("fromDate", fromDate);
			request.setAttribute("toDate", toDate);
			request.setAttribute("email", email);
			
			request.getRequestDispatcher("input.jsp").forward(request, response);
			//response.sendRedirect("input.jsp");
			//response.getWriter().append("Not a valid e-mail address.");
		}
	}
	
	@SuppressWarnings("unused")
	private void printCookies(Cookie[] cookies, HttpServletResponse response) throws IOException {
		response.setContentType("text/html");
		PrintWriter writer = response.getWriter();
		writer.append("<html><body>");
		for (int i=0;i<cookies.length;i++) {
			response.getWriter().append(cookies[i].getName() + ":" + cookies[i].getValue() + "<br>");
		}
		writer.append("</body></html>");
	}
	
	private boolean validateEmail(String adress) {
		return EmailValidator.getInstance().isValid(adress);
	}

}
