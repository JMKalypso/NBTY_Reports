package com.nbty.plm.px.reports;

import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServletRequest;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;

public class AgileSession {

	public IAgileSession getSession() throws APIException {
		IAgileSession session = null;
		Map<Integer, String> params = new HashMap<Integer, String>();
		params.put(AgileSessionFactory.URL, "http://172.18.24.39:40701/Agile");
		params.put(AgileSessionFactory.USERNAME, "administrator");
		params.put(AgileSessionFactory.PASSWORD, "agile9");
		session = AgileSessionFactory.createSessionEx(params);
		return session;
	}

	public IAgileSession getSession(String url, String user, String pass) throws APIException {
		IAgileSession session = null;
		Map<Integer, String> params = new HashMap<Integer, String>();
		params.put(AgileSessionFactory.URL, url);
		params.put(AgileSessionFactory.USERNAME, user);
		params.put(AgileSessionFactory.PASSWORD, pass);
		session = AgileSessionFactory.createSessionEx(params);
		return session;
	}

	public IAgileSession getSession(HttpServletRequest request, String URL) throws APIException { 
		IAgileSession session = null;
		AgileSessionFactory factory = AgileSessionFactory.getInstance(URL);
		HashMap<Object,Object> params = new HashMap<Object,Object>(); 
		params.put(AgileSessionFactory.PX_REQUEST, request); 
		session = factory.createSession(params); 
		return session; 
	} 
	
	public IAgileSession getSession(Cookie[] cookies, String URL) throws Exception {
		IAgileSession session = null; 
		AgileSessionFactory factory = AgileSessionFactory.getInstance(URL); 
		Map<Integer, String> params = new HashMap<Integer, String>();
		String username = null; 
		String pwd = null; 
		for (int i = 0; i < cookies.length; i++) { 
			if (cookies[i].getName().equals("j_username")) username = cookies[i].getValue(); 
			else if (cookies[i].getName().equals("j_password")) pwd = cookies[i].getValue(); 
		} 
		params.put(AgileSessionFactory.PX_USERNAME, username); 
		params.put(AgileSessionFactory.PX_PASSWORD, pwd);
		
		session = factory.createSession(params); 
		return session; 
	} 
	
}
