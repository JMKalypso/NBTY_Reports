package com.nbty.plm.px.reports;

import java.util.HashMap;
import java.util.Map;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;

public class AgileSession {

	public IAgileSession getSession() throws APIException {
		IAgileSession session = null;
		Map<Integer, String> params = new HashMap<Integer, String>();
		params.put(AgileSessionFactory.URL, "https://agiledev.nbty.net/Agile");
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

	public static void main(String args[]) throws Exception {
		System.out.printf("%f%n", (float)24/(1000*1000));
//		IAgileSession session = new AgileSession().getSession("http://agilenbtytest.oracleoutsourcing.com/Agile", "administrator", "agile9");
//		System.out.println("Session established as: " + session.getCurrentUser().getName());
	}
	
}
