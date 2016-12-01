package com.jun.test;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class TestClass {

	private String driver = "oracle.jdbc.driver.OracleDriver";
	private String url = "jdbc:oracle:thin:@10.0.5.152:1521/jknc";
	private String user = "jknc02";
	private String password = "jknc02";
	private Connection conn;

	public Connection getConnection() throws SQLException, ClassNotFoundException {
		if (conn == null) {
			Class.forName(driver);
			return DriverManager.getConnection(url, user, password);
		} else {
			return conn;
		}
	}

	public static void main(String[] args) throws ClassNotFoundException, SQLException {
//		TestClass test = new TestClass();
//		System.out.println(test.getConnection());
		List l = new ArrayList();
		l.add(13);
		l.add(4);
		l.add(15);
		l.add(2);
		System.out.println(l);
		Collections.sort(l, Collections.reverseOrder());
		System.out.println(l);
	}
}
