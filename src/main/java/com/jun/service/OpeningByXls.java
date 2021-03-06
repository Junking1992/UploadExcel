package com.jun.service;

import java.math.BigDecimal;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class OpeningByXls extends ProgressUtil {
	Map pubdocMap = new HashMap();
	Map defDocMap = new HashMap();
	Map productMap = new HashMap();
	Map storeMap = new HashMap();
	Map jarNumByStoreMap = new HashMap();
	Map jarCubageByStoreMap = new HashMap();
	String areaNo = "";
	String buildingNo = "";
	String storeNo = "";
	String status = "";

	public void initIquantityDate(List<Map<String, String>> xlsAryList) throws Exception {
		// 设置进度总数
		setAllCount(xlsAryList.size());
		try {
			conn = getConnection();
			conn.setAutoCommit(false);

			// 取得片区栋库编码
			pubdocMap = initPubDocMap();
			// 取得年份轮次香型等级备注酒编码
			defDocMap = initDefDocMap();
			// 取得库号CODE
			//storeMap = initStoreMap();
			// 取得产品名称
			productMap = initProductMap();
			// 有坛号的酒记录
			Map jarMap = new HashMap();
			//记录片区MAP
			Map areaMap = new HashMap();

			// 主逻辑 先干掉整个05片区资料
			// String sql = "update mtws_iquantity set dr=2 where dr=0 and
			// pk_area='1001A4100000000015O2'";
			// update(sql);

			for (int i = 0; i < xlsAryList.size(); i++) {
				Map map = (Map) xlsAryList.get(i);
				// 错误行号
				errorRowNum = getStrMapValue(map, "RN");
				String jarMemo = getStrMapValue(map, "Z");
				if (jarMemo.equals(""))
					continue;
				areaNo = "";
				buildingNo = "";
				storeNo = "";
				Map<String, String> abstoreMape = getABStore(getStrMapValue(map, "C"));
				String pk_area = getStrMapValue(abstoreMape, areaNo);
				if(getStrMapValue(areaMap, areaNo).equals("")){
					areaMap.put(areaNo, areaNo);
					initJarNumByStoreMap();
					initJarCubageByStoreMap();
				}
				String pk_store = getStrMapValue(abstoreMape, storeNo);
				String pk_building = getStrMapValue(abstoreMape, buildingNo);
				String tubCode = getTubCode(getStrMapValue(map, "D"), getStrMapValue(map, "J"), getStrMapValue(map, "E"), getStrMapValue(map, "F"));
				String iyear = getIyear(getStrMapValue(map, "E"));
				String oriDate = getStrMapValue(map, "F");
				String inDate = getInDate(getStrMapValue(map, "G"), getStrMapValue(map, "H"), getStrMapValue(map, "I"));
				String pdType = getPdType(getStrMapValue(map, "J"));
				String sweetype = getsweetType(getStrMapValue(map, "K"));
				String iturns = getIturns(getStrMapValue(map, "L"));
				BigDecimal qtySum = getDecMapValue(map, "M");
				BigDecimal ptySum = getDecMapValue(map, "N");
				String bzMessage = getBzMessage(getStrMapValue(map, "O"));
				BigDecimal exQtySum = getDecMapValue(map, "P");
				// BigDecimal tubQty = getDecMapValue(map, "Q");
				String grade = getGrade(getStrMapValue(map, "R"));
				String memo = getStrMapValue(map, "S");
				String exDate = getStrMapValue(map, "T");
				String oriTransMemo = getStrMapValue(map, "U");
				String memo2 = getStrMapValue(map, "V");
				String direction = getStrMapValue(map, "W");
				String dirMemo = getStrMapValue(map, "X");
				String updateDate = getStrMapValue(map, "Y");
				List jarList = getJarList(jarMemo);
				int ptySumInt = Integer.parseInt(ptySum.toString());
				int jarnum = ptySumInt > jarList.size() ? jarList.size() : ptySumInt;
				if (jarnum < 1)
					throw new Exception("Excel汇入的库[" + storeNo + "]桶[" + tubCode + "]资料的总坛数[" + jarnum + "]不可为空!");
				// 检查坛数是否合理
				BigDecimal jarNumByStore = getDecMapValue(jarNumByStoreMap, areaNo + buildingNo + storeNo);
				if (ptySum.compareTo(jarNumByStore) > 0)
					throw new Exception("本次期初坛数[" + ptySum + "]不可大于库总坛数[" + jarNumByStore + "]!");
				BigDecimal avgJar = qtySum.divide(new BigDecimal(jarnum), 2, BigDecimal.ROUND_HALF_UP);
				BigDecimal qtyJar = new BigDecimal("0");
				String pk_jar = "";
				// if(jarList.size()!=Integer.parseInt(ptySum.toString())){
				// throw new
				// Exception("Excel汇入的库["+storeNo+"]桶["+tubCode+"]资料的总坛数["+ptySum+"]与每桶对应坛数["+jarList.size()+"]不一致!");
				// }
				// 标准坛容
				BigDecimal jarCubage = getDecMapValue(jarCubageByStoreMap, areaNo + buildingNo + storeNo);
				for (int j = 0; j < jarnum; j++) {
					pk_jar = getPk_jar(jarList.get(j).toString());
					if(pk_jar == null || "".equals(pk_jar)){
						throw new Exception("坛号[" + jarList.get(j).toString() + "]基础档案里不存在!");
					}
					jarMap.put(jarList.get(j).toString(), pk_jar);
					qtySum = qtySum.subtract(avgJar);
					qtyJar = avgJar;
					String jarState = "满坛";
					String statusCode = status;
					if (avgJar.compareTo(jarCubage) < 0) {
						jarState = "半坛";
					}
					if (exQtySum.compareTo(avgJar) >= 0) {
						exQtySum = exQtySum.subtract(avgJar);
						qtyJar = new BigDecimal("0");
						jarState = "空坛";
						statusCode = "00";
					}
					// 最后一笔单独计算
					if (j == jarnum - 1) {
						qtyJar = qtyJar.add(qtySum).subtract(exQtySum);
						avgJar = avgJar.add(qtySum);
						if (qtyJar.compareTo(new BigDecimal("0")) <= 0) {
							jarState = "空坛";
							statusCode = "00";
						} else if (qtyJar.compareTo(jarCubage) >= 0)
							jarState = "满坛";
						else
							jarState = "半坛";
					}
					String updateSql = "update mtws_iquantity set dr=1 where nvl(dr,0)=0 and pk_jar='"+pk_jar+"'";
					update(updateSql);
					String insertSql = "insert into mtws_iquantity (pk_iquantity,pk_product,pk_area,pk_building,pk_store,"
							+ "tubcode,pk_jar,iyear,iturns,sweetype,bzmessage,iquertity,pk_measure,jarstate,def10,def14,def15,"
							+ "def16,def17,def18,def20,ts,dr)" + "values ('" + pk_jar + "','" + getProduct(pdType)
							+ "','" + pk_area + "','" + pk_building + "','" + pk_store + "'," + "'" + tubCode + "','"
							+ pk_jar + "','" + iyear + "','" + iturns + "','" + sweetype + "','" + bzMessage + "',"
							+ qtyJar + ",'1001A41000000000034A','" + jarState + "','" + statusCode + "'," + avgJar
							+ ",'" + pdType + "','" + oriDate + "','" + inDate + "'," + "'" + exDate + "','" + grade
							+ "',to_char(sysdate,'yyyy-mm-dd hh24:mi:ss')," + "'0')";
					
					create(insertSql);
				}
				// 进度增长
				addCount();
			}

			for (int i = 0; i < xlsAryList.size(); i++) {
				Map map = (Map) xlsAryList.get(i);
				// 错误行号
				errorRowNum = getStrMapValue(map, "RN");
				String jarMemo = getStrMapValue(map, "Z");
				if (!jarMemo.equals(""))
					continue;
				areaNo = "";
				buildingNo = "";
				storeNo = "";
				Map<String, String> abstoreMape = getABStore(getStrMapValue(map, "C"));
				String pk_area = getStrMapValue(abstoreMape, areaNo);
				if(getStrMapValue(areaMap, areaNo).equals("")){
					areaMap.put(areaNo, areaNo);
					initJarNumByStoreMap();
					initJarCubageByStoreMap();
				}
				String pk_store = getStrMapValue(abstoreMape, storeNo);
				String pk_building = getStrMapValue(abstoreMape, buildingNo);
				String tubCode = getTubCode(getStrMapValue(map, "D"), getStrMapValue(map, "J"),
						getStrMapValue(map, "E"), getStrMapValue(map, "F"));
				String iyear = getIyear(getStrMapValue(map, "E"));
				String oriDate = getStrMapValue(map, "F");
				String inDate = getInDate(getStrMapValue(map, "G"), getStrMapValue(map, "H"), getStrMapValue(map, "I"));
				String pdType = getPdType(getStrMapValue(map, "J"));
				String sweetype = getsweetType(getStrMapValue(map, "K"));
				String iturns = getIturns(getStrMapValue(map, "L"));
				BigDecimal qtySum = getDecMapValue(map, "M");
				BigDecimal ptySum = getDecMapValue(map, "N");
				// 检查坛数是否合理
				BigDecimal jarNum = getDecMapValue(jarNumByStoreMap, areaNo + buildingNo + storeNo);
				if (ptySum.compareTo(jarNum) > 0)
					throw new Exception("本次期初坛数[" + ptySum + "]不可大于库总坛数[" + jarNum + "]!");
				String bzMessage = getBzMessage(getStrMapValue(map, "O"));
				BigDecimal exQtySum = getDecMapValue(map, "P");
				// BigDecimal tubQty = getDecMapValue(map, "Q");
				String grade = getGrade(getStrMapValue(map, "R"));
				String memo = getStrMapValue(map, "S");
				String exDate = getStrMapValue(map, "T");
				String oriTransMemo = getStrMapValue(map, "U");
				String memo2 = getStrMapValue(map, "V");
				String direction = getStrMapValue(map, "W");
				String dirMemo = getStrMapValue(map, "X");
				String updateDate = getStrMapValue(map, "Y");
				// 计算平均坛重量
				BigDecimal avgJar = qtySum.divide(ptySum, 2, BigDecimal.ROUND_HALF_UP);
				BigDecimal qtyJar = new BigDecimal("0");
				int jarInt = 1;
				String pk_jar = "";
				// 标准坛容
				BigDecimal jarCubage = getDecMapValue(jarCubageByStoreMap, areaNo + buildingNo + storeNo);
				int ptySumInt = Integer.parseInt(ptySum.toString());
				for (int j = 0; j < ptySumInt; j++) {
					String jarNo = areaNo + buildingNo + storeNo
							+ "0000".substring(0, 4 - String.valueOf(jarInt).length()) + jarInt;
					while (!getStrMapValue(jarMap, jarNo).equals("")) {
						jarInt++;
						if (jarInt > Integer.parseInt(jarNum.toString()))
							throw new Exception("系统产生的坛号[" + ptySum + "]不可大于库总坛数[" + jarNum + "]!");
						jarNo = areaNo + buildingNo + storeNo + "0000".substring(0, 4 - String.valueOf(jarInt).length())
								+ jarInt;
					}
					pk_jar = getPk_jar(jarNo);

					qtySum = qtySum.subtract(avgJar);
					qtyJar = avgJar;
					String statusCode = status;
					String jarState = "满坛";
					if (avgJar.compareTo(jarCubage) < 0) {
						jarState = "半坛";
					}
					if (exQtySum.compareTo(avgJar) >= 0) {
						exQtySum = exQtySum.subtract(avgJar);
						qtyJar = new BigDecimal("0");
						jarState = "空坛";
						statusCode = "00";
					}
					// 最后一笔单独计算
					if (j == ptySumInt - 1) {
						qtyJar = qtyJar.add(qtySum).subtract(exQtySum);
						avgJar = avgJar.add(qtySum);
						if (qtyJar.compareTo(new BigDecimal("0")) <= 0) {
							jarState = "空坛";
							statusCode = "00";
						} else if (qtyJar.compareTo(jarCubage) >= 0)
							jarState = "满坛";
						else
							jarState = "半坛";
					}
					String updateSql = "update mtws_iquantity set dr=1 where nvl(dr,0)=0 and pk_jar='"+pk_jar+"'";
					update(updateSql);
					String insertSql = "insert into mtws_iquantity (pk_iquantity,pk_product,pk_area,pk_building,pk_store,"
							+ "tubcode,pk_jar,iyear,iturns,sweetype,bzmessage,iquertity,pk_measure,jarstate,def10,def14,def15,"
							+ "def16,def17,def18,def20,ts,dr)" + "values ('" + pk_jar + "','" + getProduct(pdType)
							+ "','" + pk_area + "','" + pk_building + "','" + pk_store + "'," + "'" + tubCode + "','"
							+ pk_jar + "','" + iyear + "','" + iturns + "','" + sweetype + "','" + bzMessage + "',"
							+ qtyJar + ",'1001A41000000000034A','" + jarState + "','" + statusCode + "'," + avgJar
							+ ",'" + pdType + "','" + oriDate + "','" + inDate + "'," + "'" + exDate + "','" + grade
							+ "',to_char(sysdate,'yyyy-mm-dd hh24:mi:ss')," + "'0')";
					create(insertSql);
				}
				// 进度增长
				addCount();
			}
			conn.commit();
		} catch (Exception e) {
			conn.rollback();
			throw new Exception("错误在：" + errorRowNum + "行." + e.toString());
		} finally {
			close(conn);
		}

	}

	private void initJarCubageByStoreMap() throws SQLException {
		// TODO 自动生成的方法存根
		// TODO 自动生成的方法存根
		// 取得05片区各库位总坛数
		if (areaNo.equals(""))
			areaNo = "XXXXX";
		String jarNumStoreSql = "select substr(code,1,9) as store ,jarcubage from mtws_jar where code like '" + areaNo
				+ "%' and dr=0 " + " group by substr(code,1,9),jarcubage ";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(jarNumStoreSql);
		while (rs.next()) {
			jarCubageByStoreMap.put(rs.getString("store"), rs.getString("jarcubage"));
		}
		Stmt.close();
		rs.close();
	}

	private Map initStoreMap() throws SQLException {
		// TODO 自动生成的方法存根
//		String sql = "select substr(code,6,4) as storeno,code from mtws_pubdoc where def2='2' and dr=0";
		String sql = "select substr(code,6,4) as storeno,code from mtws_pubdoc where dr=0 and name like '%库'";
		
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(sql);
		Map rstMap = new HashMap();
		while (rs.next()) {
			rstMap.put(rs.getString("storeno"), rs.getString("code"));
		}
		Stmt.close();
		rs.close();
		return rstMap;
	}

	/**
	 * 获取库号PK
	 * 
	 * @param storeStr
	 * @return
	 * @throws BusinessException
	 */
	private Map<String, String> getABStore(String storeStr) throws Exception {
		storeStr = String.format("%04d", Integer.parseInt(storeStr.trim()));
		if(storeStr.equals("")){
			throw new Exception("库号[" + storeStr + "]格式化后不可为空");
		}
		String sql = "select c.pk_pubdoc,c.code,b.pk_pubdoc,b.code,a.pk_pubdoc,a.code from mtws_pubdoc a,mtws_pubdoc b,mtws_pubdoc c where   a.pid=b.pk_pubdoc and b.pid=c.pk_pubdoc and a.code like '_____"+storeStr+"' and a.name like '%库'";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(sql);
		Object[] objs = new Object[rs.getMetaData().getColumnCount()];
		List<Object[]> result = new ArrayList<Object[]>();
		while (rs.next()) {
			for(int col=0; col<rs.getMetaData().getColumnCount();col++){
				objs[col]=rs.getObject(col+1);
			}
			result.add(objs);
		}
		Stmt.close();
		rs.close();
		if (result.size()<1) {
			throw new Exception("查无库号[" + storeStr + "]的片区冻库档案资料");
		}
		for(int i=0; i<6; i++){			
			if(result.get(0)[i]==null || "".equals(result.get(0)[i])){
				throw new Exception("库号[" + storeStr + "]的片区栋库资料["+result.get(0)[i]+"]不可为空");
			}
		}
		Map<String, String> rstMap = new HashMap<String, String>();
		areaNo = result.get(0)[1].toString();
		buildingNo = result.get(0)[3].toString();
		if(buildingNo.length()!=5){
			throw new Exception("栋号编码[" + buildingNo + "]的长度不为5码");
		}
		buildingNo = buildingNo.substring(2, 5);
		storeNo = result.get(0)[5].toString();
		if(storeNo.length()!=9){
			throw new Exception("库号编码[" + storeNo + "]的长度不为9码");
		}
		storeNo = storeNo.substring(5, 9);
		rstMap.put(areaNo, result.get(0)[0].toString());
		rstMap.put(buildingNo, result.get(0)[2].toString());
		rstMap.put(storeNo, result.get(0)[4].toString());
		
		return rstMap;
	}
	
	private void initJarNumByStoreMap() throws SQLException {
		// TODO 自动生成的方法存根
		// 取得05片区各库位总坛数
		if (areaNo.equals(""))
			areaNo = "XXXXX";
		String jarNumStoreSql = "select substr(code,1,9) as store ,count(code) jarnum from mtws_jar where code like '"
				+ areaNo + "%' and dr=0 " + " group by substr(code,1,9) ";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(jarNumStoreSql);
		while (rs.next()) {
			jarNumByStoreMap.put(rs.getString("store"), rs.getString("jarnum"));
		}
		Stmt.close();
		rs.close();
	}

	public int create(String sql) throws SQLException {
		// PreparedStatement:是预编译的,对于批量处理可以大大提高效率.也叫JDBC存储过程
		// Statement:在对数据库只执行一次性存取的时侯，用 Statement对象进行处理。
		Statement Stmt = conn.createStatement();
		// 返回新增或更新数据量
		int i = Stmt.executeUpdate(sql);
		Stmt.close();
		return i;
	}

	public int update(String sql) throws SQLException {
		Statement Stmt = conn.createStatement();
		// 返回新增或更新数据量
		int i = Stmt.executeUpdate(sql);
		Stmt.close();
		return i;
	}

	public int delete(String pk_jar) throws SQLException {
		Statement Stmt = conn.createStatement();
		// 返回新增或更新数据量
		int i = Stmt.executeUpdate("delete from mtws_iquantity where pk_jar='" + pk_jar + "' ");
		Stmt.close();
		return i;
	}

	private String getProduct(String pdType) {
		// TODO 自动生成的方法存根
		String productName = "";
		if (pdType.equals("NW"))
			productName = "新酒——完工待检";
		else if (pdType.equals("PG"))
			productName = "盘勾酒——正常盘勾";
		else if (pdType.equals("GD"))
			productName = "普通勾兑酒——完工待检";
		else if (pdType.equals("TD"))
			productName = "坛底酒——勾兑坛底酒";
		return getStrMapValue(productMap, productName);
	}

	private Map initProductMap() throws SQLException {
		// TODO 自动生成的方法存根
		String sql = "select name,pk_material from bd_material_v where dr=0";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(sql);
		Map rstMap = new HashMap();
		while (rs.next()) {
			rstMap.put(rs.getString("name"), rs.getString("pk_material"));
		}
		Stmt.close();
		rs.close();
		return rstMap;
	}

	private String getPk_jar(String jarNo) throws SQLException {
		String sql = "select pk_jar from mtws_jar where code='" + jarNo + "' and dr=0";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(sql);
		String pk_jar = "";
		while (rs.next()) {
			pk_jar = rs.getString("pk_jar");
		}
		Stmt.close();
		rs.close();
		return pk_jar;
	}

	private List getJarList(String strMapValue) {
		List jarList = new ArrayList();
		// TODO 自动生成的方法存根
		strMapValue = strMapValue.replaceAll("--", "-");
		strMapValue = strMapValue.replaceAll("－－", "-");
		strMapValue = strMapValue.replaceAll("－", "-");
		strMapValue = strMapValue.replaceAll("、", "@");
		strMapValue = strMapValue.replaceAll(",", "@");
		strMapValue = strMapValue.replaceAll("，", "@");
		strMapValue = strMapValue.replaceAll("\\.", "@");
		String[] strAry = strMapValue.split("@");
		for (int i = 0; i < strAry.length; i++) {
			String strTmp = strAry[i];
			if (strTmp.indexOf("-") > -1) {
				String[] strTmpAry = strTmp.split("-");
				int min = Integer.parseInt(strTmpAry[0]);
				int max = Integer.parseInt(strTmpAry[1]);
				if (min > max) {
					int temp = min;
					min = max;
					max = temp;
				}
				for (int j = 0; j <= max - min; j++) {
					String jarNo = String.valueOf(min + j);
					jarList.add(areaNo + buildingNo + storeNo + "0000".substring(0, 4 - jarNo.length()) + jarNo);
				}
			} else {
				jarList.add(areaNo + buildingNo + storeNo + "0000".substring(0, 4 - strTmp.length()) + strTmp);
			}
		}
		// if(strAry==null||strAry.length==0){
		// String[] strAry2 = strMapValue.split("-");
		// int min = Integer.parseInt(strAry2[0]);
		// int max = Integer.parseInt(strAry2[1]);
		// if(min > max){
		// int temp = min;
		// min = max;
		// max = temp;
		// }
		// for(int j=0;j<=max-min;j++){
		// String jarNo = String.valueOf(min + j);
		// jarList.add(areaNo+buildingNo+storeNo+"0000".substring(0,
		// 4-jarNo.length())+jarNo);
		// }
		// }
		Collections.sort(jarList);
		return jarList;
	}

	private String getGrade(String grade) {
		// TODO 自动生成的方法存根
		if (grade.indexOf("特") > -1)
			return "0";
		else if (grade.indexOf("一") > -1)
			return "1";
		else if (grade.indexOf("二") > -1)
			return "2";
		else if (grade.indexOf("三") > -1)
			return "3";
		else if (grade.indexOf("四") > -1)
			return "4";
		else if (grade.indexOf("未") > -1)
			return "9";
		return "";
	}

	private String getBzMessage(String strMapValue) {
		// TODO 自动生成的方法存根
		if (strMapValue.indexOf("次品") > -1 && strMapValue.indexOf("低度") > -1 && strMapValue.indexOf("泥") > -1)
			strMapValue = "低度次品泥臭味";
		else if (strMapValue.indexOf("次品") > -1 && strMapValue.indexOf("低度") < 0 && strMapValue.indexOf("泥") > -1)
			strMapValue = "次品泥臭味";
		else if (strMapValue.indexOf("次品") > -1 && strMapValue.indexOf("低度") < 0 && strMapValue.indexOf("霉") > -1)
			strMapValue = "次品霉味重";
		else if (strMapValue.indexOf("次品") > -1 && strMapValue.indexOf("低度") > -1 && strMapValue.indexOf("泥") < 0
				&& strMapValue.indexOf("霉") < 0)
			strMapValue = "低度次品";
		else if (strMapValue.indexOf("次品") > -1)
			strMapValue = "次品";
		else if (strMapValue.indexOf("霉") > -1 && strMapValue.indexOf("泥") > -1
				&& strMapValue.indexOf("霉") < strMapValue.indexOf("泥"))
			strMapValue = "霉味 泥味";
		else if (strMapValue.indexOf("霉") > -1 && strMapValue.indexOf("泥") > -1
				&& strMapValue.indexOf("霉") > strMapValue.indexOf("泥"))
			strMapValue = "泥味 霉味";
		else if (strMapValue.indexOf("霉") > -1 && strMapValue.indexOf("油") > -1
				&& strMapValue.indexOf("霉") < strMapValue.indexOf("油"))
			strMapValue = "霉味 油味";
		else if (strMapValue.indexOf("霉") > -1 && strMapValue.indexOf("油") > -1
				&& strMapValue.indexOf("霉") > strMapValue.indexOf("油"))
			strMapValue = "油味 霉味";
		else if (strMapValue.indexOf("泥") > -1 && strMapValue.indexOf("糊") > -1)
			strMapValue = "泥味 糊味";
		else if (strMapValue.indexOf("泥") > -1 && strMapValue.indexOf("酸") > -1)
			strMapValue = "泥味 算味";
		else if (strMapValue.indexOf("泥") > -1 && strMapValue.indexOf("酸") > -1)
			strMapValue = "泥味 算味";
		else if (strMapValue.indexOf("微泥") > -1 && strMapValue.indexOf("臭") > -1)
			strMapValue = "微泥臭";
		else if (strMapValue.indexOf("微泥") > -1)
			strMapValue = "微泥";
		else if (strMapValue.indexOf("馊") > -1 && strMapValue.indexOf("盐") > -1)
			strMapValue = "馊味 盐菜味";

		return getStrMapValue(defDocMap, strMapValue);
	}

	private String getIturns(String strMapValue) throws Exception {
		// TODO 自动生成的方法存根
		if (strMapValue.equals("1"))
			strMapValue = "一轮次";
		else if (strMapValue.equals("2"))
			strMapValue = "二轮次";
		else if (strMapValue.equals("3"))
			strMapValue = "三轮次";
		else if (strMapValue.equals("4"))
			strMapValue = "四轮次";
		else if (strMapValue.equals("5"))
			strMapValue = "五轮次";
		else if (strMapValue.equals("6"))
			strMapValue = "六轮次";
		else if (strMapValue.equals("7"))
			strMapValue = "七轮次";
		else if (strMapValue.equals("1.2"))
			strMapValue = "一轮次尾酒";
		else if (strMapValue.equals("2.2"))
			strMapValue = "二轮次尾酒";
		else if (strMapValue.equals("3.2"))
			strMapValue = "三轮次尾酒";
		// else throw new Exception("Excel轮次["+strMapValue+"]未取得对应的轮次编码!");
		return getStrMapValue(defDocMap, strMapValue);
	}

	private String getsweetType(String strMapValue) {
		// TODO 自动生成的方法存根
		if (strMapValue.indexOf("一") > -1 && strMapValue.indexOf("甜") > -1 && strMapValue.indexOf("标准") < 0)
			strMapValue = "一等醇甜";
		else if (strMapValue.indexOf("一") > -1 && strMapValue.indexOf("甜") > -1 && strMapValue.indexOf("标准") > -1)
			strMapValue = "一等醇甜(标准)";
		if (strMapValue.indexOf("二") > -1 && strMapValue.indexOf("甜") > -1 && strMapValue.indexOf("标准") < 0)
			strMapValue = "二等醇甜";
		else if (strMapValue.indexOf("二") > -1 && strMapValue.indexOf("甜") > -1 && strMapValue.indexOf("标准") > -1)
			strMapValue = "二等醇甜(标准)";
		if (strMapValue.indexOf("一") > -1 && strMapValue.indexOf("酱") > -1 && strMapValue.indexOf("标准") < 0)
			strMapValue = "一等酱香";
		else if (strMapValue.indexOf("一") > -1 && strMapValue.indexOf("酱") > -1 && strMapValue.indexOf("标准") > -1)
			strMapValue = "一等酱香(标准)";
		if (strMapValue.indexOf("二") > -1 && strMapValue.indexOf("酱") > -1 && strMapValue.indexOf("标准") < 0)
			strMapValue = "二等酱香";
		else if (strMapValue.indexOf("二") > -1 && strMapValue.indexOf("酱") > -1 && strMapValue.indexOf("标准") > -1)
			strMapValue = "二等酱香(标准)";
		else if (strMapValue.indexOf("一") > -1 && strMapValue.indexOf("窖") > -1)
			strMapValue = "一等窖甜";
		else if (strMapValue.indexOf("二") > -1 && strMapValue.indexOf("窖") > -1)
			strMapValue = "二等窖甜";
		else if (strMapValue.indexOf("三") > -1 && strMapValue.indexOf("窖") > -1)
			strMapValue = "三等窖甜";
		else if (strMapValue.indexOf("混合") > -1)
			strMapValue = "混合香";

		return getStrMapValue(defDocMap, strMapValue);
	}

	private String getPdType(String strMapValue) throws Exception {
		// TODO 自动生成的方法存根
		if (strMapValue.indexOf("新酒") > -1) {
			status = "10";
			return "NW";
		} else if (strMapValue.indexOf("盘勾") > -1) {
			status = "20";
			return "PG";
		} else if (strMapValue.indexOf("勾兑酒") > -1) {
			status = "30";
			return "GD";
		} else if (strMapValue.indexOf("坛底") > -1 || strMapValue.equals("")) {
			status = "01";
			return "TD";
		} else
			throw new Exception("类型[" + strMapValue + "]未设定!");

	}

	private String getInDate(String strMapValue, String strMapValue2, String strMapValue3) throws Exception {
		// TODO 自动生成的方法存根
		if (strMapValue.equals("") && strMapValue2.equals("") && strMapValue3.equals(""))
			return "";
		else if (!strMapValue.equals("") && !strMapValue2.equals("") && !strMapValue3.equals(""))
			return strMapValue + "-" + "00".substring(0, 2 - strMapValue2.length()) + strMapValue2 + "-"
					+ "00".substring(0, 2 - strMapValue3.length()) + strMapValue3;
		// else throw new
		// Exception("日期["+strMapValue+"/"+strMapValue2+"/"+strMapValue3+"]格式错误!");
		else
			return strMapValue + strMapValue2 + strMapValue3;
	}

	private String getIyear(String strMapValue) {
		// TODO 自动生成的方法存根
		return getStrMapValue(defDocMap, strMapValue);
	}

	private String getStrMapValue(Map map, String key) {
		return map != null && map.get(key) != null ? map.get(key).toString().trim() : "";
	}

	private BigDecimal getDecMapValue(Map map, String key) {
		return new BigDecimal(map != null && map.get(key) != null && !map.get(key).toString().trim().equals("")
				? map.get(key).toString().trim() : "0");
	}

	private Map initPubDocMap() throws SQLException {
		String sql = "select code,pk_pubdoc from mtws_pubdoc where dr=0";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(sql);
		Map rstMap = new HashMap();
		while (rs.next()) {
			rstMap.put(rs.getString("code"), rs.getString("pk_pubdoc"));
		}
		Stmt.close();
		rs.close();
		return rstMap;
	}

	private Map initDefDocMap() throws SQLException {
		String sql = "select name,pk_defdoc from bd_defdoc where pk_defdoclist in ('1001A91000000000292P','1001ZZ10000000007FSW','1001ZZ100000000093XT','1001ZZ10000000007KK1') and dr=0";
		Statement Stmt = conn.createStatement();
		ResultSet rs = Stmt.executeQuery(sql);
		Map rstMap = new HashMap();
		while (rs.next()) {
			rstMap.put(rs.getString("name"), rs.getString("pk_defdoc"));
		}
		Stmt.close();
		rs.close();
		return rstMap;
	}

	private String getPk_area(String area) throws Exception {
		String pk_area = "";
		if (area.indexOf("一") > -1) {
			areaNo = "01";
			pk_area = getStrMapValue(pubdocMap, "01");
		} else if (area.indexOf("二") > -1) {
			areaNo = "02";
			pk_area = getStrMapValue(pubdocMap, "02");
		} else if (area.indexOf("三") > -1) {
			areaNo = "03";
			pk_area = getStrMapValue(pubdocMap, "03");
		} else if (area.indexOf("四") > -1) {
			areaNo = "04";
			pk_area = getStrMapValue(pubdocMap, "04");
		} else if (area.indexOf("五") > -1) {
			areaNo = "05";
			pk_area = getStrMapValue(pubdocMap, "05");
		} else if (area.indexOf("六") > -1) {
			areaNo = "06";
			pk_area = getStrMapValue(pubdocMap, "06");
		} else if (area.indexOf("七") > -1) {
			areaNo = "07";
			pk_area = getStrMapValue(pubdocMap, "07");
		}
		if (areaNo.equals(""))
			throw new Exception("Excel的片区[" + area + "]未找到对应的片区!");
		if (pk_area.equals(""))
			throw new Exception("Excel的片区[" + area + "]未找到对应的片区编码!");
		return pk_area;
	}

	private String getPk_building(String building) throws Exception {
		if (building.equals(""))
			throw new Exception("Excel的栋号[" + building + "]未找到对应的栋号!");
		buildingNo = "000".substring(0, 3 - building.length()) + building;
		String pk_building = getStrMapValue(pubdocMap, areaNo + buildingNo);
		if (pk_building.equals(""))
			throw new Exception("Excel的栋号[" + building + "]未找到对应的栋号编码!");
		return pk_building;
	}

	private String getPk_store(String store) throws Exception {
		if (store.equals(""))
			throw new Exception("Excel的库号[" + store + "]未找到对应的库号!");
		storeNo = "0000".substring(0, 4 - store.length()) + store;
		String storeCode = getStrMapValue(storeMap, storeNo);
		if (storeCode.equals("") || storeCode.length() < 9)
			throw new Exception("Excel的库号[" + store + "]未找到正确的库号code!");
		buildingNo = storeCode.substring(2, 5);
		String pk_store = getStrMapValue(pubdocMap, storeCode);
		if (pk_store.equals(""))
			throw new Exception("Excel的库号[" + store + "]未找到对应的库号编码!");
		return pk_store;
	}

	private String getTubCode(String tubcode, String pdType, String year, String oriDate) throws Exception {
		String yearSub = "";
		if (!year.equals(""))
			yearSub = year.substring(2, 4);
		else if (!oriDate.equals(""))
			yearSub = oriDate.substring(2, 4);
		else
			yearSub = (Calendar.getInstance().get(Calendar.YEAR) + "").substring(2, 4);

		if (pdType.indexOf("新酒") > -1) {
			if (!tubcode.equals(""))
				throw new Exception("Excel该行资料的类型为新酒，桶号[" + tubcode + "]应该为空!");
			return "";
		} else if (!pdType.equals("")) {
			if (tubcode.equals(""))
				throw new Exception("Excel该行资料的类型非新酒，桶号[" + tubcode + "]不可为空!");
			if (pdType.indexOf("盘勾") > -1) {
				return "PG" + areaNo + buildingNo + storeNo + yearSub + "000".substring(0, 3 - tubcode.length())
						+ tubcode;
			} else if (pdType.indexOf("勾兑酒") > -1) {
				return "GD" + areaNo + buildingNo + storeNo + yearSub + "000".substring(0, 3 - tubcode.length())
						+ tubcode;
			} else if (pdType.indexOf("坛底") > -1) {
				return "TD" + areaNo + buildingNo + storeNo + yearSub + "000".substring(0, 3 - tubcode.length())
						+ tubcode;
			} else
				throw new Exception("Excel该行资料类型[" + pdType + "]不存在!");
		} else
			throw new Exception("类型不可为空!");

	}

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

	public void close(Connection conn) throws SQLException {
		if (conn != null) {
			conn.close();
		}
	}

}
