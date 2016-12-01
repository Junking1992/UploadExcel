package com.jun.service;

public class ProgressUtil {
	
	//总进度
	public int allCount;
	//当进度
	public int currentCount;
	
	public long startTime;
	
	public String errorRowNum = "";
	
	public ProgressUtil() {
		this.allCount = 0;
		this.currentCount = 0;
		this.startTime = System.currentTimeMillis();
	}
	
	public String getProgress(){
		return currentCount + "/" + allCount;
	}

	public synchronized void addCount() {
		this.currentCount++;
	}
	
	public String getTime(){
		return (System.currentTimeMillis()-startTime)/1000 + "秒";
	}

	public int getAllCount() {
		return allCount;
	}

	public void setAllCount(int allCount) {
		this.allCount = allCount;
	}

	public int getCurrentCount() {
		return currentCount;
	}

	public void setCurrentCount(int currentCount) {
		this.currentCount = currentCount;
	}

}
