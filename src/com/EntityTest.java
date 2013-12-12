package com;

import java.lang.reflect.Field;

public class EntityTest {
	public static void main(String[] arvh) throws Exception, Exception{
		Student stu = new Student();
		Field[] f =stu.getClass().getDeclaredFields();
		
		 for(int i=0;i<f.length;i++)
		  {
			 f[i].set(stu, "sf"+i);
			 System.out.println((i+1)+":"+f[i].getName()+":"+f[i].get(stu));
		   
		  }
	}
}
