package com.ayg.poiutils;

import java.lang.reflect.Type;
import java.lang.reflect.TypeVariable;
import java.util.ArrayList;
import java.util.List;

import com.whl.poi.demo.UserInfo;

import sun.reflect.generics.reflectiveObjects.ParameterizedTypeImpl;

public class ListTTest {
	public static void main(String[] args) {
		List<UserInfo> list = new ArrayList<UserInfo>();
		Type type = list.getClass().getGenericSuperclass();
		Class<?> clazz = (Class<?>) ((ParameterizedTypeImpl)list.getClass().getGenericSuperclass()).getActualTypeArguments()[0];
		System.out.println(clazz);
	}

}
