package com.yelisoft;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

public class Utils {
    static private Map<String, Integer> monthsMap = new HashMap<String, Integer>(){{
        put("январь", 1);
        put("февраль", 2);
        put("март", 3);
        put("апрель", 4);
        put("май", 5);
        put("июнь", 6);
        put("июль", 7);
        put("август", 8);
        put("сентябрь", 9);
        put("октябрь", 10);
        put("ноябрь", 11);
        put("декабрь", 12);
    }};

    public static Integer getMonthNumber(String monthName) {
        return monthsMap.get(monthName.toLowerCase(Locale.ROOT));
    }
}
