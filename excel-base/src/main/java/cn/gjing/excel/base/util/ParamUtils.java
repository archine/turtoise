package cn.gjing.excel.base.util;

import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**
 * Param utils
 *
 * @author Gjing
 **/
public final class ParamUtils {
    /**
     * Whether the array contains a value
     *
     * @param arr array
     * @param val value
     * @return boolean
     */
    public static boolean contains(String[] arr, String val) {
        if (arr == null || arr.length == 0) {
            return false;
        }
        for (String o : arr) {
            if (o.equals(val)) {
                return true;
            }
        }
        return false;
    }

    /**
     * MD5 encryption
     *
     * @param body need to encryption
     * @return encrypted string
     */
    public static String encodeMd5(String body) {
        StringBuilder buf = new StringBuilder();
        try {
            MessageDigest md = MessageDigest.getInstance("MD5");
            md.update(body.getBytes());
            byte[] b = md.digest();
            int i;
            for (byte b1 : b) {
                i = b1;
                if (i < 0) {
                    i += 256;
                }
                if (i < 16) {
                    buf.append("0");
                }
                buf.append(Integer.toHexString(i));
            }
        } catch (NoSuchAlgorithmException e) {
            e.printStackTrace();
            return "";
        }
        return buf.toString();
    }

    /**
     * Whether it's equal or not
     *
     * @param param1     param1
     * @param param2     param2
     * @return boolean
     */
    public static boolean equals(Object param1, Object param2) {
        return param1 == param2 || param1.equals(param2);
    }

    /**
     * Number to English letter
     *
     * @param number number
     * @return letter
     */
    public static String numberToEn(int number) {
        char prefix = 'A';
        if (number < 26) {
            return "" + (char) ('A' + number);
        }
        char suffix;
        if ((number - 25) % 26 == 0) {
            suffix = (char) (prefix + 25);
            prefix = (char) (prefix + (number - 25) / 26 - 1);
        } else {
            suffix = (char) ('A' + (number - 25) % 26 - 1);
            prefix = (char) (prefix + (number - 25) / 26);
        }
        return "" + prefix + suffix;
    }
}
