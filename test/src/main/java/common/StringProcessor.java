package common;

import org.apache.commons.lang.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.*;

public class StringProcessor {

    //    final static String spe ="`!@#$%^&*()_+':\"-=\\[\\]{};\\|,.<>/?~";
    final static String any_special_symbol = "`|!|@|#|$|%|^|&|*|(|)|_|+|'|:|\"|\\-|=|\\[|\\]|{|}|;|\\||,|.|<|>|/|?|~";
    final static String custom_special_symbol = "`|!|@|#|%|^|&|*|(|)|_|+|'|\"|\\-|=|\\[|\\]|{|}|;|\\||,|.|<|>|/|?|~";
    final static String contain_special_symbol = "[.*" + any_special_symbol + ".*]+";

    //IF(ISBLANK(C7),"",(C8-C9)*100/(C8-C7)) ---> ISBLANK(C7),"",(C8-C9)*100/(C8-C7)
    public static  String removeProrgrammingMethodFromMostOuter(String formula){
        //find the first occurence of (
        if(StringUtils.isNotBlank(formula)){
            int leftParentheses = formula.indexOf("(");
            int rightParentheses = formula.lastIndexOf(")");
            return formula.substring(leftParentheses +1 ,rightParentheses);
        }
        else return "";

    }

    public static boolean ifContain(String toBeCompared, String CellValue) {
        if (StringUtils.isNotBlank(CellValue) && CellValue.toLowerCase().contains(toBeCompared)) return true;
        else return false;
    }

    public static boolean ifEqual(String toBeCompared, String CellValue) {
        if (StringUtils.isNotBlank(CellValue) && replaceSpecialSymbol(CellValue.toLowerCase(), "").equals(toBeCompared))
            return true;
        else return false;
    }
    public static boolean ifContainSpecialSymbol(String strToCheck){
        return Pattern.matches(contain_special_symbol, strToCheck);
    }

    public static String replaceSpecialSymbol(String strToCheck, String replacement){
        return  strToCheck.replaceAll("["+any_special_symbol+"]+", replacement);
    }
    public static String replaceCustomSymbol(String strToCheck, String replacement){
        return  strToCheck.replaceAll("["+custom_special_symbol+"]+", replacement);
    }

    /*"ifBlank(C3)"-->true
      "C3*17+O2"-->false */
    public static boolean ifContainMethod(String StrToChecl){
        return Pattern.matches("[a-zA-Z]{2,}.*",replaceSpecialSymbol(StrToChecl,""));
    }

    public static String CapitalCharToLowerCaseWithDelimitor(String strToCheck, String Delimitor){

        List<Integer> indexOfCapital =IndexOfMatches(strToCheck,"[A-Z]");
        String result =strToCheck;
        for(Integer i: indexOfCapital){
            String capitalStr = String.valueOf(strToCheck.charAt(i));
            result = result.replace(capitalStr, Delimitor+capitalStr.toLowerCase());
        }
        return  result.trim();
    }

public static List<Integer> IndexOfMatches(String StrToCheck, String regularExp){
    List<Integer> result = new ArrayList<>();
    Pattern pattern = Pattern.compile(regularExp);
    Matcher matcher = pattern.matcher(StrToCheck);
    while (matcher.find()){
        result.add(matcher.start());//this will give you index
    }
    return result;
}

    public static void main(String... arg) {
//        System.out.println(replaceSpecialSymbol("C15"," "));
//        System.out.println(Pattern.matches(".*[A-Z]+.*","0.154"));
        System.out.println(Pattern.matches("[A-Z][0-9]+[:][A-Z][0-9]+", "A10$:A105"));
        System.out.println(replaceCustomSymbol("^RSQ(D15:D19,B15:B19)"," "));
//        System.out.println(CapitalCharToLowerCaseWithDelimitor("ActKind"," "));
//        System.out.println(ifContainMethod("C3*17+O2"));//false
//        System.out.println(ifContainMethod("ifBlank(C3)"));//true
    }
}
