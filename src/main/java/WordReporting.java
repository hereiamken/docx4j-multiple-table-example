import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.*;

public class WordReporting {

    /**
     * Main process for Process generate docx file using template
     * @param fileInputNm
     * @param fileOutputNm
     * @param mapKeyValues
     * @param valuesTables
     * @throws IOException
     * @throws Docx4JException
     * @throws JAXBException
     */
    public void wordProcess(String fileInputNm, String fileOutputNm, Map<String, String> mapKeyValues, Map<List<String>, List<Map<String, String>>> valuesTables) throws IOException, Docx4JException, JAXBException {
        WordprocessingMLPackage template = getTemplate(fileInputNm);
        MainDocumentPart documentPart = template.getMainDocumentPart();
        HashMap<String, String> mappings = new HashMap<>();

        // replace all texts with values
        for (Map.Entry<String, String> entry : mapKeyValues.entrySet()) {
            System.out.println("Key : " + entry.getKey() + " Value : " + entry.getValue());
            replacePlaceholder(template, entry.getValue(), entry.getKey());
            mappings.put(entry.getKey(), StringUtils.EMPTY);
        }

        // map mapKeyValues with key is list placeholders and values is list content to placements
        for (Map.Entry<List<String>, List<Map<String, String>>> entry : valuesTables.entrySet()) {
            String[] placeholders = entry.getKey().toArray(new String[0]);
            List<Map<String, String>> valueOfTables = entry.getValue();
            replaceTable(placeholders, valueOfTables, template);
            entry.getKey().stream().forEach(s -> {
                mappings.put(s, StringUtils.EMPTY);
            });
        }

        // create break page
        documentPart.getContent().add(createPageBreak());

        // remove all placeholders which are empty.
        documentPart.variableReplace(mappings);

        // write to docx file.
        writeDocxToStream(template, fileOutputNm);
    }

    /**
     * @param name
     * @return
     * @throws Docx4JException
     */
    private WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
        File f = new File(WordReporting.class.getResource(name).getFile());
        InputStream targetStream = new FileInputStream(f);
        WordprocessingMLPackage template = WordprocessingMLPackage.load(targetStream);
        return template;
    }

    /**
     * @param obj
     * @param toSearch
     * @return
     */
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch)) result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }

    /**
     * @param template
     * @param name
     * @param placeholder
     */
    private void replacePlaceholder(WordprocessingMLPackage template, String name, String placeholder) {
        List<Object> texts = getAllElementFromObject(template.getMainDocumentPart(), Text.class);

        for (Object text : texts) {
            Text textElement = (Text) text;
            if (textElement.getValue().equals(placeholder)) {
                textElement.setValue(name);
            }
        }
    }

    /**
     * @param placeholders
     * @param textToAdd
     * @param template
     * @throws Docx4JException
     * @throws JAXBException
     */
    private void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd, WordprocessingMLPackage template) throws Docx4JException, JAXBException {
        List<Object> tables = getAllElementFromObject(template.getMainDocumentPart(), Tbl.class);

        // 1. find the table
        Tbl tempTable = getTemplateTable(tables, placeholders[0]);
        List<Object> rows = getAllElementFromObject(tempTable, Tr.class);

        // first row is header, second row is content
        if (rows.size() == 2) {
            // this is our template row
            Tr templateRow = (Tr) rows.get(1);

            for (Map<String, String> replacements : textToAdd) {
                // 2 and 3 are done in this method
                addRowToTable(tempTable, templateRow, replacements);
            }

            // 4. remove the template row
            tempTable.getContent().remove(templateRow);
        }
    }

    /**
     * @param tables
     * @param templateKey
     * @return
     */
    private Tbl getTemplateTable(List<Object> tables, String templateKey) {
        for (Iterator<Object> iterator = tables.iterator(); iterator.hasNext(); ) {
            Object tbl = iterator.next();
            List<?> textElements = getAllElementFromObject(tbl, Text.class);
            for (Object text : textElements) {
                Text textElement = (Text) text;
                if (textElement.getValue() != null && (textElement.getValue().equals(templateKey)) || textElement.getValue().equals("${" + templateKey + "}")) {
                    return (Tbl) tbl;
                }
            }
        }
        return null;
    }

    /**
     * @param reviewtable
     * @param templateRow
     * @param replacements
     */
    private static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements) {
        Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);
        List<?> textElements = getAllElementFromObject(workingRow, Text.class);
        for (Object object : textElements) {
            Text text = (Text) object;
            String replacementValue = (String) getObjectByKey(replacements, text.getValue());

            if (replacementValue != null) {
                System.out.println("Before " + text.getValue());
                text.setValue(addPlaceHolder(text.getValue()));
                System.out.println("After " + text.getValue());
                text.setValue(replacementValue);
            }
        }
        reviewtable.getContent().add(workingRow);
    }

    /**
     * @param template
     * @param target
     * @throws Docx4JException
     */
    private void writeDocxToStream(WordprocessingMLPackage template, String target) throws Docx4JException {
        File f = new File(target);
        template.save(f);
    }

    public static void main(String[] args) throws Exception {
        String fileName = "WordReporting.docx";
        String outPutName = "out_WordReporting.docx";
        Map<String, String> mapKeyValues = new HashMap<>();
        mapKeyValues.put("tblNameA", "Danh sách nhân viên phòng A");
        mapKeyValues.put("tblNameB", "Danh sách vợ nhân viên");

        Map<List<String>, List<Map<String, String>>> valuesTables = new HashMap<>();
        List<String> placeholder1 = Arrays.asList("no", "name", "age", "position", "department");
        List<String> placeholder2 = Arrays.asList("noWife", "nameWife", "marriageStatus", "country", "ageWife");

        Map<String, String> tblA = new HashMap<>();
        tblA.put("no", "1");
        tblA.put("name", "Jan");
        tblA.put("age", "25");
        tblA.put("position", "DEV3");
        tblA.put("department", "FPT1");

        Map<String, String> tblB = new HashMap<>();
        tblB.put("no", "2");
        tblB.put("name", "Feb");
        tblB.put("age", "25");
        tblB.put("position", "DEV3");
        tblB.put("department", "FPT1");

        Map<String, String> tblC = new HashMap<>();
        tblC.put("no", "3");
        tblC.put("name", "March");
        tblC.put("age", "25");
        tblC.put("position", "DEV3");
        tblC.put("department", "FPT1");

        Map<String, String> wifeA = new HashMap<>();
        wifeA.put("noWife", "1");
        wifeA.put("nameWife", "March 2");
        wifeA.put("marriageStatus", "Yes");
        wifeA.put("country", "Hanoi");
        wifeA.put("ageWife", "22");

        Map<String, String> wifeB = new HashMap<>();
        wifeB.put("noWife", "2");
        wifeB.put("nameWife", "Jun 2");
        wifeB.put("marriageStatus", "No");
        wifeB.put("country", "Hanoi");
        wifeB.put("ageWife", "23");

        valuesTables.put(placeholder1, Arrays.asList(tblA, tblB, tblC));
        valuesTables.put(placeholder2, Arrays.asList(wifeA, wifeB));

        new WordReporting().wordProcess(fileName, outPutName, mapKeyValues, valuesTables);
    }

    private static String addPlaceHolder(String str) {
        if (Objects.isNull(str)) {
            return null;
        }
        if (str.contains("${") && str.contains("}")) {
            return str;
        } else {
            return "${" + str + "}";
        }
    }

    private static String removePlaceHolder(String str) {
        str = str.replaceAll("\\s", "");
        str = str.replace("$", "");
        str = str.replace("{", "");
        str = str.replace("}", "");
        return str;
    }

    private static String getObjectByKey(Map<String, String> replacements, String key) {
        if (Objects.nonNull(replacements.get(key))) {
            return replacements.get(key);
        } else if (Objects.nonNull(replacements.get(addPlaceHolder(key)))) {
            return replacements.get(addPlaceHolder(key));
        } else if (Objects.nonNull(replacements.get(removePlaceHolder(key)))) {
            return replacements.get(removePlaceHolder(key));
        }
        return null;
    }

    private P createPageBreak() {
        Br br = Context.getWmlObjectFactory().createBr();
        br.setType(STBrType.PAGE);

        R run = Context.getWmlObjectFactory().createR();
        run.getContent().add(br);

        P paragraph = Context.getWmlObjectFactory().createP();
        paragraph.getContent().add(run);
        return paragraph;
    }
}
