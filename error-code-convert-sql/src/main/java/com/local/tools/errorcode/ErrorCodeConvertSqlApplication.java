package com.local.tools.errorcode;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ErrorCodeConvertSqlApplication {
    private static final String DEFAULT_OUTPUT = "errorcode.sql";
    private static final String DEFAULT_PREFIX = "RWI.";
    private static final String DEFAULT_TIMESTAMP = "2021-11-26 15:30:09.301";
    private static final String DEFAULT_CREATE_BY = "kuka";
    private static final String DEFAULT_CREATE_APP = "OptionalCollection:312";
    private static final String DEFAULT_GROUP = "robot";

    public static void main(String[] args) {
        try {
            Config config = Config.parse(args);
            if (config.help) {
                printUsage();
                return;
            }

            List<ErrorCodeRow> rows = readRows(config);
            if (rows.isEmpty()) {
                throw new IllegalArgumentException("Excel 中没有找到可生成 SQL 的数据行。");
            }

            String sql = SqlWriter.write(rows, config);
            Files.write(config.outputPath, sql.getBytes(StandardCharsets.UTF_8));
            System.out.println("生成成功: " + config.outputPath.toAbsolutePath());
            System.out.println("数据行数: " + rows.size());
        } catch (Exception ex) {
            System.err.println("生成失败: " + ex.getMessage());
            System.err.println();
            printUsage();
            System.exit(1);
        }
    }

    private static List<ErrorCodeRow> readRows(Config config) throws Exception {
        File excel = config.excelPath.toFile();
        if (!excel.isFile()) {
            throw new IllegalArgumentException("Excel 文件不存在: " + config.excelPath);
        }
        if (!excel.getName().toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
            throw new IllegalArgumentException("当前无依赖版本支持 .xlsx 文件，请先另存为 xlsx。");
        }

        XlsxSheet sheet = XlsxReader.read(excel, config.sheetName);
        Header header = findHeader(sheet.rows);
        List<ErrorCodeRow> result = new ArrayList<ErrorCodeRow>();
        for (int i = header.rowIndex + 1; i < sheet.rows.size(); i++) {
            Map<Integer, String> row = sheet.rows.get(i);
            String keySuffix = cellText(row, header.keySuffixColumn);
            String code = cellText(row, header.codeColumn);
            String cn = cellText(row, header.cnColumn);
            String en = cellText(row, header.enColumn);
            if (isBlank(keySuffix) && isBlank(code) && isBlank(cn) && isBlank(en)) {
                continue;
            }
            if (isBlank(keySuffix) || isBlank(code)) {
                continue;
            }

            result.add(new ErrorCodeRow(config.prefix + keySuffix.trim(), code.trim(), cn.trim(), en.trim()));
        }
        return result;
    }

    private static Header findHeader(List<Map<Integer, String>> rows) {
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            Map<Integer, String> row = rows.get(rowIndex);
            Map<String, Integer> columns = new HashMap<String, Integer>();
            for (Map.Entry<Integer, String> entry : row.entrySet()) {
                String value = normalizeHeader(entry.getValue());
                if (!isBlank(value)) {
                    columns.put(value, entry.getKey());
                }
            }

            Integer keySuffixColumn = firstPresent(columns, "传参值", "param", "parameter", "参数值", "i18n_key");
            Integer codeColumn = firstPresent(columns, "key", "code", "error_code", "错误码");
            Integer cnColumn = firstPresent(columns, "中文", "cn", "zh", "zh_cn", "chinese");
            Integer enColumn = firstPresent(columns, "英文", "en", "english");
            if (keySuffixColumn != null && codeColumn != null && cnColumn != null && enColumn != null) {
                return new Header(rowIndex, keySuffixColumn, codeColumn, cnColumn, enColumn);
            }
        }
        throw new IllegalArgumentException("未找到表头，需要包含: 传参值 / key / 中文 / 英文。");
    }

    private static Integer firstPresent(Map<String, Integer> columns, String... names) {
        for (String name : names) {
            Integer index = columns.get(normalizeHeader(name));
            if (index != null) {
                return index;
            }
        }
        return null;
    }

    private static String normalizeHeader(String value) {
        return value == null ? "" : value.trim().toLowerCase(Locale.ROOT).replace("-", "_").replace(" ", "_");
    }

    private static String cellText(Map<Integer, String> row, int columnIndex) {
        String value = row.get(columnIndex);
        return value == null ? "" : value.trim();
    }

    private static boolean isBlank(String value) {
        return value == null || value.trim().isEmpty();
    }

    private static void printUsage() {
        System.out.println("用法:");
        System.out.println("  java -jar target/error-code-convert-sql-1.0.0.jar <xlsx文件> [输出SQL文件]");
        System.out.println();
        System.out.println("可选参数:");
        System.out.println("  -o, --output <file>      输出 SQL 文件，默认 errorcode.sql");
        System.out.println("  --sheet <name>           指定 Sheet，默认第一个 Sheet");
        System.out.println("  --prefix <prefix>        i18n_key 前缀，默认 RWI.");
        System.out.println("  --timestamp <time>       create/update 时间");
        System.out.println("  --create-by <name>       create_by/last_update_by，默认 kuka");
        System.out.println("  --create-app <name>      create_app/last_update_app，默认 OptionalCollection:312");
        System.out.println("  --group <name>           i18n_group，默认 robot");
    }

    private static class Config {
        private boolean help;
        private Path excelPath;
        private Path outputPath = Paths.get(DEFAULT_OUTPUT);
        private String sheetName;
        private String prefix = DEFAULT_PREFIX;
        private String timestamp = DEFAULT_TIMESTAMP;
        private String createBy = DEFAULT_CREATE_BY;
        private String createApp = DEFAULT_CREATE_APP;
        private String group = DEFAULT_GROUP;

        private static Config parse(String[] args) {
            Config config = new Config();
            List<String> positional = new ArrayList<String>();
            for (int i = 0; i < args.length; i++) {
                String arg = args[i];
                if ("-h".equals(arg) || "--help".equals(arg)) {
                    config.help = true;
                    return config;
                } else if ("-o".equals(arg) || "--output".equals(arg)) {
                    config.outputPath = Paths.get(requireValue(args, ++i, arg));
                } else if ("--sheet".equals(arg)) {
                    config.sheetName = requireValue(args, ++i, arg);
                } else if ("--prefix".equals(arg)) {
                    config.prefix = requireValue(args, ++i, arg);
                } else if ("--timestamp".equals(arg)) {
                    config.timestamp = requireValue(args, ++i, arg);
                } else if ("--create-by".equals(arg)) {
                    config.createBy = requireValue(args, ++i, arg);
                } else if ("--create-app".equals(arg)) {
                    config.createApp = requireValue(args, ++i, arg);
                } else if ("--group".equals(arg)) {
                    config.group = requireValue(args, ++i, arg);
                } else {
                    positional.add(arg);
                }
            }

            if (positional.isEmpty()) {
                throw new IllegalArgumentException("请传入 Excel 文件路径。");
            }
            if (positional.size() > 2) {
                throw new IllegalArgumentException("位置参数过多。");
            }
            config.excelPath = Paths.get(positional.get(0));
            if (positional.size() == 2) {
                config.outputPath = Paths.get(positional.get(1));
            }
            return config;
        }

        private static String requireValue(String[] args, int index, String option) {
            if (index >= args.length) {
                throw new IllegalArgumentException(option + " 缺少参数值。");
            }
            return args[index];
        }
    }

    private static class Header {
        private final int rowIndex;
        private final int keySuffixColumn;
        private final int codeColumn;
        private final int cnColumn;
        private final int enColumn;

        private Header(int rowIndex, int keySuffixColumn, int codeColumn, int cnColumn, int enColumn) {
            this.rowIndex = rowIndex;
            this.keySuffixColumn = keySuffixColumn;
            this.codeColumn = codeColumn;
            this.cnColumn = cnColumn;
            this.enColumn = enColumn;
        }
    }

    private static class ErrorCodeRow {
        private final String i18nKey;
        private final String code;
        private final String cn;
        private final String en;

        private ErrorCodeRow(String i18nKey, String code, String cn, String en) {
            this.i18nKey = i18nKey;
            this.code = code;
            this.cn = cn;
            this.en = en;
        }
    }

    private static class XlsxSheet {
        private final List<Map<Integer, String>> rows;

        private XlsxSheet(List<Map<Integer, String>> rows) {
            this.rows = rows;
        }
    }

    private static class XlsxReader {
        private static XlsxSheet read(File file, String sheetName) throws Exception {
            try (ZipFile zip = new ZipFile(file)) {
                List<String> sharedStrings = readSharedStrings(zip);
                String sheetPath = resolveSheetPath(zip, sheetName);
                Document sheetDocument = readXml(zip, sheetPath);
                NodeList rowNodes = sheetDocument.getElementsByTagName("row");
                List<Map<Integer, String>> rows = new ArrayList<Map<Integer, String>>();
                int expectedRow = 1;
                for (int i = 0; i < rowNodes.getLength(); i++) {
                    Element rowElement = (Element) rowNodes.item(i);
                    int rowNumber = parseInt(rowElement.getAttribute("r"), expectedRow);
                    while (expectedRow < rowNumber) {
                        rows.add(new LinkedHashMap<Integer, String>());
                        expectedRow++;
                    }
                    rows.add(readRow(rowElement, sharedStrings));
                    expectedRow = rowNumber + 1;
                }
                return new XlsxSheet(rows);
            }
        }

        private static Map<Integer, String> readRow(Element rowElement, List<String> sharedStrings) {
            Map<Integer, String> row = new LinkedHashMap<Integer, String>();
            NodeList cellNodes = rowElement.getElementsByTagName("c");
            int nextColumn = 0;
            for (int i = 0; i < cellNodes.getLength(); i++) {
                Element cell = (Element) cellNodes.item(i);
                String cellRef = cell.getAttribute("r");
                int column = isBlank(cellRef) ? nextColumn : columnIndex(cellRef);
                row.put(column, readCell(cell, sharedStrings));
                nextColumn = column + 1;
            }
            return row;
        }

        private static String readCell(Element cell, List<String> sharedStrings) {
            String type = cell.getAttribute("t");
            if ("inlineStr".equals(type)) {
                return textOfDescendant(cell, "t");
            }
            String value = directChildText(cell, "v");
            if ("s".equals(type)) {
                int index = parseInt(value, -1);
                return index >= 0 && index < sharedStrings.size() ? sharedStrings.get(index) : "";
            }
            return normalizeNumber(value);
        }

        private static List<String> readSharedStrings(ZipFile zip) throws Exception {
            ZipEntry entry = zip.getEntry("xl/sharedStrings.xml");
            List<String> values = new ArrayList<String>();
            if (entry == null) {
                return values;
            }
            Document document = readXml(zip, entry.getName());
            NodeList siNodes = document.getElementsByTagName("si");
            for (int i = 0; i < siNodes.getLength(); i++) {
                values.add(textOfDescendant((Element) siNodes.item(i), "t"));
            }
            return values;
        }

        private static String resolveSheetPath(ZipFile zip, String sheetName) throws Exception {
            Document workbook = readXml(zip, "xl/workbook.xml");
            Document relationships = readXml(zip, "xl/_rels/workbook.xml.rels");
            Map<String, String> targets = new HashMap<String, String>();
            NodeList relationshipNodes = relationships.getElementsByTagName("Relationship");
            for (int i = 0; i < relationshipNodes.getLength(); i++) {
                Element relationship = (Element) relationshipNodes.item(i);
                targets.put(relationship.getAttribute("Id"), relationship.getAttribute("Target"));
            }

            NodeList sheetNodes = workbook.getElementsByTagName("sheet");
            if (sheetNodes.getLength() == 0) {
                throw new IllegalArgumentException("Excel 没有 Sheet。");
            }
            Element selected = null;
            for (int i = 0; i < sheetNodes.getLength(); i++) {
                Element sheet = (Element) sheetNodes.item(i);
                if (isBlank(sheetName) || sheetName.equals(sheet.getAttribute("name"))) {
                    selected = sheet;
                    break;
                }
            }
            if (selected == null) {
                throw new IllegalArgumentException("未找到 Sheet: " + sheetName);
            }

            String relationshipId = selected.getAttribute("r:id");
            if (isBlank(relationshipId)) {
                relationshipId = selected.getAttribute("id");
            }
            String target = targets.get(relationshipId);
            if (isBlank(target)) {
                throw new IllegalArgumentException("无法定位 Sheet 文件: " + selected.getAttribute("name"));
            }
            target = target.replace('\\', '/');
            if (target.startsWith("/")) {
                return target.substring(1);
            }
            return target.startsWith("xl/") ? target : "xl/" + target;
        }

        private static Document readXml(ZipFile zip, String path) throws Exception {
            ZipEntry entry = zip.getEntry(path);
            if (entry == null) {
                throw new IllegalArgumentException("xlsx 内缺少文件: " + path);
            }
            byte[] bytes = readAll(zip.getInputStream(entry));
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            factory.setXIncludeAware(false);
            factory.setExpandEntityReferences(false);
            DocumentBuilder builder = factory.newDocumentBuilder();
            return builder.parse(new InputSource(new StringReader(new String(bytes, StandardCharsets.UTF_8))));
        }

        private static byte[] readAll(InputStream input) throws IOException {
            try (InputStream in = input; ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                byte[] buffer = new byte[8192];
                int length;
                while ((length = in.read(buffer)) >= 0) {
                    out.write(buffer, 0, length);
                }
                return out.toByteArray();
            }
        }

        private static String directChildText(Element element, String tagName) {
            NodeList children = element.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node child = children.item(i);
                if (child instanceof Element && tagName.equals(child.getNodeName())) {
                    return child.getTextContent();
                }
            }
            return "";
        }

        private static String textOfDescendant(Element element, String tagName) {
            StringBuilder text = new StringBuilder();
            NodeList nodes = element.getElementsByTagName(tagName);
            for (int i = 0; i < nodes.getLength(); i++) {
                text.append(nodes.item(i).getTextContent());
            }
            return text.toString();
        }

        private static int columnIndex(String cellRef) {
            int value = 0;
            for (int i = 0; i < cellRef.length(); i++) {
                char ch = cellRef.charAt(i);
                if (ch >= 'A' && ch <= 'Z') {
                    value = value * 26 + (ch - 'A' + 1);
                } else if (ch >= 'a' && ch <= 'z') {
                    value = value * 26 + (ch - 'a' + 1);
                } else {
                    break;
                }
            }
            return Math.max(0, value - 1);
        }

        private static int parseInt(String value, int defaultValue) {
            try {
                return isBlank(value) ? defaultValue : Integer.parseInt(value.trim());
            } catch (NumberFormatException ex) {
                return defaultValue;
            }
        }

        private static String normalizeNumber(String value) {
            if (value == null) {
                return "";
            }
            String trimmed = value.trim();
            if (trimmed.endsWith(".0")) {
                return trimmed.substring(0, trimmed.length() - 2);
            }
            return trimmed;
        }
    }

    private static class SqlWriter {
        private static String write(List<ErrorCodeRow> rows, Config config) {
            StringBuilder sql = new StringBuilder();
            appendDelete(sql, rows);
            appendInsert(sql, rows, "cn", config);
            appendInsert(sql, rows, "en", config);
            return sql.toString();
        }

        private static void appendDelete(StringBuilder sql, List<ErrorCodeRow> rows) {
            sql.append("-- delete  ------------------- \n");
            sql.append("DELETE FROM internationalization\n");
            sql.append("WHERE i18n_key IN (\n");
            for (int i = 0; i < rows.size(); i++) {
                sql.append("                   '").append(escape(rows.get(i).i18nKey)).append("'");
                sql.append(i == rows.size() - 1 ? ");\n\n" : ",\n");
            }
        }

        private static void appendInsert(StringBuilder sql, List<ErrorCodeRow> rows, String language, Config config) {
            sql.append("-- ").append(language).append(" -------------------\n");
            sql.append("INSERT INTO internationalization\n");
            sql.append("(i18n_key, i18n_code, i18n_value, is_delete, create_time, create_by, create_app, last_update_time, last_update_by, last_update_app, i18n_group)\n");
            sql.append("VALUES\n");
            for (int i = 0; i < rows.size(); i++) {
                ErrorCodeRow row = rows.get(i);
                String text = "cn".equals(language) ? row.cn : row.en;
                String value = row.code + " " + text;
                sql.append("    ('")
                        .append(escape(row.i18nKey)).append("', '")
                        .append(escape(language)).append("', '")
                        .append(escape(value)).append("', 0, '")
                        .append(escape(config.timestamp)).append("', '")
                        .append(escape(config.createBy)).append("', '")
                        .append(escape(config.createApp)).append("', '")
                        .append(escape(config.timestamp)).append("', '")
                        .append(escape(config.createBy)).append("', '")
                        .append(escape(config.createApp)).append("', '")
                        .append(escape(config.group)).append("')");
                sql.append(i == rows.size() - 1 ? ";\n\n" : ",\n");
            }
        }

        private static String escape(String value) {
            return value == null ? "" : value.replace("'", "''");
        }
    }
}
