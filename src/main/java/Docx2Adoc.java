import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.OPCPackage;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;

public class Docx2Adoc {
    static Logger LOGGER;
    private XWPFDocument docxFile;
    private File adocFile;
    private String pathToAdoc;
    private FileWriter adocWriter;
    List<IBodyElement> docxElements;
    private XWPFStyles styles;

    public XWPFDocument getDocxFile() {
        return docxFile;
    }

    public File getAdocFile() {
        return adocFile;
    }

    public String getPathToAdoc() {
        return pathToAdoc;
    }

    public FileWriter getAdocWriter() {
        return adocWriter;
    }

    public List<IBodyElement> getDocxElements() {
        return docxElements;
    }

    public XWPFStyles getStyles() {
        return styles;
    }

    public Docx2Adoc(String pathDocx, String pathAdoc) {
        try (FileInputStream fileInputStream = new FileInputStream(pathDocx)) {

            adocFile = new File(pathAdoc);
            pathToAdoc = adocFile.getParentFile().toString();
            docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            styles = docxFile.getStyles();
            docxElements = docxFile.getBodyElements();

            try (FileInputStream ins = new FileInputStream("C:\\Users\\Genius\\IdeaProjects\\ASCIIDoc\\log.config")) {
                LogManager.getLogManager().readConfiguration(ins);
                LOGGER = Logger.getLogger(Main.class.getName());
            } catch (Exception logEx) {
                logEx.printStackTrace();
            }

            adocWriter = new FileWriter(adocFile, false);

        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    public XWPFRun checkSlash(XWPFRun run) {
        char[] runChars = run.text().toCharArray();
        if (runChars[runChars.length - 1] == '\\') {
            run.setText( run.text() + " ", 0);
        }
        return run;
    }

    public XWPFParagraph checkTextFormatting(XWPFParagraph paragraph) {
        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            XWPFRun run = paragraph.getRuns().get(i);
            if (run.isBold()) {
                if (i > 0) checkSlash(run);
                run.setText( "**" + run.text() + "**", 0);
            }
            if (run.isItalic()) {
                if (i > 0) checkSlash(run);
                run.setText( "__" + run.text() + "__", 0);
            }
        }
        return paragraph;
    }

    public void writeAdoc(XWPFParagraph paragraph, String type, BigInteger numFmt) throws IOException {
        if (paragraph.getText().equals("")) {
            return;
        }
        paragraph = checkTextFormatting(paragraph);
        String text = paragraph.getText();
        if (type != null) {
            String[] splitType = type.split(" ");
            if (type.equals("Body Text")) {
                adocWriter.write('\n' + text + '\n');
                adocWriter.flush();
                LOGGER.log(Level.INFO, "Успешная миграция параграфа");
            } else if (splitType[0].equals("heading")) {
                int lvl = Integer.parseInt(splitType[1]);
                String sectionLvl = ("=".repeat(Math.max(0, ++lvl)));
                adocWriter.write("\n\n" +sectionLvl + " " + text + '\n');
                adocWriter.flush();
                LOGGER.log(Level.INFO, "Успешная миграция заголовка с уровнем: " + --lvl);
            } else if (type.equals("Title")) {
                adocWriter.write("\n\n== " + text + '\n');
                adocWriter.flush();
                LOGGER.log(Level.INFO, "Успешная миграция заголовка");
            } else if (type.equals("Subtitle")) {
                adocWriter.write("\n==== " + text + '\n');
                adocWriter.flush();
                LOGGER.log(Level.INFO, "Успешная миграция подзаголовка");
            }
        } else if (numFmt != null) {
            numFmt = numFmt.add(BigInteger.TWO);
            String sectionLvl = ("=".repeat(Math.max(0, numFmt.intValue())));
            adocWriter.write("\n\n" + sectionLvl + " " + text + '\n');
            adocWriter.flush();
            LOGGER.log(Level.INFO, "Успешная миграция заголовка с уровнем: " + numFmt.subtract(BigInteger.ONE));
        }
    }

    public void writeListAdoc(XWPFParagraph paragraph, BigInteger numIlvl, String listType, boolean ifTableLog, boolean additionalP) throws IOException {
        paragraph = checkTextFormatting(paragraph);
        String text = paragraph.getText();
        int inumIlvlInt = 1;
        for (BigInteger a = BigInteger.ZERO;
             a.compareTo(numIlvl) < 0;
             a = a.add(BigInteger.ONE)) {
            inumIlvlInt++;
        }
        if (listType.equals("bullet")) {
            String bullets = ("*".repeat(Math.max(0, inumIlvlInt)));
            adocWriter.write('\n' + bullets + " " + text + '\n');
            if (additionalP)
                adocWriter.append("+\n");
            adocWriter.flush();
            if (ifTableLog) LOGGER.log(Level.INFO, "Успешная миграция элемента списка в ячейке таблицы");
            else LOGGER.log(Level.INFO, "Успешная миграция элемента списка");
        } else {
            String dots = (".".repeat(Math.max(0, inumIlvlInt)));
            adocWriter.write('\n' + dots + " " + text + '\n');
            if (additionalP)
                adocWriter.append("+\n");
            adocWriter.flush();
            if (ifTableLog) LOGGER.log(Level.INFO, "Успешная миграция элемента списка в ячейке таблицы");
            else LOGGER.log(Level.INFO, "Успешная миграция элемента списка");
        }
    }

    public void writeTableAdoc(XWPFTable table) throws IOException {
        List<BigInteger> columnsSize = new ArrayList<>();
        BigInteger tableColumnsSize = BigInteger.ZERO;
        for (XWPFTableCell tableCell : table.getRow(0).getTableCells()) {
            if (tableCell.getWidthType().toString().equals("DXA")) {
                columnsSize.add((BigInteger) tableCell.getCTTc().getTcPr().getTcW().getW());
                tableColumnsSize = tableColumnsSize.add(columnsSize.get(columnsSize.size() - 1));
            }
            else {
                columnsSize.add(BigInteger.ONE);
                tableColumnsSize = tableColumnsSize.add(BigInteger.ONE);
            }
        }

        adocWriter.write("\n[cols=\"");
        int pos = 0;
        for (BigInteger inter : columnsSize) {
            double inter1 = (inter.doubleValue() / (tableColumnsSize).doubleValue());
            if (inter1 != 1) inter1 = inter1 * 10;
            if (pos + 1 < columnsSize.size())
                adocWriter.write((int) Math.round(inter1) + "a,");
            else adocWriter.write((int) Math.round(inter1) + "a");
            columnsSize.set(pos++, BigInteger.valueOf(Math.round(inter1)));
        }
        adocWriter.write("\"]\n|===");
        adocWriter.flush();

        for (XWPFTableRow tableRow : table.getRows()) {
            adocWriter.append('\n');
            for (XWPFTableCell tableCell : tableRow.getTableCells()) {
                adocWriter.write("|");
                adocWriter.flush();
                for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                    if (paragraph.getCTPPr().getNumPr() != null) {
                        writeListAdoc(paragraph, paragraph.getNumIlvl(), paragraph.getNumFmt(), true, false);
                    } else {
                        if (checkPicture(paragraph) != null) {
                            writePictureAdoc(checkPicture(paragraph), true);
                        }
                        if (paragraph.getText().equals("")) {
                            continue;
                        }
                        adocWriter.write('\n' + paragraph.getText());
                        adocWriter.flush();
                        LOGGER.log(Level.INFO, "Успешная миграция параграфа в ячейке таблицы");
                    }
                }
            }
        }
        adocWriter.write("\n\n|===\n");
        adocWriter.flush();
        LOGGER.log(Level.INFO, "Успешная миграция таблицы");
    }

    public void writePictureAdoc(XWPFPicture picture, boolean ifTableLog) throws IOException {
        XWPFPictureData pictureData = picture.getPictureData();
        byte[] pictureBytes = pictureData.getData();
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(pictureBytes));
        File file = new File(pathToAdoc + "\\" + pictureData.getFileName());
        ImageIO.write(image, pictureData.suggestFileExtension(), file);
        adocWriter.write("\nimage::" + pictureData.getFileName() + "[]\n");
        adocWriter.flush();
        if (ifTableLog) LOGGER.log(Level.INFO, "Успешная миграция иллюстрации в ячейке таблицы");
        else LOGGER.log(Level.INFO, "Успешная миграция иллюстрации");
    }

    public XWPFPicture checkPicture(XWPFParagraph paragraph) {
        XWPFPicture picture = null;
        for (XWPFRun run : paragraph.getRuns()) {
            if (run.getEmbeddedPictures().isEmpty())
                continue;
            for (XWPFPicture runPicture : run.getEmbeddedPictures()) {
                picture = runPicture;
            }
        }
        return picture;
    }

    public void convert() throws IOException {
        for (int i = 0; i < docxElements.size(); i++) {
            IBodyElement docxElement = docxElements.get(i);
            if (docxElement instanceof XWPFParagraph paragraph) {
                if (checkPicture(paragraph) != null) {
                    writePictureAdoc(checkPicture(paragraph), false);
                    continue;
                }
                if (paragraph.getCTPPr().getNumPr() != null) {
                    int j = i;
                    int space = 0;
                    BigInteger numId = paragraph.getNumID();
                    while (j + 1 < docxElements.size()) {
                        j++;
                        if (docxElements.get(j) instanceof XWPFParagraph paragraph2 && paragraph2.getCTPPr().getNumPr() == null || docxElements.get(j) instanceof XWPFTable) {
                            if (docxElements.get(j) instanceof XWPFTable) space++;
                            else if (docxElements.get(j) instanceof XWPFParagraph paragraph2 && checkPicture(paragraph2) != null) space++;
                            else if (docxElements.get(j) instanceof XWPFParagraph paragraph2 && !paragraph2.getText().equals("")) space++;
                            else break;
                        }
                        else if (docxElements.get(j) instanceof XWPFParagraph paragraph2 && !numId.equals(paragraph2.getNumID())) {
                            space = 0;
                            break;
                        }
                    }
                    writeListAdoc(paragraph, paragraph.getNumIlvl(), paragraph.getNumFmt(), false, space == 1);
                }
                else if (paragraph.getStyleID() != null) {
                    String styleId = paragraph.getStyleID();
                    XWPFStyle style = styles.getStyle(styleId);
                    if (style != null) {
                        writeAdoc(paragraph, style.getName(), null);
                    }
                } else if (paragraph.getCTPPr().getOutlineLvl() != null) {
                    writeAdoc(paragraph, null, paragraph.getCTPPr().getOutlineLvl().getVal());
                } else {
                    writeAdoc(paragraph, "Body Text", null);
                }
            } else if (docxElement instanceof XWPFTable table) {
                writeTableAdoc(table);
            } else {
                LOGGER.log(Level.WARNING, "Неизвестый объект");
            }
        }
    }
}
