package org.docx4j.convert.out.html;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.fonts.BestMatchingMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class OMMLTest {

    static final String MATH_ML_STR = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mrow><mfenced separators=\"|\"><mrow><mi>x</mi><mo>+</mo><mi>a</mi></mrow></mfenced></mrow><mrow><mi>n</mi></mrow></msup><mo>=</mo><mrow><msubsup><mo stretchy=\"false\">∑</mo><mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mrow><mi>n</mi></mrow></msubsup><mrow><mfenced separators=\"|\"><mrow><mfrac linethickness=\"0pt\"><mrow><mi>n</mi></mrow><mrow><mi>k</mi></mrow></mfrac></mrow></mfenced><msup><mrow><mi>x</mi></mrow><mrow><mi>k</mi></mrow></msup><msup><mrow><mi>a</mi></mrow><mrow><mi>n</mi><mo>-</mo><mi>k</mi></mrow></msup></mrow></mrow></math>";
    @Test
    public void ommlIsConvertedToMathML() throws Exception {
        String filePath = System.getProperty("user.dir") + "/sample_docs/equation.docx";
        Assert.assertTrue(Files.exists(Path.of(filePath)));
        String newFilePath = tmpDir() + "docx4j-omml-test.html";
        genFile(filePath, newFilePath);
        Path htmPath = Path.of(newFilePath);
        Assert.assertTrue(Files.exists(htmPath));
        var htmString = new String(Files.readAllBytes(htmPath));
        var strs = this.extractMathElements(htmString);

        Assert.assertEquals(1, strs.size());
        Assert.assertEquals(MATH_ML_STR, strs.get(0));
    }


    private List<String> extractMathElements(String html) {
        int fromIndex = 0, toIndex = 0, initial = 0;
        List<String> list = new ArrayList<>();
        while (true) {
            fromIndex = html.indexOf(START_INDEX_STR, initial);
            if (fromIndex == -1) {
                break;
            }
            toIndex = html.indexOf(END_INDEX_STR, fromIndex);
            if (toIndex == -1) {
                throw new IllegalStateException(String.format("end index during xml parse is -1 for str: %s", html));
            }
            list.add(html.substring(fromIndex, toIndex + END_INDEX_STR.length()));
            initial = toIndex + END_INDEX_STR.length();
        }
        return list;
    }

    private static final String START_INDEX_STR = "<math";

    private static final String END_INDEX_STR = "</math>";

    private String tmpDir() {
        if(System.getProperty("os.name").toLowerCase().contains("windows")) {
            return "C:\\Windows\\Temp\\";
        }
        return "/tmp/";
    }

    private static void genFile(String inputFilePath, String newFileName) throws Exception {

        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(inputFilePath));

        HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
        String imagePath = newFileName + "_files";
        htmlSettings.setImageDirPath(imagePath);
        String uri = newFileName.substring(newFileName.lastIndexOf("/") + 1) + "_files";
        htmlSettings.setImageTargetUri(uri);
        htmlSettings.setOpcPackage(wordMLPackage);
        // use browser defaults for ol, ul, li
        String userCSS = "html, body, div, span, h1, h2, h3, h4, h5, h6, p, a, table, caption, tbody, tfoot, thead, tr, th, td " +
                "body { margin: 0;padding: 0; border: 0;}" +
                "body { margin-left: 30px; margin-top: 30px; width: 612.43px; line-height: 0.82 !important; -webkit-text-stroke-width: 0.2px !important; font-size: 14.97px !important}"  +
                "span[style*='vertical-align: sub'] { position: relative !important;top: -3px !important; font-size: 9.35px;}" +
                "span[style*='vertical-align: super'] { font-size: 9.35px }";
        // TODO: fork docx 4j and add super script/ subscript px values


        htmlSettings.setUserCSS(userCSS);

        // SdtWriter.registerTagHandler("HTML_ELEMENT", customSDTListHandler);
        Mapper fontMapper = new BestMatchingMapper(); // better for Linux
        wordMLPackage.setFontMapper(fontMapper);

        // .. example of mapping font Times New Roman which doesn't have certain Arabic glyphs
        // eg Glyph "ي" (0x64a, afii57450) not available in font "TimesNewRomanPS-ItalicMT".
        // eg Glyph "ج" (0x62c, afii57420) not available in font "TimesNewRomanPS-ItalicMT".
        // to a font which does
        PhysicalFont font = PhysicalFonts.get("Arial Unicode MS");
        OutputStream outputStream = new FileOutputStream(newFileName);
        Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.", true);

        Docx4J.toHTML(htmlSettings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);
        outputStream.close();
        //        Docx4J.toHTML(wordMLPackage, inputFilePath,
//                uri, outputStream);

    }



}
