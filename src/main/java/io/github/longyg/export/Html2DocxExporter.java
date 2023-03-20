package io.github.longyg.export;

import io.github.longyg.export.exception.ExportException;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * Export HTML to docx using docx4j-xhtml2word library
 *
 * @author longyg
 */
@Slf4j
public class Html2DocxExporter {
    private static final String EXPORT_ERROR_MSG = "Exception while exporting HTML to word";

    private String hyperlinkStyleId = "Hyperlink";

    public String getHyperlinkStyleId() {
        return hyperlinkStyleId;
    }

    public void setHyperlinkStyleId(String hyperlinkStyleId) {
        this.hyperlinkStyleId = hyperlinkStyleId;
    }

    /**
     * Export input stream of HTML to output stream of docx
     *
     * @param html input stream of HTML
     * @param os   output stream of generated docx
     */
    public void export(InputStream html, OutputStream os) throws ExportException {
        try {
            export(html, os, WordprocessingMLPackage.createPackage());
        } catch (Exception e) {
            log.error(EXPORT_ERROR_MSG, e);
            throw new ExportException(e);
        }
    }

    /**
     * Export input stream of HTML to output stream of docx, using a docx as base
     *
     * @param html     input stream of HTML
     * @param baseDocx input stream of base docx
     * @param os       output stream of generated docx
     */
    public void export(InputStream html, InputStream baseDocx, OutputStream os) throws ExportException {
        try {
            WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(baseDocx);
            preprocess(wmlPackage);
            export(html, os, wmlPackage);
        } catch (Docx4JException e) {
            log.error(EXPORT_ERROR_MSG, e);
            throw new ExportException(e);
        }
    }

    /**
     * Export HTML string to output stream of docx, using original docx as base
     *
     * @param htmlString HTML string
     * @param baseDocx   input stream of base docx
     * @param os         output stream of generated docx
     */
    public void export(String htmlString, InputStream baseDocx, OutputStream os) throws ExportException {
        try {
            WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(baseDocx);
            preprocess(wmlPackage);
            export(htmlString, os, wmlPackage);
        } catch (Exception e) {
            log.error(EXPORT_ERROR_MSG, e);
            throw new ExportException(e);
        }
    }

    private void export(InputStream is, OutputStream os,
                        WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        XHTMLImporterImpl xhtmlImporter = getXHTMLImporterImpl(wordMLPackage);
        saveResult(wordMLPackage, os, xhtmlImporter.convert(is, null));
    }

    private void export(String html, OutputStream os, WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        XHTMLImporterImpl xhtmlImporter = getXHTMLImporterImpl(wordMLPackage);
        saveResult(wordMLPackage, os, xhtmlImporter.convert(html, null));
    }

    private XHTMLImporterImpl getXHTMLImporterImpl(WordprocessingMLPackage wordMLPackage) {
        XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordMLPackage);
        xhtmlImporter.setHyperlinkStyle(getHyperlinkStyleId());
        return xhtmlImporter;
    }

    private static void saveResult(WordprocessingMLPackage wordMLPackage,
                                   OutputStream os, List<Object> result) throws Docx4JException {
        wordMLPackage.getMainDocumentPart().getContent().clear();
        wordMLPackage.getMainDocumentPart().getContent().addAll(result);
        wordMLPackage.save(os);
    }

    protected void preprocess(WordprocessingMLPackage wmlPackage) {
        // do nothing here
        // could be overwritten by subclass if you want to do something for the base docx
    }
}
