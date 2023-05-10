package io.github.longyg.export;

import io.github.longyg.export.exception.ExportException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.relationships.Relationship;

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
        export(html, null, os);
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
            if (null == baseDocx) {
                export(html, os, WordprocessingMLPackage.createPackage());
            } else {
                WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(baseDocx);
                preprocess(wmlPackage);
                export(html, os, wmlPackage);
            }
        } catch (Docx4JException e) {
            log.error(EXPORT_ERROR_MSG, e);
            throw new ExportException(e);
        }
    }

    /**
     * Export HTML string to output stream of docx
     *
     * @param htmlString HTML string
     * @param os         output stream of generated docx
     */
    public void export(String htmlString, OutputStream os) throws ExportException {
        export(htmlString, null, os);
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
            if (null == baseDocx) {
                export(htmlString, os, WordprocessingMLPackage.createPackage());
            } else {
                WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(baseDocx);
                preprocess(wmlPackage);
                export(htmlString, os, wmlPackage);
            }
        } catch (Exception e) {
            log.error(EXPORT_ERROR_MSG, e);
            throw new ExportException(e);
        }
    }

    /**
     * Export HTML string with header and footer to output stream of docx
     *
     * @param htmlString HTML string
     * @param header     header
     * @param footer     footer
     * @param os         output stream of generated docx
     * @throws ExportException
     */
    public void export(String htmlString, String header, String footer, OutputStream os) throws ExportException {
        export(htmlString, header, footer, null, os);
    }

    /**
     * Export HTML string with header and footer to output stream of docx, using original docx as base
     *
     * @param htmlString HTML string
     * @param header     header
     * @param footer     footer
     * @param baseDocx   input stream of base docx
     * @param os         output stream of generated docx
     * @throws ExportException
     */
    public void export(String htmlString, String header, String footer, InputStream baseDocx, OutputStream os) throws ExportException {
        try {
            if (null == baseDocx) {
                export(htmlString, header, footer, os, WordprocessingMLPackage.createPackage());
            } else {
                WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(baseDocx);
                preprocess(wmlPackage);
                export(htmlString, header, footer, os, wmlPackage);
            }
        } catch (Exception e) {
            log.error(EXPORT_ERROR_MSG, e);
            throw new ExportException(e);
        }
    }

    private void export(InputStream is, OutputStream os,
                        WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        XHTMLImporterImpl xhtmlImporter = getXHTMLImporter(wordMLPackage);
        saveResult(wordMLPackage, os, xhtmlImporter.convert(is, null));
    }

    private void export(String html, OutputStream os, WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        XHTMLImporterImpl xhtmlImporter = getXHTMLImporter(wordMLPackage);
        saveResult(wordMLPackage, os, xhtmlImporter.convert(html, null));
    }

    private void export(String html, String header, String footer,
                        OutputStream os, WordprocessingMLPackage wordMLPackage)
            throws Docx4JException {
        // if there is header and footer to be generated, we need first clean the existing header and footer from word package if any
        if (!StringUtils.isEmpty(header) || !StringUtils.isEmpty(footer)) {
            HeaderFooterCleaner.clean(wordMLPackage);
        }
        if (!StringUtils.isEmpty(header)) {
            // convert header and save to word package
            HeaderPart headerPart = new HeaderPart();
            Relationship relationship = wordMLPackage.getMainDocumentPart().addTargetPart(headerPart);
            XHTMLImporterImpl importer = getXHTMLImporter(wordMLPackage, headerPart);
            HeaderFooterCreator.addHeaderPart(wordMLPackage, headerPart, relationship, importer.convert(header, null));
        }
        if (!StringUtils.isEmpty(footer)) {
            // convert footer and save to word package
            FooterPart footerPart = new FooterPart();
            Relationship relationship = wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
            XHTMLImporterImpl importer = getXHTMLImporter(wordMLPackage, footerPart);
            HeaderFooterCreator.addFooterPart(wordMLPackage, footerPart, relationship, importer.convert(footer, null));
        }

        if (!StringUtils.isEmpty(html)) {
            // convert main html and save to word package
            addMainDocument(wordMLPackage, getXHTMLImporter(wordMLPackage).convert(html, null));
        }

        // save word package to output stream
        wordMLPackage.save(os);
    }

    private void addMainDocument(WordprocessingMLPackage wordMLPackage, List<Object> objects) {
        wordMLPackage.getMainDocumentPart().getContent().clear();
        wordMLPackage.getMainDocumentPart().getContent().addAll(objects);
    }

    private XHTMLImporterImpl getXHTMLImporter(WordprocessingMLPackage wordMLPackage, Part sourcePart) {
        XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage, sourcePart);
        importer.setHyperlinkStyle(getHyperlinkStyleId());
        return importer;
    }

    private XHTMLImporterImpl getXHTMLImporter(WordprocessingMLPackage wordMLPackage) {
        XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage);
        importer.setHyperlinkStyle(getHyperlinkStyleId());
        return importer;
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
