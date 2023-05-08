package io.github.longyg.export;

import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.util.List;

/**
 * @author longyg
 */
public class HeaderFooterCreator {
    private HeaderFooterCreator() {
    }

    public static void createHeader(WordprocessingMLPackage wordMLPackage, List<Object> objects) throws InvalidFormatException {
        Relationship part = createHeaderPart(wordMLPackage, objects);
        HeaderReference headerReference = createHeaderReference(part);
        SectPr sectPr = getSectPr(wordMLPackage);
        sectPr.getEGHdrFtrReferences().add(headerReference);
    }

    public static void createFooter(WordprocessingMLPackage wordMLPackage, List<Object> objects) throws InvalidFormatException {
        Relationship part = createFooterPart(wordMLPackage, objects);
        FooterReference footerReference = createFooterReference(part);
        SectPr sectPr = getSectPr(wordMLPackage);
        sectPr.getEGHdrFtrReferences().add(footerReference);
    }

    private static SectPr getSectPr(WordprocessingMLPackage wordMLPackage) {
        List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        // There is always a section wrapper, but it might not contain a sectPr
        if (sectPr == null) {
            sectPr = Context.getWmlObjectFactory().createSectPr();
            wordMLPackage.getMainDocumentPart().addObject(sectPr);
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }
        return sectPr;
    }

    private static HeaderReference createHeaderReference(Relationship relationship) {
        HeaderReference headerReference = Context.getWmlObjectFactory().createHeaderReference();
        headerReference.setId(relationship.getId());
        headerReference.setType(HdrFtrRef.DEFAULT);
        return headerReference;
    }

    private static FooterReference createFooterReference(Relationship relationship) {
        FooterReference footerReference = Context.getWmlObjectFactory().createFooterReference();
        footerReference.setId(relationship.getId());
        footerReference.setType(HdrFtrRef.DEFAULT);
        return footerReference;
    }

    private static Relationship createHeaderPart(WordprocessingMLPackage wordprocessingMLPackage, List<Object> objects)
            throws InvalidFormatException {
        HeaderPart headerPart = new HeaderPart();
        Relationship rel = wordprocessingMLPackage.getMainDocumentPart()
                .addTargetPart(headerPart);

        Hdr hdr = Context.getWmlObjectFactory().createHdr();
        hdr.getContent().addAll(objects);
        headerPart.setJaxbElement(hdr);
        return rel;
    }

    private static Relationship createFooterPart(WordprocessingMLPackage wordMLPackage, List<Object> objects)
            throws InvalidFormatException {
        FooterPart footerPart = new FooterPart();
        Relationship rel = wordMLPackage.getMainDocumentPart()
                .addTargetPart(footerPart);

        Ftr ftr = Context.getWmlObjectFactory().createFtr();
        ftr.getContent().addAll(objects);
        footerPart.setJaxbElement(ftr);
        return rel;
    }

}
