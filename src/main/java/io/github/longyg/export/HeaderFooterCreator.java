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

    public static void addHeaderPart(WordprocessingMLPackage wordMLPackage,
                                     HeaderPart headerPart,
                                     Relationship relationship,
                                     List<Object> objects) {
        Hdr hdr = Context.getWmlObjectFactory().createHdr();
        hdr.getContent().addAll(objects);
        headerPart.setJaxbElement(hdr);

        HeaderReference headerReference = createHeaderReference(relationship);
        SectPr sectPr = getSectPr(wordMLPackage);
        sectPr.getEGHdrFtrReferences().add(headerReference);
    }

    public static void addFooterPart(WordprocessingMLPackage wordMLPackage,
                                     FooterPart footerPart,
                                     Relationship relationship,
                                     List<Object> objects) {
        Ftr ftr = Context.getWmlObjectFactory().createFtr();
        ftr.getContent().addAll(objects);
        footerPart.setJaxbElement(ftr);

        FooterReference footerReference = createFooterReference(relationship);
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
}
