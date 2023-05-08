package io.github.longyg.export;

import org.docx4j.TraversalUtil;
import org.docx4j.finders.SectPrFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.SectPr;

import java.util.ArrayList;
import java.util.List;

/**
 * Clean header and footer for word package if any
 *
 * @author longyg
 */
public class HeaderFooterCleaner {

    private HeaderFooterCleaner() {
    }

    public static void clean(WordprocessingMLPackage wordMLPackage) {
        if (null == wordMLPackage || null == wordMLPackage.getMainDocumentPart()) return;

        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();

        // Remove from sectPr
        SectPrFinder finder = new SectPrFinder(mdp);
        new TraversalUtil(mdp.getContent(), finder);
        for (SectPr sectPr : finder.getOrderedSectPrList()) {
            sectPr.getEGHdrFtrReferences().clear();
        }

        // Remove rels
        List<Relationship> hfRels = new ArrayList<>();
        for (Relationship rel : mdp.getRelationshipsPart().getRelationships().getRelationship()) {

            if (rel.getType().equals(Namespaces.HEADER)
                    || rel.getType().equals(Namespaces.FOOTER)) {
                hfRels.add(rel);
            }
        }
        for (Relationship rel : hfRels) {
            mdp.getRelationshipsPart().removeRelationship(rel);
        }
    }
}
