import argparse
import zipfile
from lxml import etree
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor


def ensure_custom_styles_exist(docx_path):
    """Ensure the 'CommentHighlight' and 'CommentQuery' character styles exist."""
    doc = Document(docx_path)
    styles = doc.styles

    if "CommentHighlight" not in styles:
        highlight_style = styles.add_style(
            "CommentHighlight", WD_STYLE_TYPE.CHARACTER
        )
        highlight_style.font.color.rgb = RGBColor(180, 0, 0)  # Dark Red
        highlight_style.font.underline = True

    if "CommentQuery" not in styles:
        query_style = styles.add_style("CommentQuery", WD_STYLE_TYPE.CHARACTER)
        query_style.font.color.rgb = RGBColor(0, 70, 180)  # Dark Blue
        query_style.font.italic = True
        query_style.font.bold = True

    doc.save(docx_path)


def convert_comments_to_inline_styled_text(input_docx, output_docx):
    """Reads Word document comments, highlights body text, inserts comment text

    inline using custom character styles, and outputs a valid .docx file.
    """
    # 1. First ensure styles exist in the input file
    ensure_custom_styles_exist(input_docx)

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    with zipfile.ZipFile(input_docx, "r") as docx_in:
        file_list = docx_in.namelist()

        # 2. Extract comments map: { comment_id: comment_text }
        comments_dict = {}
        if "word/comments.xml" in file_list:
            comments_tree = etree.fromstring(docx_in.read("word/comments.xml"))
            for cmt in comments_tree.xpath("//w:comment", namespaces=ns):
                c_id = cmt.get(f"{{{ns['w']}}}id")
                texts = cmt.xpath(".//w:t/text()", namespaces=ns)
                comments_dict[c_id] = "".join(texts).strip()

        # 3. Modify document.xml
        doc_tree = etree.fromstring(docx_in.read("word/document.xml"))
        active_comment_ids = set()

        # Elements to remove after processing (to prevent corruption)
        nodes_to_remove = []

        for p in doc_tree.xpath("//w:p", namespaces=ns):
            for child in list(p):
                tag = child.tag.split("}")[-1]

                # Start of comment range
                if tag == "commentRangeStart":
                    c_id = child.get(f"{{{ns['w']}}}id")
                    active_comment_ids.add(c_id)
                    nodes_to_remove.append(child)

                # End of comment range -> Insert Query text inline
                elif tag == "commentRangeEnd":
                    c_id = child.get(f"{{{ns['w']}}}id")
                    nodes_to_remove.append(child)

                    if c_id in comments_dict and comments_dict[c_id]:
                        # Create run (<w:r>) with style 'CommentQuery'
                        query_run = etree.Element(f"{{{ns['w']}}}r")
                        rPr = etree.SubElement(query_run, f"{{{ns['w']}}}rPr")
                        rStyle = etree.SubElement(rPr, f"{{{ns['w']}}}rStyle")
                        rStyle.set(f"{{{ns['w']}}}val", "CommentQuery")

                        t = etree.SubElement(query_run, f"{{{ns['w']}}}t")
                        t.set(
                            "{http://www.w3.org/XML/1998/namespace}space",
                            "preserve",
                        )
                        t.text = f" [QUERY: {comments_dict[c_id]}]"

                        child.addnext(query_run)

                    if c_id in active_comment_ids:
                        active_comment_ids.remove(c_id)

                # Reference markers inside text (commentReference)
                elif tag == "r":
                    # Check for comment reference icons inside runs and remove them
                    comment_refs = child.xpath(
                        ".//w:commentReference", namespaces=ns
                    )
                    if comment_refs:
                        nodes_to_remove.extend(comment_refs)

                    # Apply character style 'CommentHighlight' if inside active comment range
                    if active_comment_ids:
                        rPr = child.find(f"{{{ns['w']}}}rPr")
                        if rPr is None:
                            rPr = etree.Element(f"{{{ns['w']}}}rPr")
                            child.insert(0, rPr)

                        rStyle = rPr.find(f"{{{ns['w']}}}rStyle")
                        if rStyle is None:
                            rStyle = etree.SubElement(
                                rPr, f"{{{ns['w']}}}rStyle"
                            )
                        rStyle.set(f"{{{ns['w']}}}val", "CommentHighlight")

        # Safely remove comment marker tags from document tree
        for node in nodes_to_remove:
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)

        # 4. Save updated ZIP archive without deleting internal docx structures
        with zipfile.ZipFile(
            output_docx, "w", zipfile.ZIP_DEFLATED
        ) as docx_out:
            for item in file_list:
                if item == "word/document.xml":
                    docx_out.writestr(
                        item,
                        etree.tostring(
                            doc_tree, xml_declaration=True, encoding="UTF-8"
                        ),
                    )
                else:
                    docx_out.writestr(item, docx_in.read(item))

    print(
        f"? Successfully converted without corruption!\nSaved to: {output_docx}"
    )


def main():
    parser = argparse.ArgumentParser(
        description="Convert Word comments to inline styled text safely."
    )
    parser.add_argument(
        "-i", "--input", required=True, help="Path to input .docx file"
    )
    parser.add_argument(
        "-o", "--output", required=True, help="Path to output .docx file"
    )

    args = parser.parse_args()
    convert_comments_to_inline_styled_text(args.input, args.output)


if __name__ == "__main__":
    main()