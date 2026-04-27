#!/usr/bin/env python3
"""
Test script to validate the Abuhamad document with the new features
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from Referencenumvalidation import ReferenceProcessor
from docx import Document

# Test file path
TEST_FILE = "S4C-Processed-Documents/e9b4386288e145cf890dcf6123167849/Abuhamad9781975242831-ch002-Tagged.docx"

def test_document():
    print("=" * 70)
    print("TESTING: Abuhamad9781975242831-ch002-Tagged.docx")
    print("=" * 70)

    if not os.path.exists(TEST_FILE):
        print(f"[FAIL] File not found: {TEST_FILE}")
        return False

    try:
        # Load document
        doc = Document(TEST_FILE)
        processor = ReferenceProcessor(doc)

        print("\n[1] DOCUMENT STRUCTURE")
        print(f"    Total paragraphs: {len(doc.paragraphs)}")

        # Get bibliography info
        refs_found, ref_objects = processor.get_references_in_bibliography()
        print(f"    Bibliography entries (REF-N): {len(refs_found)}")
        print(f"    Reference IDs found: {sorted(refs_found)}")

        # Get citation info
        all_cited, appearance_order = processor.get_citations_in_text()
        unique_cited = set(all_cited)
        print(f"    Total citations: {len(all_cited)}")
        print(f"    Unique citations: {len(unique_cited)}")
        print(f"    Citation order: {appearance_order}")

        # Get validation stats
        print("\n[2] VALIDATION STATS (BEFORE)")
        stats = processor.get_validation_stats()
        print(f"    Missing references: {stats['missing_references']}")
        print(f"    Unused references: {stats['unused_references']}")
        print(f"    Duplicate references: {len(stats['duplicate_references'])} found")

        if stats['duplicate_references']:
            print("\n    DUPLICATES DETECTED:")
            for dup in stats['duplicate_references']:
                print(f"      - Ref {dup['id']} is duplicate of Ref {dup['duplicate_of']} ({dup['score']}% similar)")

        print(f"    Sequence issues: {len(stats['sequence_issues'])} found")
        if stats['sequence_issues']:
            print("\n    SEQUENCE ISSUES:")
            for issue in stats['sequence_issues']:
                print(f"      - Position {issue['position']}: expected {issue['expected']}, got {issue['current']}")

        print(f"    Is perfect: {stats['is_perfect']}")

        # Test duplicate resolution
        if stats['duplicate_references']:
            print("\n[3] RESOLVING DUPLICATES")
            merge_log = processor.resolve_duplicates()
            print(f"    Duplicates resolved: {len(merge_log)}")
            for merge in merge_log:
                print(f"      - Ref {merge['removed_id']} -> Ref {merge['canonical_id']} ({merge['score']}% similar)")
                print(f"        {merge['citations_updated']} citations updated")

            # Re-validate after merge
            print("\n[4] VALIDATION STATS (AFTER MERGE)")
            stats_after = processor.get_validation_stats()
            print(f"    Missing references: {stats_after['missing_references']}")
            print(f"    Unused references: {stats_after['unused_references']}")
            print(f"    Duplicate references: {len(stats_after['duplicate_references'])}")
            print(f"    Sequence issues: {len(stats_after['sequence_issues'])}")

        # Test renumbering
        print("\n[5] RENUMBERING")
        mapping = processor.renumber()
        print(f"    Renumbering mapping: {mapping if mapping else 'No changes needed'}")

        # Final validation
        print("\n[6] FINAL VALIDATION STATS")
        stats_final = processor.get_validation_stats()
        print(f"    Missing references: {stats_final['missing_references']}")
        print(f"    Unused references: {stats_final['unused_references']}")
        print(f"    Is perfect: {stats_final['is_perfect']}")

        print("\n" + "=" * 70)
        print("[PASS] TEST COMPLETED SUCCESSFULLY!")
        print("=" * 70)
        return True

    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_document()
    sys.exit(0 if success else 1)
