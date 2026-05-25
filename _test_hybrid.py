from ReferenceConversion import _lookup_journal_local

print("TEST 1: Local Lookup (Known Abbreviations)")
print("=" * 60)

test_cases = [
    ("Harv. L. Rev.", "Harvard Law Review"),
    ("J Affect Disord.", "Journal of Affective Disorders"),
    ("Yale L.J.", "Yale Law Journal"),
    ("discov psychol", "Discovery Psychology"),
    ("n.y.u. l. rev.", "New York University Law Review"),
    ("Unknown Abbrev.", None),
]

all_pass = True
for abbrev, expected in test_cases:
    result = _lookup_journal_local(abbrev)
    passed = result == expected
    all_pass = all_pass and passed
    status = "PASS" if passed else "FAIL"
    print(f"  [{status}] {abbrev:25} -> {result or 'NOT FOUND'}")
    if not passed:
        print(f"        Expected: {expected}")

print()
print("TEST 2: Abbreviation Detection")
print("=" * 60)

abbrev_patterns = [
    ("Harv. L. Rev.", True, "has dots"),
    ("J Affect Disord.", True, "has dots"),
    ("Journal of Affective Disorders", False, "full name, multiple words"),
    ("Nature", False, "single word, full name"),
    ("J. Clin. Psychiatry", True, "abbreviated"),
]

for name, should_be_abbrev, reason in abbrev_patterns:
    word_count = len(name.split())
    has_dots = "." in name
    is_single_word_caps = word_count == 1 and name.isupper() and len(name) <= 4
    is_abbrev = has_dots or is_single_word_caps
    status = "PASS" if is_abbrev == should_be_abbrev else "FAIL"
    print(f"  [{status}] {name:35} is_abbrev={is_abbrev} ({reason})")

print()
print("=" * 60)
all_pass_abbrev = True
for name, should_be, _ in abbrev_patterns:
    word_count = len(name.split())
    has_dots = "." in name
    is_single_word_caps = word_count == 1 and name.isupper() and len(name) <= 4
    is_abbrev = has_dots or is_single_word_caps
    if is_abbrev != should_be:
        all_pass_abbrev = False

print("Local Lookup:", "ALL PASSED" if all_pass else "SOME FAILED")
print("Abbrev Detection:", "ALL PASSED" if all_pass_abbrev else "SOME FAILED")
