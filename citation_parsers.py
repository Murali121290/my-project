"""
Citation Style Parsers for Citation Checker

This module provides parsers for different citation styles including:
- APA (American Psychological Association)
- Vancouver (Numeric and Author-Year variants)
- Chicago (Author-Year)

Each parser can extract citations from text and parse bibliography entries.
"""

import re
from typing import List, Tuple, Dict, Optional


class CitationParser:
    """Base class for citation parsers."""
    
    def __init__(self, style_name: str):
        self.style_name = style_name
    
    def parse_citation(self, text: str) -> List[Dict]:
        """
        Parse citations from text.
        
        Args:
            text: Text containing citations
            
        Returns:
            List of dicts with keys: 'author', 'year', 'type' (parenthetical/narrative), 
            'warnings' (list), and 'raw' (original text)
        """
        raise NotImplementedError
    
    def parse_reference(self, text: str) -> Optional[Dict]:
        """
        Parse a bibliography reference.
        
        Args:
            text: Reference text from bibliography
            
        Returns:
            Dict with keys 'author', 'year', 'full_author', 'abbreviations', or None if invalid
        """
        raise NotImplementedError


class APACitationParser(CitationParser):
    """Parser for APA (American Psychological Association) style citations."""
    
    def __init__(self):
        super().__init__("APA")
        # Updated year pattern to CAPTURE the optional letter suffix (e.g., 2024a)
        self.year_pattern = re.compile(r'\b((?:19|20)\d{2}[a-z]?)\b')
        # Updated pattern to support:
        # (2020)
        # (2020, January 15)
        # (n.d.) or (n.d)
        # (in press)
        # Optional period at end to be robust
        # Updated pattern to robustly capture the Year block.
        # It looks for a parenthetical group starting with a year (19xx/20xx), n.d., or in press.
        # This prevents greedily capturing "(CDC)" or "(Abbr)" as the date part.
        self.reference_pattern = re.compile(r'^(.+?)\s*\(((?:(?:19|20)\d{2}[a-z]?|n\.?d\.?(?:-[a-z])?|in press).*?)\)\.?')
    
    def parse_citation(self, cite_text: str) -> List[Dict]:
        """
        Parse APA-style citations from text.
        Supports:
        - Parenthetical: (Smith, 2020) embedded in text
        - Narrative: Smith (2020) embedded in text
        
        Args:
            cite_text: Text segment containing citations
            
        Returns:
            List of dicts with: author, year, type, warnings
        """
        results = []
        cite_text_clean = cite_text.strip()
        
        # Track parenthetical citations (author, year pairs) to avoid duplicates in narrative parsing
        parenthetical_pairs = set()
        
        # Check for Month names (exclusion) - simplistic check, might need refinement so valid citations aren't skipped if paragraph mentions a month.
        # But if the PASSAGE is just a date "January 2020", we skip.
        # If it's a paragraph, we shouldn't return [] immediately unless the MATCH is a date.
        # Removing global month check for paragraph-level parsing.
            
        # 1. Parse Parenthetical Citations: (Smith, 2020)
        # PRIORITY: Parse parentheses first, collect valid citations
        # Find content inside parentheses
        parenthetical_pattern = re.compile(r'\(([^()]+)\)')
        
        for match in parenthetical_pattern.finditer(cite_text_clean):
             raw_content = match.group(1).strip()
             full_match = match.group(0) # We might need to approximate raw text for sub-segments
             
             # Split by semicolon for multi-citations
             segments = [s.strip() for s in raw_content.split(';') if s.strip()]
             
             for segment in segments:
                 # Skip if it is likely NOT a citation
                 # 1. Page numbers only: (p. 12)
                 if re.match(r'^p\.?\s*\d+', segment, re.IGNORECASE):
                     continue
                 # 2. Dates only: (January 2020)
                 if re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)', segment, re.IGNORECASE):
                     continue
                     
                 # Check for common format errors first
                 warnings = []
                 
                 # Missing comma: (Smith 2020) - but ensures it looks like Name Year
                 if re.search(r'^[A-Z][a-z]+\s+\d{4}$', segment):
                     warnings.append("Format Error: Missing comma between author and year (APA requires comma)")
                 
                 # Wrong conjunction: (Smith and Jones, 2020)
                 if ' and ' in segment and '&' not in segment:
                     warnings.append("Format Error: Use '&' inside parentheses, not 'and'")
                     
                 # Now extract years (including n.d. and in press)
                 years = self.year_pattern.findall(segment)
                 
                 # Also check for n.d. (with optional suffix like n.d.-a) or in press
                 nd_match = re.search(r'\bn\.?d\.?(?:-[a-z])?\b', segment, re.IGNORECASE)
                 has_in_press = bool(re.search(r'\bin\s+press\b', segment, re.IGNORECASE))
                 
                 if nd_match:
                     # Capture the actual n.d. string (e.g. "n.d.-a") as the year
                     years = [nd_match.group(0)] 
                 elif has_in_press:
                     years = ['in press']
                 
                 # Remove prefixes
                 # Match longest prefixes first to prevent partial matches (e.g. "see, for example" vs "see")
                 clean_content = re.sub(r'^(?:see, for example,|see,\s*for example|see also|for example|also|see|cf\.?|e\.g\.?,?|i\.e\.?,?)(?:,)?\s+', '', segment, flags=re.IGNORECASE).strip()
                 
                 # Strip page numbers aggressively (p. 23, pp. 40-41) anywhere in the string
                 clean_content = re.sub(r',?\s*\bpp?\.?\s*\d+[-–]?\d*', '', clean_content, flags=re.IGNORECASE).strip()
                 
                 # Extract Author part by removing years (including n.d. and in press)
                 author_part = clean_content
                 if years:
                     # Remove regular years
                     author_part = re.sub(r',?\s*(?:19|20)\d{2}[a-z]?,?\s*', '', author_part).strip()
                     # Remove n.d. (and suffix) and in press
                     author_part = re.sub(r',?\s*\bn\.?d\.?(?:-[a-z])?\b,?\s*', '', author_part, flags=re.IGNORECASE).strip()
                     author_part = re.sub(r',?\s*\bin\s+press\b,?\s*', '', author_part, flags=re.IGNORECASE).strip()
                 
                 # Strip leading/trailing punctuation (like comma from 'see, for example, Author')
                 author_part = re.sub(r'^[.,;: ]+', '', author_part).strip()
                 author_part = re.sub(r'[.,;: ]+$', '', author_part).strip()
                 # author_part = re.sub(r',\s*$', '', author_part).strip() # Handled above
                 author_part = re.sub(r'^\d+,?\s*', '', author_part).strip()
                 
                 # Handle Year-Only Citations (e.g. (1991), (1982-1990))
                 # Treat empty or non-letter strings as "Unknown" if years exist
                 # BUT: Suppress if preceded by an excluded abbreviation (e.g. "PMHNP (2020)")
                 # We don't want "Unknown (2020)" if we intentionally ignored "PMHNP".
                 
                 preceding_text = cite_text_clean[:match.start()].strip()
                 # Check if preceding text looks like a Narrative Author (ends with Capitalized Word)
                 # If so, we assume the Narrative Parser will handle it (either accept or exclude).
                 # This prevents:
                 # 1. Duplicates: Smith (2020) -> Narrative identifies "Smith", Parenthetical shouldn't add "Unknown".
                 # 2. Ignored Abbreviations: PMHNP (2020) -> Narrative excludes "PMHNP", Parenthetical shouldn't add "Unknown".
                 # 3. Et al: Browne et al. (2017) -> Narrative identifies "Browne et al.", Parenthetical shouldn't add "Unknown".
                 # 4. Possessives: Rumbaut's (2005) -> Narrative identifies "Rumbaut's", Parenthetical shouldn't add "Unknown".
                 
                 preceded_by_potential_author = False
                 # Check for capitalized word (Name) OR "et al."/ "et al" OR possessive forms (Name's)
                 # Updated pattern to include possessive apostrophes: ['']s?
                 if re.search(r"([A-Z][\w\.]*[''']?s?|et al\.?)\s*$", preceding_text):
                     preceded_by_potential_author = True
                 
                 if years and (not author_part or not re.search(r'[A-Za-z]', author_part)):
                     if preceded_by_potential_author:
                         continue # Skip "Year Only" logic, defer to Narrative Parser
                     
                     author_part = "Unknown"
                     warnings.append("Warning: Missing Author")
                 
                 if author_part and len(author_part) > 1:
                     # Exclusions: Table, Figure, Ed/Eds/Vol/Suppl, Appendix, Chapter, etc.
                     # Also excluding provided false positives: "between", "except", "Ruth", "Johnny", "Z"
                     if re.match(r'^(Table|Figure|Fig|Eds?|Vol|Suppl|Appendix|Chapter|Section|Part|between|except|Ruth|Johnny|Z\.?|UK)\b', author_part, re.IGNORECASE):
                         continue
                     
                     # Secondary Citation Check: (Beckman in Shimrat, 1997) -> Match Shimrat
                     secondary_match = re.search(r'\b(?:as cited in|in)\s+([A-Z][A-Za-z\s&]+)', author_part)
                     if secondary_match:
                         author_part = secondary_match.group(1).strip()
                         
                     if years:
                         # Garbage Filter: Check if author_part has any letters
                         # Note: "Unknown" has letters, so it passes.
                         # Real garbage like "---" without replacement would be filtered, but we replaced it with Unknown above.
                         if not re.search(r'[A-Za-z]', author_part):
                             continue

                         # Case: Author, 2019, 2020 -> Multiple citations
                         for year in years:
                             results.append({
                                 'author': author_part,
                                 'year': year,
                                 'type': 'parenthetical',
                                 'warnings': warnings,
                                 'raw': f"({segment})" # Approx raw
                             })
                             # Track this author-year pair to skip in narrative parsing
                             # Normalize: strip trailing periods/punctuation for consistent matching
                             author_normalized = re.sub(r'[.,;: ]+$', '', author_part).strip()
                             parenthetical_pairs.add((author_normalized.lower(), year))
                     # MISSING YEAR DETECTION DISABLED PER USER REQUEST
                     # else: pass
                         
        # 2. Parse Narrative Citations: Smith (2020)
        # Regex for Name (Content with Year)
        # Restrict to ~100 chars max to avoid capturing entire sentences
        # Allow dots (et al.), quotes (possessives) in name.
        # Allow extra content in parens (p. numbers, multiple years).
        # ADDED HYPHEN SUPPORT
        narrative_pattern = re.compile(r"\b([A-Z][A-Za-z\s&.'\"`’–-]{0,100}?)\s*\(([^)]+)\)")
        
        for match in narrative_pattern.finditer(cite_text_clean):
            author_raw = match.group(1).strip()
            parens_content = match.group(2)
            
            # Check if parens content actually contains a year (or n.d. or in press)
            years = self.year_pattern.findall(parens_content)
            
            # Also check for n.d. or in press
            has_nd = bool(re.search(r'\bn\.?\s*d\.?\b', parens_content, re.IGNORECASE))
            has_in_press = bool(re.search(r'\bin\s+press\b', parens_content, re.IGNORECASE))
            
            if not years and not has_nd and not has_in_press:
                continue
            
            # If we have n.d. or in press, use that as the year
            if has_nd:
                years = ['n.d.']
            elif has_in_press:
                years = ['in press']

            # Skip if parentheses contain a FULL citation (author, year) not just year
            # Example: "Indigenous cultures (First Nations Pedagogy Online, 2019)"
            # This is a parenthetical citation, not a narrative one
            # Check if parens_content has both letters AND year (indicating full citation)
            # BUT: Don't skip if it's just "n.d." or "in press" (these are valid narrative citations)
            parens_has_letters = bool(re.search(r'[A-Za-z]{3,}', parens_content))
            is_just_nd_or_in_press = has_nd or has_in_press
            if parens_has_letters and years and not is_just_nd_or_in_press:
                # This looks like a full parenthetical citation, skip it
                continue


            # Clean Author Name
            # Remove possessives ('s, ’s, or just ’)
            author = re.sub(r'[\'’]s?\b', '', author_raw).strip()
            
            # Trim away common lowercase words from the beginning
            # This handles cases like "An interesting outcome of Rumbaut's" -> "Rumbaut's"
            # Split by spaces and find the last sequence of capitalized words
            words = author.split()
            if len(words) > 1:
                # Find the rightmost capitalized word that's not a stopword
                # Work backwards to find where the actual author name starts
                stopwords = {'of', 'the', 'and', 'for', 'a', 'an', 'in', 'on', 'at', 'to', 'from', 'with', 'by', 'as', 'et', 'al', 'al.'}
                cap_start_idx = None
                
                for i in range(len(words) - 1, -1, -1):
                    word = words[i]
                    # If this is a capitalized word that's not a stopword, this could be the start
                    if word and word[0].isupper() and word.lower() not in stopwords:
                        cap_start_idx = i
                    # If we hit a lowercase word that's not a stopword, stop searching
                    elif word and word[0].islower() and word.lower() not in stopwords:
                        break
                
                # Trim to only the capitalized sequence
                if cap_start_idx is not None and cap_start_idx > 0:
                    author = ' '.join(words[cap_start_idx:])
            
            # Remove common prefixes
            author = re.sub(r'^(According to|As cited by|As stated by|See also)\s+', '', author, flags=re.IGNORECASE).strip()

            # Skip exclusions
            # Expanded exclusion list: common non-citation patterns
            if re.match(r'^(Table|Figure|Fig|Eds?|Vol|Suppl|Appendix|Chapter|Section|Part|between|except|Ruth|Johnny|Z\.?|UK)\b', author, re.IGNORECASE):
                continue
            
            # Exclude Tool/Instrument/Scale/Measure names (e.g., "APA Psychotherapy Competency Tool")
            # These appear as "ToolName (author, year)" but are NOT citations to be matched
            # Pattern: if author contains Tool, Scale, Measure, Assessment, Instrument, Inventory, Index, etc.
            if re.search(r'\b(Tool|Scale|Measure|Assessment|Instrument|Inventory|Index|Test|Battery|Questionnaire|Survey|Protocol)\b', author, re.IGNORECASE):
                continue
            
            # Generic Abbreviation Exclusion: All Caps (ICU, PMHNP) or Plural Caps (NPs, APPNs)
            # Heuristic: 2+ uppercase letters, optional 's' at end.
            # Use search at end of string ($) to handle "The PMHNP" or "Many NPs"
            if re.search(r'\b[A-Z]{2,}s?$', author):
                continue
            
            if author.lower() in ['see', 'date', 'january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']:
                continue

            # Word Count Limit & Title Case Heuristic
            words = author.split()
            word_count = len(words)
            
            # 1. Hard fail on very long sentences (typical author: 1-5 words)
            # "Smith" (1), "Smith Jones" (2), "Smith, Jones & Brown" (3-4), "Brown Family Research Group" (4-5)
            if word_count > 6:
                continue
                
            # 2. Strict Capitalization for Multi-Word Authors (>2 words)
            # Ignore common stopwords (of, and, the, etc) to allow "University of Ottawa"
            # Also ignore "et" and "al" for et al. citations (e.g., "Vives et al.")
            if word_count > 2:
                stopwords = {'of', 'and', 'the', 'for', 'a', 'an', 'in', 'on', 'at', 'to', 'from', 'with', 'by', 'as', 'et', 'al', 'al.'}
                
                # Filter words that are NOT stopwords
                meaningful_words = [w for w in words if w.lower() not in stopwords]
                
                if not meaningful_words:
                    continue
                    
                capitalized_meaningful = sum(1 for w in meaningful_words if w[0].isupper())
                ratio = capitalized_meaningful / len(meaningful_words)
                
                # Require > 70% capitalized meaningful words
                # "University of Ottawa" -> ["University", "Ottawa"] -> 2/2 -> 1.0 (Pass)
                # "Toronto as a city counsellor" -> ["Toronto", "city", "counsellor"] -> 1/3 -> 0.33 (Fail)
                if ratio < 0.7:
                    continue

            # Validation for Narrative
            warnings = []
            if '&' in author:
                warnings.append("Format Error: Use 'and' in narrative citations, not '&'")
            
            # Create results for EACH year found (Multi-year narrative: Shapiro (2001, 2012))
            # Skip if already extracted from parenthetical citations
            for year in years:
                # Check if this pair was already extracted from parenthetical citations
                # Normalize: strip trailing periods/punctuation for consistent matching
                author_normalized = re.sub(r'[.,;: ]+$', '', author).strip()
                if (author_normalized.lower(), year) in parenthetical_pairs:
                    continue  # Skip - already have this citation from parenthetical
                
                results.append({
                    'author': author,
                    'year': year,
                    'type': 'narrative',
                    'warnings': warnings,
                    'raw': match.group(0)
                })
        return results
    
    def parse_reference(self, text: str) -> Optional[Dict]:
        """
        Parse APA-style reference like:
        Smith, J. (2020). Title of work. Publisher.
        Also supports: (2020, May 15), (n.d.), (in press)
        """
        match = self.reference_pattern.match(text)
        if not match:
            return None
        
        author_part = match.group(1).strip()
        date_part = match.group(2).strip()
        year = ""
        
        # Extract simple year for matching logic
        # 1. Try finding 4-digit year with optional letter
        year_match = re.search(r'\b((?:19|20)\d{2}[a-z]?)\b', date_part)
        if year_match:
            year = year_match.group(1)
        # 2. Handle n.d. (with optional suffix)
        elif 'n.d' in date_part.lower():
            # Try to extract "n.d.-a" or similar
            nd_match = re.search(r'\bn\.?d\.?(?:-[a-z])?\b', date_part, re.IGNORECASE)
            if nd_match:
                year = nd_match.group(0)
            else:
                year = "n.d."
        # 3. Handle in press
        elif 'in press' in date_part.lower():
            year = "in press"
        else:
            # Fallback: Just valid if something is there, but might not be usable for year matching
            # Return raw date_part as year for display at least
             year = date_part
        
        # Remove (Eds.) or (Ed.) from author part
        author_part = re.sub(r'\s*\(\s*Eds?\.?\s*\)', '', author_part, flags=re.IGNORECASE).strip()
        
        # Remove trailing comma
        author_part = re.sub(r',\s*$', '', author_part)
        
        # Extract abbreviations if present (in square brackets OR parentheses)
        # e.g. "American Nurses Association [ANA]" or "Centers... (CDC)"
        abbreviations = re.findall(r'[\[\(]([A-Z]{2,})[\]\)]', author_part)
        
        # Clean author part of abbreviations (remove both [ANA] and (CDC))
        author_clean = re.sub(r'\s*[\[\(][^\]\)]+[\]\)]\s*', ' ', author_part).strip()
        
        # Infer abbreviations from author names if not explicitly provided
        if not abbreviations and ',' not in author_clean and len(author_clean.split()) >= 2:
            words = author_clean.split()
            # Only infer if looks like an Organization (all caps start) and not just a person name
            # Heuristic: If it has "Association", "Organization", "Society", "Group" etc. OR strict Capitalization?
            # Existing logic was simple: First letter of each word.
            if all(w[0].isupper() for w in words if w):
                inferred_abbr = ''.join(w[0] for w in words if w[0].isupper())
                if len(inferred_abbr) >= 3: # Increase threshold to avoid initials like AB
                    abbreviations.append(inferred_abbr)
        
        # Format display author logic
        if '&' in author_clean:
            authors = author_clean.split('&')
            if len(authors) == 2:
                first_author = authors[0].strip().split(',')[0].strip()
                second_author = authors[1].strip().split(',')[0].strip()
                display_author = f"{first_author} & {second_author}"
            else:
                # 3 or more, APA 7: First et al. if 21+? citations usually list up to 20.
                # But for *in-text* citation matching, we need the first surname.
                first_author = authors[0].strip().split(',')[0].strip()
                display_author = first_author # Usually used for matching
        elif ',' in author_clean:
            parts = author_clean.split(',')
            # Standard Last, F. M.
            display_author = parts[0].strip()
        else:
            display_author = author_clean.strip()
        
        return {
            'author': display_author,
            'year': year,
            'full_author': author_part,
            'abbreviations': abbreviations,
            'raw_text': text
        }

class VancouverCitationParser(CitationParser):
    """
    Parser for Vancouver style citations.
    
    Vancouver style has two variants:
    1. Numeric: [1], [2], [1-3], [1,3,5] - NOT implemented here (author-year only)
    2. Author-Year: (Smith 2020) - no comma between author and year
    """
    
    def __init__(self):
        super().__init__("Vancouver")
        self.year_pattern = re.compile(r'\b((?:19|20)\d{2})[a-z]?\b')
        # Vancouver references often use: Smith J. Title. Journal. 2020;10(2):123-45.
        self.reference_pattern = re.compile(r'^([A-Z][^.]+?)\.\s*([^.]+?)\.\s*.*?(\d{4})')
    
    def parse_citation(self, cite_text: str) -> List[Dict]:
        """
        Parse Vancouver author-year citations like (Smith 2020) or (Jones and Brown 2019).
        Note: Vancouver uses 'and' instead of '&', and no comma before year.
        
        Args:
            cite_text: Text segment containing citations (e.g. a full paragraph)
            
        Returns:
            List of dicts with author, year, type, etc.
        """
        results = []
        cite_text_clean = cite_text.strip()
        
        # Regex to find content inside parentheses: (Smith 2020)
        parenthetical_pattern = re.compile(r'\(([^()]+)\)')
        
        for match in parenthetical_pattern.finditer(cite_text_clean):
            content = match.group(1).strip()
            
            # Remove square brackets and common prefixes from the CONTENT
            content = re.sub(r'\[[^\]]+\]', '', content).strip()
            content = re.sub(r'^(see|cf\.?|e\.g\.?,?|i\.e\.?,?)\s+', '', content, flags=re.IGNORECASE).strip()
            
            # Skip page references
            if re.search(r'\bp\.?\s*\d+', content, re.IGNORECASE):
                continue
            
            # Skip month names
            if re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)', 
                         content, re.IGNORECASE):
                continue
            
            # Find years in the content
            years = self.year_pattern.findall(content)
            if not years:
                continue
            
            # Extract author part (Vancouver has no comma before year)
            # Pattern: "Smith 2020" or "Smith and Jones 2020"
            # Remove the years we found
            author_part = content
            for year in years:
                author_part = re.sub(fr'\b{year}[a-z]?\b', '', author_part)
                
            author_part = author_part.strip()
            # Remove trailing/leading punctuation/digits
            author_part = re.sub(r'^\d+\s*', '', author_part).strip()
            author_part = re.sub(r'[.,;:]+$', '', author_part).strip()
            
            # Hard Word Count Limit to prevent matching long non-citation text
            # e.g. (Although a wide variety of terms ... 2017)
            if len(author_part.split()) > 6:
                continue
                
            if not author_part or len(author_part) < 2:
                continue
            
            # Add results
            for year in years:
                 results.append({
                     'author': author_part,
                     'year': year,
                     'type': 'parenthetical',
                     'warnings': [],
                     'raw': f"({content})"
                 })
                 
        return results
    
    def parse_reference(self, text: str) -> Optional[Dict]:
        """
        Parse Vancouver-style reference like:
        Smith J, Jones M. Title of article. Journal Name. 2020;10(2):123-45.
        
        Args:
            text: Reference line text
            
        Returns:
            Dict with author, year, and other metadata
        """
        match = self.reference_pattern.match(text)
        if not match:
            return None
        
        author_part = match.group(1).strip()
        year = match.group(3).strip()
        
        # Vancouver uses abbreviated names: Smith J, Jones M
        # Extract first author for display
        if ',' in author_part:
            # Multiple authors
            authors = author_part.split(',')
            first_author = authors[0].strip()
            
            if len(authors) >= 3:
                display_author = first_author + " et al"
            elif len(authors) == 2:
                second_author = authors[1].strip()
                display_author = f"{first_author} and {second_author}"
            else:
                display_author = first_author
        else:
            display_author = author_part.strip()
        
        return {
            'author': display_author,
            'year': year,
            'full_author': author_part,
            'abbreviations': [],
            'raw_text': text
        }


class ChicagoCitationParser(CitationParser):
    """
    Parser for Chicago Manual of Style (Author-Year) citations.
    Similar to Vancouver but uses 'and' and may have comma before year.
    """
    
    def __init__(self):
        super().__init__("Chicago")
        self.year_pattern = re.compile(r'\b((?:19|20)\d{2})[a-z]?\b')
        self.reference_pattern = re.compile(r'^([A-Z][^.]+?)\.\s*(\d{4})\.\s*')
    
    def parse_citation(self, cite_text: str) -> List[Dict]:
        """
        Parse Chicago author-year citations like (Smith 2020) or (Smith and Jones 2020).
        
        Args:
            cite_text: Citation text from parentheses
            
        Returns:
            List of dicts with author, year, etc.
        """
        # Chicago is very similar to Vancouver for author-year
        # Use same logic as Vancouver
        vancouver_parser = VancouverCitationParser()
        return vancouver_parser.parse_citation(cite_text)
    
    def parse_reference(self, text: str) -> Optional[Dict]:
        """
        Parse Chicago-style reference like:
        Smith, John. 2020. Title of Book. Publisher.
        
        Args:
            text: Reference line text
            
        Returns:
            Dict with author, year, and other metadata
        """
        match = self.reference_pattern.match(text)
        if not match:
            return None
        
        author_part = match.group(1).strip()
        year = match.group(2).strip()
        
        # Chicago uses full names: Smith, John
        if ',' in author_part:
            parts = author_part.split(',')
            if len(parts) >= 3:
                display_author = parts[0].strip() + " et al."
            else:
                display_author = parts[0].strip()
        else:
            display_author = author_part.strip()
        
        return {
            'author': display_author,
            'year': year,
            'full_author': author_part,
            'abbreviations': [],
            'raw_text': text
        }


def get_parser(style: str) -> CitationParser:
    """
    Get the appropriate citation parser for a given style.
    
    Args:
        style: Citation style name ('apa', 'vancouver', 'chicago')
        
    Returns:
        CitationParser instance
        
    Raises:
        ValueError: If style is not supported
    """
    style = style.lower().strip()
    
    parsers = {
        'apa': APACitationParser,
        'vancouver': VancouverCitationParser,
        'chicago': ChicagoCitationParser
    }
    
    if style not in parsers:
        raise ValueError(f"Unsupported citation style: {style}. Supported styles: {', '.join(parsers.keys())}")
    
    return parsers[style]()


def auto_detect_style(text_sample: str) -> str:
    """
    Attempt to auto-detect citation style from a text sample.
    
    Args:
        text_sample: Sample text containing citations
        
    Returns:
        Detected style name ('apa', 'vancouver', or 'chicago')
        Defaults to 'apa' if uncertain
    """
    # Look for characteristic patterns
    
    # APA uses comma before year: (Smith, 2020)
    apa_pattern = r'\([A-Z][a-z]+,\s*\d{4}\)'
    
    # Vancouver/Chicago use no comma or space: (Smith 2020)
    vancouver_pattern = r'\([A-Z][a-z]+\s+\d{4}\)'
    
    apa_matches = len(re.findall(apa_pattern, text_sample))
    vancouver_matches = len(re.findall(vancouver_pattern, text_sample))
    
    if apa_matches > vancouver_matches:
        return 'apa'
    elif vancouver_matches > 0:
        # Default to Vancouver if we see the pattern
        # (Chicago and Vancouver are very similar in author-year format)
        return 'vancouver'
    else:
        # Default to APA
        return 'apa'
