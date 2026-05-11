"""
Hierarchy Manager module.
Implements Stage 2 (Validate) and Stage 3 (Standardize) of the STM styling workflow.
"""

import logging
from typing import List, Dict, Any
from .rules_loader import get_rules_loader

logger = logging.getLogger(__name__)

class HierarchyManager:
    def __init__(self):
        self.rules_loader = get_rules_loader()
        self.config = self.rules_loader.get_heading_hierarchy()
        
    def refine_annotations(self, annotations: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Refine annotations based on STM hierarchy rules.
        - Normalizes synonyms
        - Enforces mandatory levels
        - Validates and fixes hierarchy (no skipping levels)
        """
        if not self.config:
            logger.warning("No heading hierarchy configuration found")
            return annotations
            
        synonyms = self.config.get("synonym_normalization", {})
        mandatory_h1 = set(self.config.get("mandatory_h1_sections", []))
        constraints = self.config.get("constraints", {})
        
        current_level = 0  # 0 indicates no heading seen yet
        
        # Map style names to levels
        style_levels = {
            "CT": 0,
            "Title": 0,
            "H1": 1,
            "Heading 1": 1,
            "H2": 2,
            "Heading 2": 2,
            "H3": 3,
            "Heading 3": 3,
            "H4": 4,
            "Heading 4": 4,
            "Normal": 99,
            "TXT": 99,
            "BODY_TEXT": 99
        }
        
        # Reverse map for level -> style
        level_styles = {
            1: "H1",
            2: "H2",
            3: "H3",
            4: "H4"
        }

        refined_annotations = []
        
        for idx, item in enumerate(annotations):
            para = item["para"]
            tag = item["tag"]
            style = item["style"]
            text = para.text.strip()
            
            # Skip empty or non-heading items mostly, but we need to track context
            # Actually, we process all, modify if needed.
            
            # 1. Synonym Normalization
            if text in synonyms:
                new_text = synonyms[text]
                logger.info(f"Normalizing '{text}' to '{new_text}'")
                if para.runs:
                    # Only safe to rewrite when the full text lives in one run;
                    # multi-run paragraphs may have mixed formatting — leave them.
                    if len(para.runs) == 1 or para.runs[0].text.strip() == text:
                        para.runs[0].text = new_text
                    else:
                        logger.warning(f"Skipping synonym rewrite for multi-run paragraph '{text}' to avoid formatting loss")
                else:
                    para.add_run(new_text)
                text = new_text
            
            # 2. Mandatory H1 Enforcement
            if text in mandatory_h1:
                if style != "H1":
                    logger.info(f"Enforcing H1 for mandatory section '{text}'")
                    style = "H1"
                    tag = "H1"
            
            # 3. Hierarchy Validation (Auto-fix)
            # Check if this item determines a level
            level = style_levels.get(style, 99)
            
            if level <= 4: # It is a heading
                # Check constraints
                if constraints.get("require_h1_first", False):
                    # Only enforce if skipping level 1 AND it's not a root element like Title (level 0)
                    if level > 1 and current_level == 0:
                        logger.warning(f"Heading '{text}' (H{level}) appears before first H1. Promoting to H1.")
                        level = 1
                        style = "H1"
                        
                if constraints.get("no_skipping_levels", False):
                    # valid: current=1, next=2. invalid: current=1, next=3
                    if level > current_level + 1:
                        # Auto-fix: Reduce level to be sequential
                        new_level = current_level + 1
                        # But ensure we don't go deeper than max_depth or stay 0 if current is 0?
                        # If current is 0 (Title/Start), next can be 1.
                        # If current is 1 (H1), next can be 2.
                        # If current is 1, next cannot be 3.
                        
                        # However, sometimes we jump back up (H3 -> H1). That is allowed.
                        # Skipping is only forbidden downwards (H1 -> H3).
                        
                        logger.warning(f"Hierarchy violation: H{current_level} -> H{level}. Auto-fixing to H{new_level}.")
                        level = new_level
                        style = level_styles.get(level, style)
                
                # Update current level context
                current_level = level
                
            item["style"] = style
            item["tag"] = tag
            refined_annotations.append(item)
            
        return refined_annotations

def enforce_hierarchy(annotations: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    manager = HierarchyManager()
    return manager.refine_annotations(annotations)
