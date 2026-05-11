"""
Rules configuration loader.
Loads rules from YAML and validates them against the document template.
"""

import re
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

try:
    import yaml
except ImportError:
    yaml = None

from .logger_config import get_logger

logger = get_logger(__name__)


class RuleValidator:
    """Validates rules for correctness and conflicts"""
    
    @staticmethod
    def validate_regex_pattern(pattern: str, max_length: int = 100) -> Tuple[bool, str]:
        """
        Validate regex pattern is correct and safe.
        
        Args:
            pattern: Regex pattern to validate
            max_length: Maximum pattern length to prevent DOS
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        if len(pattern) > max_length:
            return False, f"Pattern too long ({len(pattern)} > {max_length})"
        
        try:
            re.compile(pattern)
            return True, ""
        except re.error as e:
            return False, f"Invalid regex: {str(e)}"
    
    @staticmethod
    def validate_rules(rules: Dict[str, Any]) -> List[str]:
        """
        Validate loaded rules for consistency.
        
        Args:
            rules: Rules dictionary from YAML
        
        Returns:
            List of error messages (empty if valid)
        """
        errors = []
        
        # Get validation settings
        max_length = 100
        if "validation" in rules and "max_pattern_length" in rules["validation"]:
            max_length = rules["validation"]["max_pattern_length"]

        # Validate paragraph rules
        if "paragraphs" in rules:
            seen_priorities = set()
            for idx, rule in enumerate(rules["paragraphs"]):
                # Check required fields
                if "pattern" not in rule:
                    errors.append(f"Paragraph rule {idx}: missing 'pattern'")
                elif not isinstance(rule["pattern"], str):
                    errors.append(f"Paragraph rule {idx}: pattern must be string")
                else:
                    valid, msg = RuleValidator.validate_regex_pattern(rule["pattern"], max_length)
                    if not valid:
                        errors.append(f"Paragraph rule {idx}: {msg}")
                
                if "tag" not in rule:
                    errors.append(f"Paragraph rule {idx}: missing 'tag'")
                if "style" not in rule:
                    errors.append(f"Paragraph rule {idx}: missing 'style'")
                
                # Check priority
                if "priority" in rule:
                    priority = rule["priority"]
                    if priority in seen_priorities:
                        logger.warning(f"Duplicate priority {priority} in rule {idx}")
                    seen_priorities.add(priority)
        
        # Validate block definitions
        if "blocks" in rules:
            if "start_markers" in rules["blocks"]:
                markers = rules["blocks"]["start_markers"]
                if not isinstance(markers, dict):
                    errors.append("blocks.start_markers must be a dict")
            
            if "end_patterns" in rules["blocks"]:
                for idx, pattern_def in enumerate(rules["blocks"]["end_patterns"]):
                    if "pattern" not in pattern_def:
                        errors.append(f"blocks.end_patterns[{idx}]: missing 'pattern'")
                    else:
                        valid, msg = RuleValidator.validate_regex_pattern(pattern_def["pattern"])
                        if not valid:
                            errors.append(f"blocks.end_patterns[{idx}]: {msg}")
        
        # Validate list patterns
        if "lists" in rules:
            lists_cfg = rules["lists"]
            if "bullet_pattern" in lists_cfg:
                valid, msg = RuleValidator.validate_regex_pattern(lists_cfg["bullet_pattern"])
                if not valid:
                    errors.append(f"lists.bullet_pattern: {msg}")
            if "number_pattern" in lists_cfg:
                valid, msg = RuleValidator.validate_regex_pattern(lists_cfg["number_pattern"])
                if not valid:
                    errors.append(f"lists.number_pattern: {msg}")
            if "roman_pattern" in lists_cfg:
                valid, msg = RuleValidator.validate_regex_pattern(lists_cfg["roman_pattern"])
                if not valid:
                    errors.append(f"lists.roman_pattern: {msg}")

        # Validate box patterns
        if "boxes" in rules:
            boxes_cfg = rules["boxes"]
            for key in ("open_patterns", "close_patterns"):
                for idx, pattern in enumerate(boxes_cfg.get(key, [])):
                    valid, msg = RuleValidator.validate_regex_pattern(pattern)
                    if not valid:
                        errors.append(f"boxes.{key}[{idx}]: {msg}")
        
        return errors


class RulesLoader:
    """Load and manage rules from YAML configuration"""
    
    def __init__(self, rules_path: Optional[str] = None):
        """
        Initialize rules loader.
        
        Args:
            rules_path: Path to rules.yaml file
        """
        if rules_path is None:
            rules_path = Path(__file__).parent / "rules.yaml"
        
        self.rules_path = Path(rules_path)
        self.rules: Dict[str, Any] = {}
        self.loaded = False
    
    def load(self, validate: bool = True) -> bool:
        """
        Load rules from YAML file.
        
        Args:
            validate: Whether to validate rules after loading
        
        Returns:
            True if loaded successfully
        """
        if not self.rules_path.exists():
            logger.error(f"Rules file not found: {self.rules_path}")
            return False
        
        if yaml is None:
            logger.error("PyYAML not installed. Install with: pip install pyyaml")
            return False
        
        try:
            with open(self.rules_path, 'r', encoding='utf-8') as f:
                self.rules = yaml.safe_load(f)
            
            logger.info(f"Loaded rules from {self.rules_path}")
            
            if validate:
                errors = RuleValidator.validate_rules(self.rules)
                if errors:
                    for error in errors:
                        logger.error(f"Rule validation error: {error}")
                    return False
                logger.info("Rules validation passed")
            
            self.loaded = True
            return True
        
        except Exception as e:
            logger.error(f"Failed to load rules: {e}", exc_info=True)
            return False
    
    def get_paragraphs(self) -> List[Dict[str, Any]]:
        """Get paragraph rules sorted by priority"""
        if not self.loaded:
            return []
        
        paragraphs = self.rules.get("paragraphs", [])
        return sorted(paragraphs, key=lambda x: x.get("priority", 999))
    
    def get_block_start_markers(self) -> Dict[str, str]:
        """Get block start markers"""
        if not self.loaded:
            return {}
        
        return self.rules.get("blocks", {}).get("start_markers", {})
    
    def get_block_end_patterns(self) -> List[Dict[str, str]]:
        """Get block end patterns"""
        if not self.loaded:
            return []
        
        return self.rules.get("blocks", {}).get("end_patterns", [])
    
    def get_list_patterns(self) -> Dict[str, str]:
        """Get list detection patterns"""
        if not self.loaded:
            return {}
        
        return self.rules.get("lists", {})
    
    def get_table_config(self) -> Dict[str, Any]:
        """Get table detection configuration"""
        if not self.loaded:
            return {}
        
        return self.rules.get("tables", {})
    
    def get_heading_hierarchy(self) -> Dict[str, Any]:
        """Get heading hierarchy rules"""
        if not self.loaded:
            return {}
        
        return self.rules.get("heading_hierarchy", {})
    
    def get_state_machine_config(self) -> Dict[str, Any]:
        """Get state machine configuration"""
        if not self.loaded:
            return {}
        
        return self.rules.get("state_machine", {})

    def get_cross_references(self) -> Dict[str, Any]:
        """Get cross-reference configuration"""
        if not self.loaded:
            return {}
        
        return self.rules.get("cross_references", {})

    def get_box_config(self) -> Dict[str, Any]:
        """Get box handling configuration"""
        if not self.loaded:
            return {}

        return self.rules.get("boxes", {})

    def get_keyterm_config(self) -> Dict[str, Any]:
        """Get key-term handling configuration"""
        if not self.loaded:
            return {}

        return self.rules.get("keyterms", {})

    def get_normalization_config(self) -> Dict[str, Any]:
        """Get style/tag normalization configuration"""
        if not self.loaded:
            return {}

        return self.rules.get("normalization", {})


# Global rules instance
_rules_loader: Optional[RulesLoader] = None


def get_rules_loader(rules_path: Optional[str] = None) -> RulesLoader:
    """Get or create global rules loader instance"""
    global _rules_loader
    
    requested_path = Path(rules_path) if rules_path is not None else None

    if (
        _rules_loader is None
        or (
            requested_path is not None
            and _rules_loader.rules_path != requested_path
        )
    ):
        _rules_loader = RulesLoader(rules_path)
        if not _rules_loader.load():
            raise RuntimeError(
                f"Failed to load or validate rules from {_rules_loader.rules_path}. "
                "Check logs for validation errors."
            )

    return _rules_loader


def reload_rules(rules_path: Optional[str] = None) -> bool:
    """Reload rules from file"""
    global _rules_loader
    
    _rules_loader = RulesLoader(rules_path)
    return _rules_loader.load()


if __name__ == "__main__":
    # Test loader
    loader = get_rules_loader()
    
    if loader.loaded:
        print("✓ Rules loaded successfully")
        print(f"\nParagraph rules: {len(loader.get_paragraphs())}")
        print(f"Block markers: {loader.get_block_start_markers()}")
        print(f"Heading hierarchy: {loader.get_heading_hierarchy()}")
    else:
        print("✗ Failed to load rules")
